"""Writing long values and overflow rows.

The thresholds and layouts asserted here were measured on what the
engine wrote (docs/access_engine.md): 64 bytes inline, 3816 bytes as one
row on a shared LVAL page, longer values chained one 4072-byte chunk per
page; a row that no longer fits its page moves behind a pointer.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._lval import LVAL_INLINE_MAX, LVAL_SINGLE_MAX, memo_bytes
from pyopenvba.access._pages import (
    GLOBAL_USAGE_MAP_PAGE,
    GLOBAL_USAGE_MAP_ROW,
    ROW_DELETED,
    ROW_OVERFLOW,
    read_usage_map,
    read_usage_map_ref,
    row_slots,
)
from pyopenvba.access._rows import LongValueRef, decode_long_value_ref, split_row
from pyopenvba.access.database import RowId, Table
from pyopenvba.access_read import AccessError
from test_access_write import check_indexes

LARGE = Path(__file__).parent / "live_access_test" / "New Microsoft Access Database.accdb"


def _lv_ref(table: Table, rid: RowId, column: str) -> LongValueRef:
    raw = table.fetch_row(rid.page, rid.slot)
    assert raw is not None
    value = split_row(table.definition, raw).values[table.definition.column(column).number]
    assert value is not None
    return decode_long_value_ref(value)


def test_memo_bytes_follow_the_engine() -> None:
    assert memo_bytes("a") == "a".encode("utf-16-le")             # compression would not shorten it
    assert memo_bytes("abcdefghij") == b"\xff\xfeabcdefghij"      # inline and shorter compressed
    assert memo_bytes("a" * 33) == ("a" * 33).encode("utf-16-le")  # outside the row: never compressed
    assert LVAL_INLINE_MAX == 64 and LVAL_SINGLE_MAX == 3816


def test_ole_values_of_every_kind_round_trip(tmp_path: Path) -> None:
    """MSysAccessStorage carries an OLE column with real long-value usage
    maps; the fixture's user tables have no such column."""
    db = AccessDatabase(LARGE)
    table = db.table("MSysAccessStorage")
    cases = {
        "inline": bytes(range(60)),
        "single": bytes(range(256)) * 10,
        "chained": bytes(range(256)) * 40,
        "chained long": bytes(range(256)) * 90,
    }
    for label, value in cases.items():
        table.insert_row({"Name": f"lv {label}", "Lv": value})
    db.save(tmp_path / "lv.accdb")
    again = AccessDatabase(tmp_path / "lv.accdb")
    table = again.table("MSysAccessStorage")
    found = {str(r["Name"])[3:]: r["Lv"] for r in table.rows() if str(r["Name"]).startswith("lv ")}
    assert found == cases
    check_indexes(table)


def test_memo_text_round_trips(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    table = db.table("MSysQueries")
    texts = {1001: "short memo", 1002: "Привет 日本 inline", 1003: "m" * 1500, 1004: "long " * 3000}
    for object_id, text in texts.items():
        table.insert_row({"ObjectId": object_id, "Attribute": 1, "Expression": text})
    db.save(tmp_path / "memo.accdb")
    again = AccessDatabase(tmp_path / "memo.accdb")
    rows = {r["ObjectId"]: r["Expression"] for r in again.table("MSysQueries").rows()}
    for object_id, text in texts.items():
        assert rows[object_id] == text


def test_long_value_kinds_and_freeing(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    table = db.table("MSysAccessStorage")
    lv = table.definition.column("Lv")
    inline = table.insert_row({"Name": "k inline", "Lv": b"\x01" * 60})
    single = table.insert_row({"Name": "k single", "Lv": b"\x02" * 3000})
    chained = table.insert_row({"Name": "k chained", "Lv": b"\x03" * 9000})

    assert _lv_ref(table, inline, "Lv").kind == LongValueRef.KIND_INLINE
    assert _lv_ref(table, single, "Lv").kind == LongValueRef.KIND_SINGLE_PAGE
    chain = _lv_ref(table, chained, "Lv")
    assert chain.kind == LongValueRef.KIND_CHAINED and chain.length == 9000
    owned_ref, _free = table.definition.column_usage_maps[lv.number]
    assert chain.page in set(read_usage_map_ref(db.store, owned_ref).pages())
    # The first chunk page carries the stamp the definition carries.
    assert int.from_bytes(db.store.read(chain.page)[8:12], "little") == db.lval_stamp

    # Deleting the chained row gives its pages back to the global map.
    table.delete_row(chained)
    free_now = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert chain.page in free_now
    assert chain.page not in set(read_usage_map_ref(db.store, owned_ref).pages())
    # Replacing a single-page value tombstones its old row.
    old = _lv_ref(table, single, "Lv")
    table.update_row(single, {"Lv": b"\x04" * 100})
    assert row_slots(db.store.read(old.page))[old.row] & 0xC000 == 0xC000
    db.save(tmp_path / "kinds.accdb")
    again = AccessDatabase(tmp_path / "kinds.accdb")
    rows = {r["Name"]: r["Lv"] for r in again.table("MSysAccessStorage").rows()}
    assert rows["k inline"] == b"\x01" * 60
    assert rows["k single"] == b"\x04" * 100
    assert "k chained" not in rows


def test_unique_index_rejects_a_duplicate_key() -> None:
    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    with pytest.raises(AccessError):
        table.insert_row({"ID": 1, "Field1": "clashes with the primary key"})
    assert table.row_count == 3


def test_row_that_outgrows_its_page_moves_and_comes_back(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    home_page = table.data_pages()[0]
    victim = table.insert_row({"Field1": "small"})
    victim_id = 4
    while True:  # fill the page so the next growth cannot stay put
        if table.insert_row({"Field1": "filler row " * 5}).page != home_page:
            break
    assert not row_slots(db.store.read(victim.page))[victim.slot] & 0xC000
    table.update_row(victim, {"Field1": "grown " * 40})
    entry = row_slots(db.store.read(victim.page))[victim.slot]
    assert entry & ROW_OVERFLOW and not entry & ROW_DELETED
    moved = table._moved_to(victim)  # pyright: ignore[reportPrivateUsage]
    assert moved is not None and moved[0] != victim.page
    assert row_slots(db.store.read(moved[0]))[moved[1]] & 0xC000 == ROW_DELETED
    assert {r["ID"]: r["Field1"] for r in table.rows()}[victim_id] == "grown " * 40
    check_indexes(table)
    # Shrink it: it comes home and the copy dies.
    table.update_row(victim, {"Field1": "tiny"})
    assert not row_slots(db.store.read(victim.page))[victim.slot] & 0xC000
    assert row_slots(db.store.read(moved[0]))[moved[1]] & 0xC000 == 0xC000
    # Grow again, then delete: both slots end up dead.
    table.update_row(victim, {"Field1": "grown again " * 20})
    moved2 = table._moved_to(victim)  # pyright: ignore[reportPrivateUsage]
    assert moved2 is not None
    table.delete_row(victim)
    assert row_slots(db.store.read(victim.page))[victim.slot] & 0xC000 == 0xC000
    assert row_slots(db.store.read(moved2[0]))[moved2[1]] & 0xC000 == 0xC000
    db.save(tmp_path / "overflow.accdb")
    again = AccessDatabase(tmp_path / "overflow.accdb")
    table = again.table("Table1")
    assert victim_id not in {r["ID"] for r in table.rows()}
    check_indexes(table)
