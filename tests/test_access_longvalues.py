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
TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


def _lv_ref(table: Table, rid: RowId, column: str) -> LongValueRef:
    raw = table.fetch_row(rid.page, rid.slot)
    assert raw is not None
    value = split_row(table.definition, raw).values[table.definition.column(column).number]
    assert value is not None
    return decode_long_value_ref(value)


def _memo_table(db: AccessDatabase, name: str) -> Table:
    from pyopenvba.access import ColumnSpec, IndexSpec

    return db.create_table(name, [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("M", "Memo", compressed=False)], [IndexSpec("PK", ("Id",), primary=True)])


def _placement(table: Table) -> dict[int, tuple[int, int, int]]:
    d = table.definition
    out: dict[int, tuple[int, int, int]] = {}
    for _page, _slot, raw in table.raw_rows():
        parts = split_row(d, raw)
        lv = parts.values[d.column("M").number]
        assert isinstance(lv, bytes)
        ref = decode_long_value_ref(lv)
        out[int(table.decode(parts)["Id"])] = (ref.page, ref.row, ref.length)  # pyright: ignore[reportArgumentType]
    return out


def _lv_maps(table: Table) -> tuple[list[int], list[int], dict[int, int]]:
    db = table.database
    owned, fsp = table.definition.column_usage_maps[table.definition.column("M").number]
    pages = list(read_usage_map_ref(db.store, owned).pages())
    return pages, list(read_usage_map_ref(db.store, fsp).pages()), {p: int.from_bytes(db.store.read(p)[2:4], "little") for p in pages}


def test_single_row_values_follow_the_engines_cursor_then_the_free_space_map() -> None:
    """DAO in one session: two 3000-byte values took pages A and B; a
    900-byte value went to B (the page last written), the next one, B
    now full, to A; both pages ended at 178 free and unlisted."""
    db = AccessDatabase(TEMPLATE)
    table = _memo_table(db, "L3")
    for text in ("a" * 1500, "b" * 1500, "c" * 450, "d" * 450):
        table.insert_row({"M": text})
    where = _placement(table)
    a, b = where[1][0], where[2][0]
    assert a < b and where[3] == (b, 1, 900) and where[4] == (a, 1, 900)
    pages, listed, free = _lv_maps(table)
    assert pages == [a, b] and listed == [] and free == {a: 178, b: 178}


def test_lval_pages_stay_listed_above_256_bytes_free() -> None:
    db = AccessDatabase(TEMPLATE)
    for free, expect_listed in ((256, False), (258, True)):
        table = _memo_table(db, f"F{free}")
        table.insert_row({"M": "a" * 650})
        table.insert_row({"M": "b" * ((2778 - free) // 2)})
        pages, listed, actual = _lv_maps(table)
        assert actual == {pages[0]: free} and bool(listed) is expect_listed


def test_updates_store_the_new_value_before_freeing_and_deletes_relist() -> None:
    """DAO: six 1300-byte values on two pages (176 free each, unlisted);
    deleting one re-lists its page at 1476; a 500-byte insert fills the
    hole; updating a value on the other page puts the new value on the
    listed page and frees the old, which lists its page; a 2800-byte
    value fits nowhere and takes a fresh page.  Each step was its own
    session, so the cursor is cleared between them."""
    db = AccessDatabase(TEMPLATE)
    table = _memo_table(db, "L")
    for i in range(6):
        table.insert_row({"M": chr(97 + i) * 650})
    ids = {row["Id"]: rid for rid, row in table.rows_with_ids()}
    p1, p2 = _lv_maps(table)[0]
    assert _lv_maps(table)[1] == [] and _lv_maps(table)[2] == {p1: 176, p2: 176}

    def new_session() -> Table:
        return AccessDatabase(db.to_bytes()).table("L")

    table = new_session(); table.delete_row(ids[2])
    assert _lv_maps(table)[1] == [p1] and _lv_maps(table)[2][p1] == 1476
    db = table.database
    table = new_session(); table.insert_row({"M": "n" * 250})
    assert _placement(table)[7] == (p1, 3, 500) and _lv_maps(table)[2][p1] == 974
    db = table.database
    table = new_session(); table.update_row({row["Id"]: rid for rid, row in table.rows_with_ids()}[4], {"M": "u" * 400})
    assert _placement(table)[4] == (p1, 4, 800)
    assert _lv_maps(table)[1] == [p2] and _lv_maps(table)[2] == {p1: 172, p2: 1476}
    db = table.database
    table = new_session(); table.insert_row({"M": "p" * 1400})
    pages, listed, free = _lv_maps(table)
    assert len(pages) == 3 and _placement(table)[8] == (pages[2], 0, 2800) and listed == [p2, pages[2]] and free[pages[2]] == 1280
    check_indexes(table)


def test_an_update_never_lands_in_the_hole_it_opens() -> None:
    """DAO: values 1-3 on page A (176 free), 4-5 on page B; updating value
    1 to 1000 bytes put it on B, then freed A, which lists again."""
    db = AccessDatabase(TEMPLATE)
    table = _memo_table(db, "L2")
    for i in range(5):
        table.insert_row({"M": chr(97 + i) * 650})
    a, b = _lv_maps(table)[0]
    table = AccessDatabase(db.to_bytes()).table("L2")
    table.update_row({row["Id"]: rid for rid, row in table.rows_with_ids()}[1], {"M": "u" * 500})
    assert _placement(table)[1] == (b, 2, 1000)
    assert _lv_maps(table)[1] == [a, b] and _lv_maps(table)[2] == {a: 1476, b: 476}


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


def test_usage_maps_grow_past_512_pages(tmp_path: Path) -> None:
    """Inline usage maps cover 512 pages at first; the engine enlarges the
    bitmap in 8-byte steps and re-bases an empty map to its first page.
    Measured: 573 pages give the global map a 72-byte bitmap, a table
    holding only page 542 a start page of 536."""
    from pyopenvba.access import ColumnSpec, IndexSpec
    from pyopenvba.access._pages import row_bytes

    template = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
    db = AccessDatabase(template)
    table = db.create_table("Memos", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("T", "Text", size=50), ColumnSpec("M", "Memo")], [IndexSpec("PK", ("Id",), primary=True)])
    for i in range(450):
        table.insert_row({"T": f"m{i}", "M": "a" * 1600})
    assert db.store.page_count > 512
    global_row = row_bytes(db.store.read(GLOBAL_USAGE_MAP_PAGE), GLOBAL_USAGE_MAP_ROW)
    assert global_row is not None and len(global_row) == 5 + 8 * (-(-db.store.page_count // 64))
    free = read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    assert not [p for p in free.pages() if p < db.store.page_count]
    d = table.definition
    owned = read_usage_map_ref(db.store, d.column_usage_maps[d.column("M").number][0])
    assert max(owned.pages()) > 512 and len(owned.pages()) == 450
    db.save(tmp_path / "grown.accdb")
    again = AccessDatabase(tmp_path / "grown.accdb")
    rows = list(again.table("Memos").rows())
    assert len(rows) == 450 and all(r["M"] == "a" * 1600 for r in rows)
    check_indexes(again.table("Memos"))


def test_unique_index_rejects_a_duplicate_key() -> None:
    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    with pytest.raises(AccessError):
        table.insert_row({"ID": 1, "Field1": "clashes with the primary key"})
    assert table.row_count == 3


def test_a_row_coming_home_retires_the_copy_page_it_empties() -> None:
    """DAO: a row grown past its page's room moved alone to a fresh page;
    shrunk again it came home and that page was retired (type 0x09,
    released, out of both maps), while a delete would have left it."""
    from pyopenvba.access import ColumnSpec, IndexSpec
    from pyopenvba.access._pages import PAGE_RETIRED, read_usage_map

    db = AccessDatabase(TEMPLATE)
    table = db.create_table("W", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("T", "Text", size=255, compressed=False)], [IndexSpec("PK", ("Id",), primary=True)])
    first = table.insert_row({"T": "a" * 30})
    for i in range(7):
        table.insert_row({"T": chr(98 + i) * 255})
    home = table.data_pages()
    assert len(home) == 1
    table.update_row(first, {"T": "z" * 255})
    moved = table._moved_to(first)  # pyright: ignore[reportPrivateUsage]
    assert moved is not None and moved[0] != home[0]
    copy_page = moved[0]
    table.update_row(first, {"T": "home again"})
    assert table._moved_to(first) is None  # pyright: ignore[reportPrivateUsage]
    assert db.store.read(copy_page)[0] == PAGE_RETIRED
    assert copy_page in set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    d = table.definition
    assert list(read_usage_map_ref(db.store, d.owned_pages_ref).pages()) == home
    assert list(read_usage_map_ref(db.store, d.free_space_pages_ref).pages()) == []
    check_indexes(table)


def test_a_freed_chains_pages_are_reused_only_when_older_than_the_session() -> None:
    """DAO: a 10 KB value's three chain pages, deleted in a later session,
    were taken again in order by the next 10 KB value; a chain created and
    freed within one session got fresh pages."""

    def chain_pages(table: Table) -> list[int]:
        d = table.definition
        return list(read_usage_map_ref(table.database.store, d.column_usage_maps[d.column("M").number][0]).pages())

    db = AccessDatabase(TEMPLATE)
    table = _memo_table(db, "K")
    first = table.insert_row({"M": "k" * 5000})
    pages = chain_pages(table)
    assert len(pages) == 3
    table.delete_row(first)
    table.insert_row({"M": "m" * 5000})
    assert set(chain_pages(table)).isdisjoint(pages)  # same session: fresh pages

    db = AccessDatabase(TEMPLATE)
    table = _memo_table(db, "K")
    table.insert_row({"M": "k" * 5000})
    pages = chain_pages(table)
    later = AccessDatabase(db.to_bytes()).table("K")
    later.delete_row(next(rid for rid, _ in later.rows_with_ids()))
    later.insert_row({"M": "m" * 5000})
    assert chain_pages(later) == pages  # a later session: the same pages, in order


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
