"""Writing rows: data pages, B-trees, counters and allocation, checked the
way the engine would check them -- by reading everything back.

The page-level behaviour asserted here (row layout, tombstones, free
space, split policy) was measured on pages the ACE engine wrote; see
docs/access_engine.md.  The live gate hands the result to the engine.
"""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._datapage import DataPage
from pyopenvba.access._index import decode_key, leaf_entries, node_pages, parse_index_page
from pyopenvba.access._rows import encode_row, split_row
from pyopenvba.access.database import RowId, Table
from pyopenvba.access_read import AccessError

FIXTURES = Path(__file__).parent / "live_access_test"
TEMPLATES = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files"
AUTHORED = [
    FIXTURES / "New Microsoft Access Database.accdb",
    FIXTURES / "module_spanning_pages.accdb",
    FIXTURES / "two_modules_one_page.accdb",
    TEMPLATES / "blank_database.accdb",
    TEMPLATES / "blank_database_module.accdb",
]
LARGE = FIXTURES / "New Microsoft Access Database.accdb"


# --- row codec ----------------------------------------------------------------


@pytest.mark.parametrize("path", AUTHORED, ids=lambda p: p.name)
def test_every_stored_row_reassembles_byte_for_byte(path: Path) -> None:
    """split_row then encode_row must give back the exact bytes for every
    row of every table -- the strongest check the row writer can have
    without the engine present."""
    db = AccessDatabase(path)
    checked = 0
    for table in db.tables(include_system=True):
        for _page, _slot, raw in table.raw_rows():
            parts = split_row(table.definition, raw)
            if parts.column_count != len(table.columns):
                continue  # written before a column was added; not re-encodable as is
            booleans = {
                n for n, present in parts.present.items()
                if present and table.definition.column_by_number(n).type_code == 0x01
            }
            # The engine leaves whatever was in its buffer under a null
            # fixed-length column (one attachment row carries stale text
            # there); the writer zeroes it.  The null mask governs either way.
            expected = bytearray(raw)
            for column in table.columns:
                if column.is_fixed and column.type_code != 0x01 and not parts.present.get(column.number):
                    start = 2 + column.fixed_offset
                    expected[start : start + column.length] = bytes(column.length)
            assert encode_row(table.definition, parts.values, booleans) == bytes(expected), table.name
            checked += 1
    assert checked > 150


# --- data pages -----------------------------------------------------------------


def test_data_page_layout_matches_the_engine() -> None:
    page = DataPage.new(owner=95)
    assert page.free_space == 4082 and page.slot_count == 0
    rows = [bytes([3, 0, i, 0, 0, 0]) + b"x" * (20 + i) for i in range(5)]
    for row in rows:
        page.add_row(row)
    assert page.slots == [4096 - len(rows[0])] + [
        4096 - sum(len(r) for r in rows[: i + 1]) for i in range(1, 5)
    ]
    assert page.free_space == 4082 - sum(len(r) for r in rows) - 2 * 5
    for i, row in enumerate(rows):
        assert page.row(i) == row

    # Delete slot 2: rows below shift up, slot 2 dies at the boundary.
    page.remove_row(2)
    start2, end2 = 4096 - sum(len(r) for r in rows[:3]), 4096 - sum(len(r) for r in rows[:2])
    assert page.slots[2] == 0xC000 | end2
    assert page.row(2) is None
    assert page.row(3) == rows[3] and page.row(4) == rows[4]
    assert page.slots[3] == start2 + len(rows[2]) - len(rows[3])
    assert page.free_space == 4082 - sum(len(r) for r in rows) - 10 + len(rows[2])
    assert page.slot_count == 5

    # Replace slot 3 with a longer row: later rows move down.
    longer = rows[3] + b"more"
    page.replace_row(3, longer)
    assert page.row(3) == longer and page.row(4) == rows[4]
    assert page.free_space == 4082 - sum(len(r) for r in rows) - 10 + len(rows[2]) - 4
    with pytest.raises(AccessError):
        page.add_row(b"y" * 5000)


# --- inserting, updating, deleting ---------------------------------------------------


def check_indexes(table: Table) -> None:
    db = table.database
    rows: dict[tuple[int, int], dict[str, object]] = {}
    for page, slot, data in table.raw_rows():
        rows[(page, slot)] = table.decode(split_row(table.definition, data))
    assert len(rows) == table.row_count
    for index in table.indexes:
        expected = len(rows)
        if index.ignores_nulls:
            expected -= sum(1 for r in rows.values() if all(r[c.name] is None for c, _ in index.columns))
        previous: bytes | None = None
        seen: set[tuple[int, int]] = set()
        count = 0
        for entry in leaf_entries(db.store, index.real.root_page):
            count += 1
            assert previous is None or entry.key + bytes([entry.page >> 16, entry.page >> 8 & 255, entry.page & 255, entry.row]) > previous
            previous = entry.key + bytes([entry.page >> 16, entry.page >> 8 & 255, entry.page & 255, entry.row])
            assert (entry.page, entry.row) in rows
            assert (entry.page, entry.row) not in seen
            seen.add((entry.page, entry.row))
            row = rows[(entry.page, entry.row)]
            for (column, _asc), value in zip(index.columns, decode_key(entry.key, index.columns)):
                if value is None:
                    assert row[column.name] is None
                elif not isinstance(value, float) and column.type_code != 0x0A:
                    assert value == row[column.name], (index.name, column.name)
        assert count == expected, (index.name, count, expected)
        for node in node_pages(db.store, index.real.root_page):
            for node_entry in node.entries:
                assert node_entry.child is not None
                child = parse_index_page(db.store, node_entry.child)
                last = child.entries[-1]
                assert (last.key, last.page, last.row) == (node_entry.key, node_entry.page, node_entry.row)


def test_unchanged_database_saves_byte_for_byte(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    out = db.save(tmp_path / "copy.accdb")
    assert out.read_bytes() == LARGE.read_bytes()


def test_insert_assigns_autonumber_and_indexes_the_row(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    before = list(table.rows())
    assert [r["ID"] for r in before] == [1, 2, 3]
    row_id = table.insert_row({"Field1": "written by pyOpenVBA"})
    assert row_id.page in table.data_pages()
    db.save(tmp_path / "t1.accdb")

    again = AccessDatabase(tmp_path / "t1.accdb")
    table = again.table("Table1")
    rows = list(table.rows())
    assert len(rows) == 4 and table.row_count == 4
    assert rows[-1] == {"ID": 4, "Field1": "written by pyOpenVBA"}
    assert table.definition.next_autonumber == 4
    check_indexes(table)
    # The text went in compressed, as the column asks.
    raw = table.fetch_row(row_id.page, row_id.slot)
    assert raw is not None and b"\xff\xfewritten" in raw


def test_update_and_delete(tmp_path: Path) -> None:
    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    ids = {r["ID"]: rid for rid, r in table.rows_with_ids()}
    table.update_row(ids[2], {"Field1": "second, edited"})
    table.delete_row(ids[3])
    db.save(tmp_path / "t1.accdb")

    again = AccessDatabase(tmp_path / "t1.accdb")
    table = again.table("Table1")
    rows = {r["ID"]: r["Field1"] for r in table.rows()}
    assert rows == {1: "TEST1", 2: "second, edited"}
    assert table.row_count == 2
    check_indexes(table)
    # The deleted row's slot is dead at its boundary, engine style.
    page = DataPage(again.store.read(ids[3].page))
    assert page.slots[ids[3].slot] & 0xC000 == 0xC000


def test_many_inserts_allocate_pages_and_split_indexes(tmp_path: Path) -> None:
    """Enough rows to overflow the data page several times and to split the
    primary key's root leaf into a node with children."""
    from pyopenvba.access._pages import GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW, read_usage_map

    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    data_pages_before = table.data_pages()
    for i in range(1500):
        table.insert_row({"Field1": f"bulk row {i:04d}"})
    db.save(tmp_path / "bulk.accdb")

    again = AccessDatabase(tmp_path / "bulk.accdb")
    table = again.table("Table1")
    assert table.row_count == 1503
    # New pages came from the global free map (this fixture has free pages
    # inside it, which the engine would also reuse first) and are no longer free.
    new_pages = set(table.data_pages()) - set(data_pages_before)
    assert len(new_pages) >= 8
    free = set(read_usage_map(again.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert not (new_pages & free)
    rows = list(table.rows())
    assert len(rows) == 1503
    # Physical order follows page numbers, and reused pages may lie below
    # the original data page, so check membership rather than position.
    assert {r["Field1"] for r in rows} >= {"TEST1", "bulk row 0000", "bulk row 1499"}
    pk = table.primary_key
    assert pk is not None
    root = parse_index_page(again.store, pk.real.root_page)
    assert not root.is_leaf, "1503 nine-byte entries need more than one leaf"
    assert [r["ID"] for r in pk.rows()] == list(range(1, 1504))
    check_indexes(table)
    # Every data page is either in the free-space map or full enough to be out of it.
    from pyopenvba.access._pages import read_usage_map_ref

    owned = set(read_usage_map_ref(again.store, table.definition.owned_pages_ref).pages())
    assert set(table.data_pages()) <= owned
    assert len(table.data_pages()) >= 12


def test_random_order_keys_split_in_the_middle(tmp_path: Path) -> None:
    """Rows inserted with out-of-order key values must still yield a valid
    tree when the leaves split."""
    import random

    db = AccessDatabase(LARGE)
    table = db.table("Table1")
    values = list(range(10, 5000))
    random.Random(3).shuffle(values)
    for v in values[:1200]:
        table.insert_row({"ID": v, "Field1": f"v{v}"})
    db.save(tmp_path / "random.accdb")
    again = AccessDatabase(tmp_path / "random.accdb")
    table = again.table("Table1")
    pk = table.primary_key
    assert pk is not None
    ids: list[int] = []
    for r in pk.rows():
        assert isinstance(r["ID"], int)
        ids.append(r["ID"])
    assert ids == sorted(ids) and len(ids) == 1203
    check_indexes(table)


def test_writes_refuse_what_is_not_supported_yet() -> None:
    db = AccessDatabase(LARGE)
    with pytest.raises(AccessError):
        db.table("Table1").insert_row({"NoSuchColumn": 1})
    with pytest.raises(AccessError):
        db.table("Table1").delete_row(RowId(page=db.table("Table1").data_pages()[0], slot=99))
    assert struct.calcsize("<I") == 4
