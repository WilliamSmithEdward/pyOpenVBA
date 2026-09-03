"""Creating and dropping tables.

The structures are the ones the engine writes for CREATE TABLE, CREATE
INDEX and DROP TABLE (docs/access_engine.md); the live gate has the engine
use the created table and compares the bytes.
"""

from __future__ import annotations

import datetime as dt
import struct
from decimal import Decimal
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec, Table
from pyopenvba.access._pages import GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW, read_usage_map, read_usage_map_ref
from pyopenvba.access._tdef import parse_table_definition
from pyopenvba.access_read import AccessError
from test_access_write import check_indexes

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
WHEN = dt.datetime(2026, 9, 2, 12, 0, 0)


def _all_types() -> list[ColumnSpec]:
    return [
        ColumnSpec("Id", "Long", autonumber=True),
        ColumnSpec("Flag", "Boolean"),
        ColumnSpec("Tiny", "Byte"),
        ColumnSpec("Small", "Integer"),
        ColumnSpec("Big", "Long"),
        ColumnSpec("Cash", "Currency"),
        ColumnSpec("Sgl", "Single"),
        ColumnSpec("Dbl", "Double"),
        ColumnSpec("Stamp", "DateTime"),
        ColumnSpec("Bin", "Binary", size=50),
        ColumnSpec("Txt", "Text", size=100),
        ColumnSpec("Blob", "OLE"),
        ColumnSpec("Story", "Memo"),
        ColumnSpec("Uid", "GUID"),
        ColumnSpec("Frac", "Decimal", size=(18, 4)),
        ColumnSpec("Huge", "BigInt"),
    ]


def test_create_table_matches_the_engines_layout(tmp_path: Path) -> None:
    db = AccessDatabase(TEMPLATE)
    pages_before = db.store.page_count
    table = db.create_table(
        "Simple",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=50, compressed=False)],
        [IndexSpec("PrimaryKey", ("Id",), primary=True), IndexSpec("IX_N", ("N",))],
        created=WHEN,
    )
    d = table.definition
    assert d.definition_length == 63 + 2 * 12 + 3 * 25 + (6 + 4 + 4) + 2 * 52 + 2 * 28 + (2 + 20 + 2 + 8) + 2
    assert d.table_type == 0x4E and d.max_columns == 3 and d.var_column_count == 1
    assert [c.name for c in d.columns_by_number()] == ["Id", "N", "T"]
    assert d.column("Id").flags == 0x07 and d.column("N").fixed_offset == 4 and d.column("T").length == 100
    assert not d.column("T").compressed_unicode
    pk = table.primary_key
    assert pk is not None and pk.name == "PrimaryKey" and pk.unique
    assert pk.real.flags == 0x89 and table.index("IX_N").real.flags == 0x80
    raw = db.store.read(d.page)
    assert int.from_bytes(raw[2:4], "little") == 4096 - d.definition_length - 8
    assert d.owned_pages_ref >> 8 == d.free_space_pages_ref >> 8
    assert d.real_indexes[0].raw[:4] == bytes.fromhex("83070000")
    assert d.logical_indexes[0].raw[21:23] == b"\x04\x04"
    # Two pages for the table plus one root per index came from the free map.
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert d.page not in free and pk.real.root_page not in free
    assert db.store.page_count >= pages_before
    entry = next(e for e in db.catalog() if e.name == "Simple")
    assert entry.id == d.page and entry.type == 1 and entry.flags == 0
    aces = [r for r in db.table("MSysACEs").rows() if r["ObjectId"] == d.page]
    assert len(aces) == 3 and all(r["ACM"] == 0xFFEFF for r in aces)

    db.save(tmp_path / "created.accdb")
    again = AccessDatabase(tmp_path / "created.accdb")
    assert "Simple" in again.table_names()
    table = again.table("Simple")
    for i in range(1, 8):
        table.insert_row({"N": i * 10, "T": f"row {i}"})
    assert [r["Id"] for r in table.rows()] == list(range(1, 8))
    check_indexes(table)


def test_every_column_type_can_be_created_and_written(tmp_path: Path) -> None:
    import uuid

    db = AccessDatabase(TEMPLATE)
    table = db.create_table("AllTypes", _all_types(), [IndexSpec("PK", ("Id",), primary=True), IndexSpec("IX_Txt", ("Txt",))], created=WHEN)
    assert table.definition.column("Uid").is_fixed is False
    assert table.definition.column("Bin").is_fixed and table.definition.column("Bin").length == 50
    assert table.definition.column("Frac").precision == 18 and table.definition.column("Frac").scale == 4
    assert set(table.definition.column_usage_maps) == {11, 12}
    row = {
        "Flag": True, "Tiny": 7, "Small": -300, "Big": 123456, "Cash": Decimal("12.5"), "Sgl": 0.25, "Dbl": 2.75,
        "Stamp": dt.datetime(2020, 5, 6, 7, 8, 9), "Bin": b"\x01\x02", "Txt": "Привет", "Blob": bytes(range(200)),
        "Story": "memo " * 500, "Uid": uuid.UUID(int=7), "Frac": Decimal("-1.2500"), "Huge": 10**12,
    }
    table.insert_row(row)
    db.save(tmp_path / "alltypes.accdb")
    again = AccessDatabase(tmp_path / "alltypes.accdb")
    got = next(iter(again.table("AllTypes").rows()))
    assert got["Id"] == 1 and got["Flag"] is True and got["Txt"] == "Привет" and got["Story"] == "memo " * 500
    assert got["Bin"] == b"\x01\x02" + bytes(48) and got["Frac"] == Decimal("-1.2500") and got["Huge"] == 10**12
    check_indexes(again.table("AllTypes"))


def test_drop_table_releases_everything(tmp_path: Path) -> None:
    db = AccessDatabase(TEMPLATE)
    before = db.to_bytes()
    table = db.create_table("Gone", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("M", "Memo")], [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    for i in range(40):
        table.insert_row({"M": "x" * (5000 if i % 10 == 0 else 20)})
    d = table.definition
    held = {d.page} | set(table.data_pages()) | {d.real_indexes[0].root_page}
    db.drop_table("Gone")
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert held <= free
    assert db.store.read(d.page)[0] == 0x08
    assert "Gone" not in db.table_names()
    assert not [r for r in db.table("MSysACEs").rows() if r["ObjectId"] == d.page]
    with pytest.raises(AccessError):
        db.table("Gone")
    # The catalog is back to its original row count and its indexes hold.
    assert db.table("MSysObjects").row_count == AccessDatabase(before).table("MSysObjects").row_count
    check_indexes(db.table("MSysObjects"))
    check_indexes(db.table("MSysACEs"))
    db.save(tmp_path / "dropped.accdb")
    AccessDatabase(tmp_path / "dropped.accdb").tables(include_system=True)


def test_definition_of_a_created_table_reparses_identically() -> None:
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("Re", [ColumnSpec("A", "Text", size=10), ColumnSpec("B", "Double")], [IndexSpec("IX_B", (("B", False),))], created=WHEN)
    parsed = parse_table_definition(db.store, table.definition.page)
    assert [(c.name, c.type_name, c.is_fixed) for c in parsed.columns_by_number()] == [("A", "Text", False), ("B", "Double", True)]
    assert parsed.real_indexes[0].columns[0].ascending is False


@pytest.mark.parametrize(
    "path",
    [
        Path(__file__).parent / "live_access_test" / "New Microsoft Access Database.accdb",
        Path(__file__).parent / "live_access_test" / "two_modules_one_page.accdb",
        TEMPLATE,
    ],
    ids=lambda p: p.name,
)
def test_every_definition_reserializes_byte_for_byte(path: Path) -> None:
    """The definition writer must give back exactly the page the engine
    wrote for every table, one-page definitions being all it writes."""
    from pyopenvba.access._pages import PAGE_TDEF
    from pyopenvba.access._schema import definition_pages, serialize_definition

    db = AccessDatabase(path)
    checked = 0
    for page in range(db.store.page_count):
        if db.store.page_type(page) != PAGE_TDEF:
            continue
        definition = parse_table_definition(db.store, page)
        if definition.pages[0] != page:
            continue  # a continuation page, checked with its first page
        images = definition_pages(serialize_definition(definition), definition.pages[1:])
        for index, (page_number, image) in enumerate(zip(definition.pages, images, strict=True)):
            raw = db.store.read(page_number)
            if index == 0:
                length = min(definition.definition_length, 4096)
                assert image[:length] == raw[:length], page_number
            else:
                carried = definition.definition_length - 4096 - 4088 * (index - 1)
                length = 8 + max(0, min(4088, carried))
                assert image[:length] == raw[:length], page_number
        checked += 1
    assert checked >= 19


def test_create_index_on_a_populated_table(tmp_path: Path) -> None:
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("Idx", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=30)], [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    for i in range(300):
        table.insert_row({"N": (i * 7919) % 1000, "T": f"t{i % 10}"})
    db.create_index("Idx", IndexSpec("IX_N", ("N",)), updated=WHEN)
    db.create_index("Idx", IndexSpec("IX_T", (("T", False),)), updated=WHEN)
    db.save(tmp_path / "idx.accdb")
    again = AccessDatabase(tmp_path / "idx.accdb")
    table = again.table("Idx")
    assert [i.name for i in table.indexes] == ["IX_N", "IX_T", "PK"]  # stored sorted by name
    assert [r.root_page for r in table.definition.real_indexes] == sorted(r.root_page for r in table.definition.real_indexes)
    assert table.index("IX_N").distinct_count == len({(i * 7919) % 1000 for i in range(300)})
    check_indexes(table)
    ns = [k[0] for k, _p, _r in table.index("IX_N").entries()]
    assert ns == sorted(ns, key=lambda v: (v is None, v))  # pyright: ignore[reportUnknownLambdaType, reportUnknownArgumentType]
    with pytest.raises(AccessError):
        db.create_index("Idx", IndexSpec("IX_N", ("N",)))


def test_create_new_gives_a_database_the_engine_would_recognize(tmp_path: Path) -> None:
    import pyopenvba

    path = tmp_path / "fresh" / "new.accdb"
    with pyopenvba.AccessDatabase.create_new(path) as db:
        assert db.header.version == 2 and db.table_names() == []
        table = db.create_table("T", [pyopenvba.ColumnSpec("Id", "Long", autonumber=True), pyopenvba.ColumnSpec("N", "Text")], [pyopenvba.IndexSpec("PK", ("Id",), primary=True)])
        table.insert_row({"N": "one"})
        db.save()
    again = AccessDatabase(path)
    assert [r["N"] for r in again.table("T").rows()] == ["one"]
    assert path.read_bytes()[4:20] == b"Standard ACE DB\x00"


def test_currency_rounds_half_even_to_four_places(tmp_path: Path) -> None:
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("Money", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Cash", "Currency")], created=WHEN)
    for value in (19.99, 0.00005, 0.00015, Decimal("1.23456"), 7, Decimal("-922337203685477.5807")):
        table.insert_row({"Cash": value})
    assert [r["Cash"] for r in table.rows()] == [
        Decimal("19.99"), Decimal("0"), Decimal("0.0002"), Decimal("1.2346"), Decimal("7"), Decimal("-922337203685477.5807"),
    ]
    for bad in (float("nan"), Decimal("Infinity"), 10**15, "1.00"):
        with pytest.raises(AccessError):
            table.insert_row({"Cash": bad})


def _wide_columns(count: int, name_length: int) -> list[ColumnSpec]:
    names = [f"Column{k:03d}".ljust(name_length, "_") for k in range(1, count)]
    return [ColumnSpec("Id", "Long", autonumber=True), *(ColumnSpec(n, "Text", size=20, compressed=False) for n in names)]


def _definition_length(columns: list[ColumnSpec], index_names: tuple[str, ...]) -> int:
    fixed = 63 + (12 + 52 + 28) * len(index_names) + 25 * len(columns) + 2
    return fixed + sum(2 + 2 * len(c.name) for c in columns) + sum(2 + 2 * len(n) for n in index_names)


def test_a_definition_over_one_page_is_chained_as_the_engine_chains_it(tmp_path: Path) -> None:
    """Four pages: the first holds 4096 bytes, continuations 4088 each after
    an 8-byte header; continuation pages are allocated ascending and
    chained in reverse; the free word is 0 everywhere but the last page,
    where it is ``4088 * pages - length``."""
    db = AccessDatabase(TEMPLATE)
    columns = _wide_columns(151, 30)
    length = _definition_length(columns, ("PK",))
    assert 4096 + 2 * 4088 < length <= 4 * 4088
    table = db.create_table("T4", columns, [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    d = table.definition
    assert d.definition_length == length and len(d.pages) == 4
    first, *chain = d.pages
    root = d.real_indexes[0].root_page
    assert chain == sorted(chain, reverse=True) and chain[-1] > root
    assert db.store.read(first)[2:8] == struct.pack("<HI", 0, chain[0])
    assert db.store.read(chain[0])[2:8] == struct.pack("<HI", 0, chain[1])
    assert db.store.read(chain[1])[2:8] == struct.pack("<HI", 0, chain[2])
    assert db.store.read(chain[2])[2:8] == struct.pack("<HI", 4 * 4088 - length, 0)
    body = 8 + (length - 4096 - 2 * 4088)
    assert db.store.read(chain[2])[body:] == bytes(4096 - body)
    parsed = parse_table_definition(db.store, first)
    assert [c.name for c in parsed.columns_by_number()] == [c.name for c in columns]
    for i in range(5):
        table.insert_row({c.name: f"r{i}c{k}" for k, c in enumerate(columns[1:]) if (k + i) % 4})
    db.save(tmp_path / "wide.accdb")
    again = AccessDatabase(tmp_path / "wide.accdb").table("T4")
    assert [r[columns[3].name] for r in again.rows()] == ["r0c2", "r1c2", None, "r3c2", "r4c2"]
    check_indexes(again)


@pytest.mark.parametrize("length, pages", [(4088, 1), (4089, 2), (4096, 2), (4100, 2), (8176, 2), (8177, 3)])
def test_definition_page_count_turns_at_4088_byte_shares(length: int, pages: int) -> None:
    from pyopenvba.access._schema import definition_page_count

    assert definition_page_count(length) == pages
    columns = _wide_columns(111, 4)
    short = _definition_length(columns, ("PK",))
    if length - short in range(0, 2000, 2):
        columns[-1] = ColumnSpec(columns[-1].name + "x" * ((length - short) // 2), "Text", size=20, compressed=False)
        db = AccessDatabase(TEMPLATE)
        d = db.create_table("Edge", columns, [IndexSpec("PK", ("Id",), primary=True)], created=WHEN).definition
        assert d.definition_length == length and len(d.pages) == pages
        last = db.store.read(d.pages[-1])
        assert int.from_bytes(last[2:4], "little") == 4088 * pages - length
        if pages == 2 and length <= 4096:
            assert last[8:] == bytes(4088)  # the continuation carries nothing but its header


def test_rewriting_a_definition_replaces_its_continuation_pages() -> None:
    db = AccessDatabase(TEMPLATE)
    columns = _wide_columns(80, 30)  # 7067 bytes: two pages before and after the index
    table = db.create_table("Two", columns, [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    first, old = table.definition.pages
    db.create_index("Two", IndexSpec("IX", (columns[1].name,)), updated=WHEN)
    d = db.table("Two").definition
    assert d.pages[0] == first and len(d.pages) == 2 and d.pages[1] != old
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert old in free and d.pages[1] not in free
    assert db.store.read(old)[0] == 0x02  # released, bytes kept
    assert d.real_indexes[1].root_page < d.pages[1]  # root first, then the fresh chain
    db.drop_table("Two")
    assert db.store.read(first)[0] == 0x08 and db.store.read(d.pages[1])[0] == 0x02
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert {first, d.pages[1]} <= free


def test_pages_released_in_a_session_are_reused_only_after_reopening() -> None:
    """The engine passes over pages it released earlier in the same session
    and hands them out again, lowest first, once the database is reopened.
    The page numbers are the ones DAO produced for this very sequence on
    the blank template."""
    db = AccessDatabase(TEMPLATE)
    spec = [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long")]
    key = [IndexSpec("PK", ("Id",), primary=True)]

    def pages(name: str, database: AccessDatabase = db) -> tuple[int, int, int]:
        d = database.create_table(name, spec, key, created=WHEN).definition
        return d.page, d.owned_pages_ref >> 8, d.real_indexes[0].root_page

    assert pages("A") == (95, 115, 116)
    assert pages("B") == (117, 118, 121)  # 119 and 120 went to the catalog's own growth
    db.drop_table("A")
    assert pages("C") == (122, 123, 124) and db.store.page_count == 125
    assert {95, 115, 116} <= set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    again = AccessDatabase(db.to_bytes())
    assert pages("D", again) == (95, 115, 116) and again.store.page_count == 125


def test_a_stamp_that_no_datetime_can_carry_survives_updates_and_indexing() -> None:
    """A serial whose last bit a datetime loses (one of the two doubles
    next to an engine stamp) is stored bit for bit when given as a float,
    kept through an update of another column, and its index entry is
    built, found and removed from the exact serial."""
    import math
    import struct

    from pyopenvba.access._rows import decode_datetime, encode_datetime, split_row

    def round_trip(value: float) -> float:
        return struct.unpack("<d", encode_datetime(decode_datetime(struct.pack("<d", value))))[0]

    stamp = 46267.46060038194  # 2026-09-02 11:03:15.873, as the engine wrote it
    serial = next(c for c in (math.nextafter(stamp, math.inf), math.nextafter(stamp, -math.inf)) if round_trip(c) != c)
    db = AccessDatabase(TEMPLATE)
    table = db.create_table(
        "Stamps",
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("When", "DateTime"), ColumnSpec("N", "Long")],
        [IndexSpec("PK", ("Id",), primary=True), IndexSpec("IX_When", ("When",))],
        created=WHEN,
    )
    row_id = table.insert_row({"When": serial, "N": 1})

    def stored() -> float:
        raw = table.fetch_row(row_id.page, row_id.slot)
        assert raw is not None
        stored_bytes = split_row(table.definition, raw).values[table.definition.column("When").number]
        assert isinstance(stored_bytes, bytes)
        return struct.unpack("<d", stored_bytes)[0]

    assert stored() == serial
    table.update_row(row_id, {"N": 2})
    assert stored() == serial
    check_indexes(table)
    assert [k[0] for k, _p, _r in table.index("IX_When").entries()] == [decode_datetime(struct.pack("<d", serial))]
    table.delete_row(row_id)
    assert list(table.index("IX_When").entries()) == []


def _filled_table(db: AccessDatabase, name: str, rows: int, memo_at: tuple[int, ...] = ()) -> Table:
    """A table whose 500-byte rows span several pages, with single-row
    long values for the rows named in ``memo_at``."""
    table = db.create_table(
        name,
        [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("T", "Text", size=255, compressed=False), ColumnSpec("M", "Memo", compressed=False)],
        [IndexSpec("PK", ("Id",), primary=True)],
        created=WHEN,
    )
    for i in range(1, rows + 1):
        table.insert_row({"T": f"t{i:02d}" * 80, "M": (f"m{i:02d}" * 500) if i in memo_at else "short"})
    return table


def test_an_emptied_page_is_retired_unless_it_is_the_first() -> None:
    from pyopenvba.access._pages import PAGE_RETIRED, row_slots

    db = AccessDatabase(TEMPLATE)
    table = _filled_table(db, "R", 24, memo_at=(3, 10, 20))
    pages = list(table.data_pages())
    assert len(pages) == 4, pages
    first, second, third, last = pages
    ids = {row["Id"]: rid for rid, row in table.rows_with_ids()}
    for i in range(13, 25):
        table.delete_row(ids[i])
    raw = db.store.read(third)
    assert raw[0] == PAGE_RETIRED and set(row_slots(raw)) == {0xD000}
    assert int.from_bytes(raw[2:4], "little") == 4096 - 14 - 2 * len(row_slots(raw))
    assert db.store.read(last)[0] == PAGE_RETIRED
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert {third, last} <= free and second not in free
    d = table.definition
    assert list(read_usage_map_ref(db.store, d.owned_pages_ref).pages()) == [first, second]
    assert list(read_usage_map_ref(db.store, d.free_space_pages_ref).pages()) == [second]  # it lost rows, so it rejoined
    # The memo of row 20 sat alone on its LVAL page, which is retired too.
    lv_owned, _lv_free = d.column_usage_maps[d.column("M").number]
    assert len(list(read_usage_map_ref(db.store, lv_owned).pages())) == 2
    # Emptying the first page keeps it.
    for i in range(1, 8):
        table.delete_row(ids[i])
    assert db.store.read(first)[0] == 0x01 and first not in set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert list(read_usage_map_ref(db.store, d.owned_pages_ref).pages()) == [first, second]
    check_indexes(table)


def test_truncate_releases_pages_untouched_and_resets_indexes() -> None:
    db = AccessDatabase(TEMPLATE)
    table = _filled_table(db, "R", 24, memo_at=(3, 10, 20))
    d = table.definition
    pages = list(table.data_pages())
    lv_pages = list(read_usage_map_ref(db.store, d.column_usage_maps[d.column("M").number][0]).pages())
    images = {p: db.store.read(p) for p in pages + lv_pages}
    table.truncate()
    assert table.row_count == 0 and list(table.rows()) == [] and d.next_autonumber == 24
    for p, image in images.items():
        assert db.store.read(p) == image  # rows left in place
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert set(images) <= free
    assert list(read_usage_map_ref(db.store, d.owned_pages_ref).pages()) == []
    for owned, fsp in d.column_usage_maps.values():
        assert list(read_usage_map_ref(db.store, owned).pages()) == [] and list(read_usage_map_ref(db.store, fsp).pages()) == []
    root = db.store.read(d.real_indexes[0].root_page)
    assert int.from_bytes(root[2:4], "little") == 0x0E20 and d.real_indexes[0].entry_count == 0
    assert list(table.index("PK").entries()) == []
    table.insert_row({"T": "again", "M": "x"})
    assert [r["Id"] for r in table.rows()] == [25]
    check_indexes(table)


def test_create_relationship_writes_both_sides_and_the_catalog() -> None:
    """A foreign key as ALTER TABLE ADD CONSTRAINT writes it: a non-unique
    index named after the relationship on the child, a kind-2 logical
    entry on each side naming the other's definition page and logical
    number, ``.rB`` then ``.rC`` on the parent, one MSysRelationships row
    per column pair, a type-8 catalog object with three permission rows."""
    from pyopenvba.access import Relationship
    from pyopenvba.access._schema import RELATIONSHIP_REFERENCED, RELATIONSHIP_REFERENCING
    from pyopenvba.access.database import OBJECT_RELATIONSHIP, RELATIONSHIP_ACMS

    db = AccessDatabase(TEMPLATE)
    key = [IndexSpec("PK", ("Id",), primary=True)]
    db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50)], key, created=WHEN)
    child = db.create_table("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key, created=WHEN)
    for i in range(5):
        child.insert_row({"ParentId": i % 2})
    before = len(db.relationships())
    rel = db.create_relationship("FK_Child_Parent", "Child", ("ParentId",), "Parent", ("Id",), created=WHEN)
    assert rel == Relationship("FK_Child_Parent", "Child", ("ParentId",), "Parent", ("Id",), 0)
    assert db.relationships()[-1] == rel and len(db.relationships()) == before + 1
    assert not rel.cascade_updates and not rel.cascade_deletes and rel.enforced

    cd, pd = db.table("Child").definition, db.table("Parent").definition
    fk = cd.logical_indexes[0]
    assert [li.name for li in cd.logical_indexes] == ["FK_Child_Parent", "PK"]
    assert fk.kind == 2 and fk.relationship_kind == RELATIONSHIP_REFERENCING and fk.relationship_table_page == pd.page and fk.relationship_index == 1
    assert cd.real_indexes[fk.real_index].flags == 0x80 and cd.real_indexes[fk.real_index].entry_count == 2  # two distinct ParentId values
    back = pd.logical_indexes[0]
    assert [li.name for li in pd.logical_indexes] == [".rB", "PK"]
    assert back.kind == 2 and back.relationship_kind == RELATIONSHIP_REFERENCED and back.relationship_table_page == cd.page and back.relationship_index == 1
    assert back.real_index == 0 and not back.cascade_updates and not back.cascade_deletes
    entry = next(e for e in db.catalog() if e.name == "FK_Child_Parent")
    assert entry.type == OBJECT_RELATIONSHIP and entry.id < 0 and entry.parent_id == db._container("Relationships").id  # pyright: ignore[reportPrivateUsage]
    aces = sorted(int(r["ACM"]) for r in db.table("MSysACEs").rows() if r["ObjectId"] == entry.id)  # pyright: ignore[reportArgumentType]
    assert aces == sorted(RELATIONSHIP_ACMS)
    check_indexes(db.table("Child"))
    check_indexes(db.table("MSysRelationships"))
    check_indexes(db.table("MSysObjects"))

    # A second child gets ``.rC`` on the parent, with cascades recorded on both sides.
    db.create_table("Child2", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key, created=WHEN)
    rel2 = db.create_relationship("FK_Child2_Parent", "Child2", ("ParentId",), "Parent", ("Id",), cascade_updates=True, cascade_deletes=True, created=WHEN)
    assert rel2.attributes == 0x1100 and rel2.cascade_updates and rel2.cascade_deletes
    pd = db.table("Parent").definition
    assert [li.name for li in pd.logical_indexes] == [".rB", ".rC", "PK"]
    assert pd.logical_indexes[1].cascade_updates and pd.logical_indexes[1].cascade_deletes and pd.logical_indexes[1].relationship_index == 1
    assert db.table("Child2").definition.logical_indexes[0].relationship_index == 2

    with pytest.raises(AccessError):
        db.create_relationship("FK_Child_Parent", "Child", ("ParentId",), "Parent", ("Id",))
    with pytest.raises(AccessError):
        db.create_relationship("FK_Bad", "Child", ("ParentId",), "Parent", ("Name",))  # no unique index on Name
    with pytest.raises(AccessError):
        db.create_relationship("FK_Bad", "Child", ("Nope",), "Parent", ("Id",))

    # Dropping the first: its index and both logical entries go, ``.rC``
    # keeps its number, the catalog object and its permissions vanish.
    fk_root = db.table("Child").definition.real_indexes[1].root_page
    db.drop_relationship("FK_Child_Parent")
    assert [r.name for r in db.relationships()][-1:] == ["FK_Child2_Parent"]
    cd, pd = db.table("Child").definition, db.table("Parent").definition
    assert [li.name for li in cd.logical_indexes] == ["PK"] and len(cd.real_indexes) == 1
    assert [(li.name, li.number) for li in pd.logical_indexes] == [(".rC", 2), ("PK", 0)]
    assert not [e for e in db.catalog() if e.name == "FK_Child_Parent"]
    assert not [r for r in db.table("MSysACEs").rows() if r["ObjectId"] == entry.id]
    assert fk_root in set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    tail = db.store.read(pd.page)[pd.definition_length :]
    assert tail[:8] == bytes(8) and tail[8:24] != bytes(16)  # eight reserved bytes zeroed, the old tail beyond stays
    with pytest.raises(AccessError):
        db.drop_relationship("FK_Child_Parent")
    check_indexes(db.table("Child"))
    check_indexes(db.table("MSysObjects"))


def test_columns_are_added_and_dropped_as_alter_table_does() -> None:
    """Measured on ALTER TABLE: an added column takes the next number, a
    fixed one the offset past the highest fixed column, a variable one the
    next variable index; a dropped column leaves numbers, offsets and the
    maximum column count alone; rows are never rewritten."""
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("W", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=40)], [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    for i in range(1, 11):
        table.insert_row({"N": i, "T": f"row {i}"})
    extra = table.add_column(ColumnSpec("Extra", "Long"), updated=WHEN)
    assert (extra.number, extra.fixed_offset, extra.is_fixed) == (3, 8, True)
    d = db.table("W").definition
    assert d.max_columns == 4 and d.var_column_count == 1 and [c.name for c in d.columns_by_number()] == ["Id", "N", "T", "Extra"]
    remark = db.table("W").add_column(ColumnSpec("Remark", "Text", size=30), updated=WHEN)
    assert (remark.number, remark.var_index, remark.is_fixed) == (4, 1, False)
    assert db.table("W").definition.var_column_count == 2
    db.table("W").drop_column("Extra", updated=WHEN)
    d = db.table("W").definition
    assert [(c.name, c.number) for c in d.columns_by_number()] == [("Id", 0), ("N", 1), ("T", 2), ("Remark", 4)] and d.max_columns == 5
    again = db.table("W").add_column(ColumnSpec("Again", "Long"), updated=WHEN)
    assert (again.number, again.fixed_offset) == (5, 8)  # the dropped column's slot is reused
    db.table("W").drop_column("T", updated=WHEN)
    d = db.table("W").definition
    assert d.var_column_count == 2 and d.column("Remark").var_index == 1 and d.max_columns == 6
    rows = list(db.table("W").rows())
    assert [r["N"] for r in rows] == list(range(1, 11)) and rows[0] == {"Id": 1, "N": 1, "Remark": None, "Again": None}
    db.table("W").insert_row({"N": 11, "Remark": "new", "Again": 7})
    assert list(db.table("W").rows())[-1] == {"Id": 11, "N": 11, "Remark": "new", "Again": 7}
    check_indexes(db.table("W"))
    with pytest.raises(AccessError):
        db.table("W").drop_column("Id")  # indexed
    with pytest.raises(AccessError):
        db.table("W").add_column(ColumnSpec("N", "Long"))
    with pytest.raises(AccessError):
        db.table("W").add_column(ColumnSpec("Serial", "Long", autonumber=True))
    from pyopenvba.access._schema import serialize_definition

    saved = AccessDatabase(db.to_bytes())
    assert saved.table("W").column_names == ["Id", "N", "Remark", "Again"]
    assert serialize_definition(saved.table("W").definition) == serialize_definition(db.table("W").definition)

    # A Memo column brings two map rows and a map pair, and gives its
    # long-value pages back untouched when dropped.
    table = db.table("W")
    map_page = table.definition.owned_pages_ref >> 8
    slots_before = len(read_usage_map_ref(db.store, table.definition.owned_pages_ref).pages()), db.store.read(map_page)[12:14]
    notes = table.add_column(ColumnSpec("Notes", "Memo"), updated=WHEN)
    d = db.table("W").definition
    assert notes.number == 6 and notes.var_index == 2 and d.var_column_count == 3 and 6 in d.column_usage_maps
    assert db.store.read(map_page)[12:14] == (int.from_bytes(slots_before[1], "little") + 2).to_bytes(2, "little")
    db.table("W").insert_row({"N": 99, "Notes": "n" * 3000})
    owned_ref = db.table("W").definition.column_usage_maps[6][0]
    lv_pages = list(read_usage_map_ref(db.store, owned_ref).pages())
    assert len(lv_pages) == 2
    db.table("W").drop_column("Notes", updated=WHEN)
    d = db.table("W").definition
    assert 6 not in d.column_usage_maps and d.var_column_count == 3 and d.max_columns == 7
    assert [c.name for c in d.columns_by_number()] == ["Id", "N", "Remark", "Again"]
    free = set(read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW).pages())
    assert set(lv_pages) <= free and all(db.store.read(p)[0] == 0x01 for p in lv_pages)
    assert [r["N"] for r in db.table("W").rows()][-1] == 99 and db.table("W").row_count == 12
    check_indexes(db.table("W"))


def test_rename_table_follows_the_catalog_and_the_relationships() -> None:
    db = AccessDatabase(TEMPLATE)
    key = [IndexSpec("PK", ("Id",), primary=True)]
    db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50)], key, created=WHEN)
    child = db.create_table("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key, created=WHEN)
    child.insert_row({"ParentId": 1})
    db.create_relationship("FK_Child_Parent", "Child", ("ParentId",), "Parent", ("Id",), created=WHEN)
    page = db.table("Parent").definition.page
    db.rename_table("Parent", "Parents", updated=WHEN)
    assert "Parents" in db.table_names() and "Parent" not in db.table_names()
    assert db.table("Parents").definition.page == page
    assert db.relationships()[-1].referenced_table == "Parents" and db.relationships()[-1].table == "Child"
    db.rename_table("Child", "Children", updated=WHEN)
    assert db.relationships()[-1].table == "Children"
    assert [r["ParentId"] for r in db.table("Children").rows()] == [1]
    with pytest.raises(AccessError):
        db.rename_table("Children", "Parents")
    with pytest.raises(AccessError):
        db.rename_table("Nope", "X")
    check_indexes(db.table("MSysObjects"))
    check_indexes(db.table("MSysRelationships"))
    again = AccessDatabase(db.to_bytes())
    assert sorted(t for t in again.table_names() if not t.startswith("MSys")) == ["Children", "Parents"]


def test_map_rows_spill_onto_a_second_map_page() -> None:
    """Measured: a table's 58th usage-map row (32 indexes, 12 Memo columns)
    went to row 0 of a fresh map page, allocated just before the index root."""
    db = AccessDatabase(TEMPLATE)
    columns = [ColumnSpec("Id", "Long", autonumber=True), *(ColumnSpec(f"C{i:02d}", "Long") for i in range(1, 33)), *(ColumnSpec(f"M{i:02d}", "Memo", compressed=False) for i in range(1, 13))]
    table = db.create_table("Many", columns, [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    first_map = table.definition.owned_pages_ref >> 8
    for i in range(1, 32):
        db.create_index("Many", IndexSpec(f"IX{i:02d}", (f"C{i:02d}",)), updated=WHEN)
    d = db.table("Many").definition
    refs = [r.usage_map_ref for r in d.real_indexes]
    assert all(ref >> 8 == first_map for ref in refs[:-1]) and refs[-1] >> 8 != first_map and refs[-1] & 0xFF == 0
    spill = db.store.read(refs[-1] >> 8)
    assert spill[0] == 0x01 and spill[4:8] == bytes(4) and int.from_bytes(spill[12:14], "little") == 1
    assert len(d.real_indexes) == 32 and len(d.pages) == 2
    db.table("Many").insert_row({"C01": 1, "M01": "x" * 100})
    check_indexes(db.table("Many"))
    again = AccessDatabase(db.to_bytes())
    assert len(again.table("Many").indexes) == 32 and [r["C01"] for r in again.table("Many").rows()] == [1]


def test_rename_column_follows_properties_and_relationships() -> None:
    db = AccessDatabase(TEMPLATE)
    key = [IndexSpec("PK", ("Id",), primary=True)]
    parent = db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50)], key, created=WHEN)
    db.create_table("Child", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("ParentId", "Long")], key, created=WHEN)
    parent.set_properties({"Caption": "Name shown"}, column="Name")
    parent.insert_row({"Name": "Ada"})
    db.create_relationship("FK_Child_Parent", "Child", ("ParentId",), "Parent", ("Id",), created=WHEN)
    header = db.table("Parent").definition.column("Name").raw
    db.table("Parent").rename_column("Name", "FullName", updated=WHEN)
    d = db.table("Parent").definition
    assert [c.name for c in d.columns_by_number()] == ["Id", "FullName"] and d.column("FullName").raw == header
    assert db.table("Parent").column_properties("FullName") == {"Caption": "Name shown"} and db.table("Parent").property_blob().column_properties.keys() == {"FullName"}
    assert [r["FullName"] for r in db.table("Parent").rows()] == ["Ada"]
    db.table("Child").rename_column("ParentId", "PId", updated=WHEN)
    assert db.relationships()[-1].columns == ("PId",) and db.relationships()[-1].referenced_columns == ("Id",)
    db.table("Parent").rename_column("Id", "Key", updated=WHEN)
    assert db.relationships()[-1].referenced_columns == ("Key",) and db.table("Parent").primary_key is not None
    with pytest.raises(AccessError):
        db.table("Parent").rename_column("FullName", "Key")
    with pytest.raises(AccessError):
        db.table("Parent").rename_column("Nope", "X")
    check_indexes(db.table("Parent"))
    check_indexes(db.table("MSysObjects"))
    again = AccessDatabase(db.to_bytes())
    assert again.table("Parent").column_names == ["Key", "FullName"] and again.table("Child").column_names == ["Id", "PId"]


def test_alter_column_retypes_and_resizes_as_the_engine_does() -> None:
    """Measured: the new column takes the old one's place and position
    under the next number, rows are re-encoded with the old column's bytes
    left in as a phantom, the definition length does not change."""
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("W", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("N", "Long"), ColumnSpec("T", "Text", size=40, compressed=False)], [IndexSpec("PK", ("Id",), primary=True)], created=WHEN)
    for i in range(1, 11):
        table.insert_row({"N": i, "T": f"row {i}"})
    length = table.definition.definition_length
    first = next(iter(table.raw_rows()))[2]
    resized = table.alter_column("T", ColumnSpec("T", "Text", size=80, compressed=False), updated=WHEN)
    d = db.table("W").definition
    assert (resized.number, resized.var_index, resized.length) == (3, 1, 160) and d.definition_length == length
    assert [(c.name, c.number) for c in d.columns] == [("Id", 0), ("N", 1), ("T", 3)] and d.max_columns == 4 and d.var_column_count == 2
    row = next(iter(db.table("W").raw_rows()))[2]
    assert len(row) == len(first) + 12 and row[:2] == b"\x04\x00"  # column count 4, phantom text kept
    retyped = db.table("W").alter_column("N", ColumnSpec("N", "Double"), updated=WHEN)
    d = db.table("W").definition
    assert (retyped.number, retyped.fixed_offset, retyped.type_name) == (4, 8, "Double")
    assert [(c.name, c.number) for c in d.columns] == [("Id", 0), ("N", 4), ("T", 3)]  # the replacement keeps its place
    assert int.from_bytes(d.column("N").raw[9:11], "little") == 1 and int.from_bytes(d.column("N").raw[7:9], "little") == 2
    rows = list(db.table("W").rows())
    assert rows[0] == {"Id": 1, "T": "row 1", "N": 1.0} and [r["N"] for r in rows] == [float(i) for i in range(1, 11)]
    db.table("W").insert_row({"N": 2.5, "T": "x" * 60})
    assert list(db.table("W").rows())[-1]["N"] == 2.5
    check_indexes(db.table("W"))
    with pytest.raises(AccessError):
        db.table("W").alter_column("Id", ColumnSpec("Id", "Double"))  # indexed and AutoNumber
    with pytest.raises(AccessError):
        db.table("W").alter_column("T", ColumnSpec("Other", "Text", size=10))
    with pytest.raises(AccessError):
        db.table("W").alter_column("T", ColumnSpec("T", "Long"))
    again = AccessDatabase(db.to_bytes())
    assert [r["N"] for r in again.table("W").rows()][:3] == [1.0, 2.0, 3.0]


def test_bad_specs_are_refused() -> None:
    db = AccessDatabase(TEMPLATE)
    with pytest.raises(AccessError):
        db.create_table("MSysObjects", [ColumnSpec("A", "Long")])
    with pytest.raises(AccessError):
        db.create_table("X", [ColumnSpec("A", "Long"), ColumnSpec("a", "Text")])
    with pytest.raises(AccessError):
        db.create_table("X", [ColumnSpec("A", "Text", autonumber=True)])
    with pytest.raises(AccessError):
        db.create_table("X", [ColumnSpec("A", "Long")], [IndexSpec("I", ("B",))])
    with pytest.raises(AccessError):
        db.create_table("X", [ColumnSpec("A", "Wibble")])
    with pytest.raises(AccessError):
        db.drop_table("MSysObjects")
