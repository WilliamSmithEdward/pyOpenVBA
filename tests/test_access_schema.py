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

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec
from pyopenvba.access._pages import GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW, read_usage_map
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
