"""Creating and dropping tables.

The structures are the ones the engine writes for CREATE TABLE, CREATE
INDEX and DROP TABLE (docs/access_engine.md); the live gate has the engine
use the created table and compares the bytes.
"""

from __future__ import annotations

import datetime as dt
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
    from pyopenvba.access._schema import serialize_definition

    db = AccessDatabase(path)
    checked = 0
    for page in range(db.store.page_count):
        if db.store.page_type(page) != PAGE_TDEF:
            continue
        definition = parse_table_definition(db.store, page)
        if len(definition.pages) != 1:
            continue
        raw = db.store.read(page)
        assert serialize_definition(definition)[: definition.definition_length] == raw[: definition.definition_length], page
        assert serialize_definition(definition)[:4] == raw[:4]
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
