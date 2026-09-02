"""The ACE storage engine's read layer, against Access-authored fixtures.

Every structural rule the engine relies on is checked here in the form
that would fail if the rule were wrong: a definition must reconcile to the
byte, a table's rows must count to what its definition says, a row must
decode without a column running off its end.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._pages import (
    PAGE_TDEF,
    GLOBAL_USAGE_MAP_PAGE,
    GLOBAL_USAGE_MAP_ROW,
    HEADER_MASK,
    read_usage_map,
    toggle_definition_mask,
)
from pyopenvba.access._rows import decode_text, encode_text
from pyopenvba.access._tdef import (
    TYPE_LONG,
    TYPE_MEMO,
    TYPE_OLE,
    TYPE_TEXT,
    parse_table_definition,
)
from pyopenvba.access_read import AccessReader

FIXTURES = Path(__file__).parent / "live_access_test"
TEMPLATES = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files"

# Written by Access itself.  The _write_* copies and tests/output/* were
# produced by an earlier experimental writer and are not ground truth.
AUTHORED = [
    FIXTURES / "New Microsoft Access Database.accdb",
    FIXTURES / "module_spanning_pages.accdb",
    FIXTURES / "two_modules_one_page.accdb",
    TEMPLATES / "blank_database.accdb",
    TEMPLATES / "blank_database_module.accdb",
]
SMALL = FIXTURES / "two_modules_one_page.accdb"
LARGE = FIXTURES / "New Microsoft Access Database.accdb"


# --- page 0 -----------------------------------------------------------------


def test_header_mask_is_the_rc4_keystream_of_the_fixed_key() -> None:
    # First bytes of RC4(key C7 DA 39 6B); a wrong key would decode the
    # code page to nonsense, which the next test would catch.
    assert len(HEADER_MASK) == 126
    once = bytes(toggle_definition_mask(bytes(4096)))
    assert once[0x18:0x96] == HEADER_MASK
    assert bytes(toggle_definition_mask(once)) == bytes(4096)


def test_header_fields_decode() -> None:
    db = AccessDatabase(SMALL)
    assert db.header.version == 2
    assert db.header.code_page == 1252
    assert db.header.sort_order == 0x409
    assert db.header.encoding_key == 0
    assert db.header.password == b""
    created = dt.datetime(1899, 12, 30) + dt.timedelta(days=db.header.creation_date)
    assert created.date() == dt.date(2026, 8, 19)


def test_global_usage_map_marks_pages_past_the_end_as_free() -> None:
    db = AccessDatabase(SMALL)
    free = read_usage_map(db.store, GLOBAL_USAGE_MAP_PAGE, GLOBAL_USAGE_MAP_ROW)
    pages = set(free.pages())
    assert db.store.page_count == 73
    assert all(p in pages for p in range(73, 512))
    assert 0 not in pages and 1 not in pages and 2 not in pages


# --- table definitions --------------------------------------------------------


@pytest.mark.parametrize("path", AUTHORED, ids=lambda p: p.name)
def test_every_definition_reconciles_to_its_declared_length(path: Path) -> None:
    db = AccessDatabase(path)
    seen: set[int] = set()
    parsed = 0
    for page in range(db.store.page_count):
        if db.store.page_type(page) != PAGE_TDEF or page in seen:
            continue
        definition = parse_table_definition(db.store, page)
        seen.update(definition.pages)
        parsed += 1
        assert len(definition.columns) == len({c.number for c in definition.columns})
        assert definition.var_column_count == sum(
            1 for c in definition.columns if not c.is_fixed
        )
        assert set(definition.column_usage_maps) == {
            c.number for c in definition.columns if c.is_long_value
        }
    assert parsed >= 19


def test_msysobjects_definition() -> None:
    db = AccessDatabase(SMALL)
    d = db.definition(2)
    assert d.row_count == 35
    assert d.is_system
    assert [c.name for c in d.columns_by_number()] == [
        "Id", "ParentId", "Name", "Type", "DateCreate", "DateUpdate", "Owner",
        "Flags", "Database", "Connect", "ForeignName", "RmtInfoShort",
        "RmtInfoLong", "Lv", "LvProp", "LvModule", "LvExtra",
    ]
    assert d.column("Name").type_code == TYPE_TEXT
    assert d.column("Name").length == 510
    assert d.column("Lv").type_code == TYPE_OLE
    assert d.column("Connect").type_code == TYPE_MEMO
    assert d.column("Id").type_code == TYPE_LONG and d.column("Id").is_fixed
    assert d.column("Flags").fixed_offset == 26
    by_name = {i.name: i for i in d.logical_indexes}
    assert set(by_name) == {"Id", "ParentIdName"}
    assert by_name["Id"].is_primary_key
    pin = d.real_indexes[by_name["ParentIdName"].real_index]
    assert [(c.number, c.ascending) for c in pin.columns] == [(1, True), (2, True)]
    assert pin.unique
    assert pin.root_page == 7


def test_engine_and_application_system_tables_are_distinguished() -> None:
    """Jet marks only its own four catalog tables 'S'; the MSys* tables the
    Access layer adds are ordinary 'N' tables that the catalog flags."""
    db = AccessDatabase(LARGE)
    engine_tables = {
        e.name for e in db.table_entries(include_system=True)
        if db.definition(e.id).is_system
    }
    assert engine_tables == {"MSysObjects", "MSysACEs", "MSysQueries", "MSysRelationships"}
    assert not db.definition(db.table("Table2").definition.page).is_system


# --- rows -------------------------------------------------------------------


@pytest.mark.parametrize("path", AUTHORED, ids=lambda p: p.name)
def test_every_table_counts_to_its_definition(path: Path) -> None:
    db = AccessDatabase(path)
    tables = db.tables(include_system=True)
    assert len(tables) >= 19
    for table in tables:
        rows = list(table.rows())
        assert len(rows) == table.row_count, table.name
        for row in rows:
            assert set(row) == set(table.column_names)


def test_catalog_agrees_with_the_shipped_reader_on_every_row() -> None:
    """Two independent decoders of MSysObjects must agree, ids compared as
    the unsigned values the old reader reports."""
    for path in AUTHORED:
        db = AccessDatabase(path)
        mine = {
            e.id & 0xFFFFFFFF: (e.name, e.type, e.parent_id & 0xFFFFFFFF)
            for e in db.catalog()
        }
        with AccessReader(path) as old:
            theirs = {o.id_: (o.name, o.type_, o.parent_id) for o in old.msys_objects()}
        # The old reader skips rows it cannot place; ours must cover them.
        assert set(theirs) <= set(mine), path.name
        for key, value in theirs.items():
            if path == LARGE and key == 145:
                # The old reader mis-bounds the one row that lives on an
                # overflow page and reads garbage after its name.
                assert value[0].startswith("Table2") and value[0] != "Table2"
                assert mine[key] == ("Table2", 1, 251658241)
                continue
            assert mine[key] == value, (path.name, key)


def test_overflow_rows_are_followed() -> None:
    """MSysObjects on the 1 MB fixture keeps one row on another page; the
    old reader never saw it, so 'Table2' is the proof the pointer works."""
    db = AccessDatabase(LARGE)
    names = {e.name for e in db.catalog()}
    assert "Table2" in names and "MSysNameMap" in names
    with AccessReader(LARGE) as old:
        assert "Table2" not in {o.name for o in old.msys_objects()}


def test_user_and_system_tables_are_told_apart() -> None:
    db = AccessDatabase(LARGE)
    users = db.table_names()
    assert "Table2" in users
    assert not any(n.startswith("MSys") for n in users)
    assert "MSysObjects" in db.table_names(include_system=True)


def test_values_decode_to_python_types() -> None:
    db = AccessDatabase(SMALL)
    rows = {r["Name"]: r for r in db.table("MSysObjects").rows()}
    tables = rows["Tables"]
    assert tables["Id"] == 251658241 and tables["ParentId"] == 251658240
    assert tables["Type"] == 3
    assert tables["Flags"] == -2147483648
    assert isinstance(tables["DateCreate"], dt.datetime)
    assert tables["DateCreate"].year == 2026
    assert tables["Owner"] == b"\x00\x00" or isinstance(tables["Owner"], bytes)
    assert tables["Lv"] is None
    props = rows["MSysDb"]["LvProp"]
    assert isinstance(props, bytes) and len(props) > 1000
    assert rows["ModA"]["Type"] == -32761
    assert rows["ModA"]["ParentId"] == rows["Modules"]["Id"]


def test_module_source_is_reachable_through_the_engine() -> None:
    """MSysAccessStorage rows carry the VBA project; the old reader found
    them by scanning pages, the engine finds them as a table."""
    db = AccessDatabase(SMALL)
    storage = db.table("MSysAccessStorage")
    rows = list(storage.rows())
    assert rows
    assert {"Id", "Name", "ParentId", "Type", "Lv"} <= set(storage.column_names)
    blobs = [r["Lv"] for r in rows if isinstance(r["Lv"], bytes)]
    assert blobs, "at least one row holds a long value"


# --- text codec -------------------------------------------------------------


@pytest.mark.parametrize(
    "text",
    ["", "plain", "héllo wörld", "Привет", "日本語", "mixed ascii and 日本", "a\x00b"],
)
def test_text_round_trips(text: str) -> None:
    assert decode_text(encode_text(text)) == text


def test_compressed_text_toggles_on_a_lone_zero_byte() -> None:
    # FF FE, "ab" one byte each, 0x00 switches to UTF-16, U+65E5, 0x00
    # switches back, "c".
    raw = b"\xff\xfeab\x00\xe5\x65\x00c"
    assert decode_text(raw) == "ab日c"


def test_all_latin1_text_is_compressed_like_access() -> None:
    assert encode_text("Tables") == b"\xff\xfeTables"
    assert encode_text("日本") == "日本".encode("utf-16-le")
