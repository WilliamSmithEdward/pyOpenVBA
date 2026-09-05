"""Properties: the ``MR2`` blob in a catalog row's LvProp, read and written.

The layout was measured on a blob DAO wrote (a table Description plus a
field Caption and Description) and holds for every Access-authored blob
in the fixtures, which serialize back byte for byte.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, IndexSpec, PropertyValue
from pyopenvba.access._props import (
    DB_BOOLEAN,
    DB_INTEGER,
    DB_LONG,
    DB_TEXT,
    PropertyBlob,
    parse_property_blob,
    serialize_property_blob,
)
from pyopenvba.access_read import AccessError
from test_access_write import check_indexes

HERE = Path(__file__).parent
TEMPLATE = HERE.parents[0] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"
FIXTURES = [
    TEMPLATE,
    HERE / "live_access_test" / "New Microsoft Access Database.accdb",
    HERE / "live_access_test" / "two_modules_one_page.accdb",
    HERE / "live_access_test" / "module_spanning_pages.accdb",
]

def test_every_authored_blob_round_trips_byte_for_byte() -> None:
    checked = 0
    for path in FIXTURES:
        db = AccessDatabase(path)
        for row in db.table("MSysObjects").rows():
            lv = row["LvProp"]
            if isinstance(lv, bytes) and lv[:4] == b"MR2\0":
                assert serialize_property_blob(parse_property_blob(lv)) == lv, (path.name, row["Name"])
                checked += 1
    assert checked >= 12


def test_database_and_table_properties_decode() -> None:
    db = AccessDatabase(HERE / "live_access_test" / "New Microsoft Access Database.accdb")
    settings = db.database_properties()
    assert settings["AccessVersion"] == "09.50" and settings["ANSI Query Mode"] == 0
    table = db.table("Table1")
    props = table.properties()
    assert props["OrderByOn"] is False and props["Orientation"] == 0
    field1 = table.column_properties("Field1")
    assert field1["Required"] is False and field1["AllowZeroLength"] is True and field1["ColumnWidth"] == -1
    blob = table.property_blob()
    assert blob.column_properties["ID"]["ColumnWidth"].type == DB_INTEGER and len(blob.column_properties["ID"]["ColumnWidth"].raw) == 4


def test_setting_properties_writes_what_dao_writes() -> None:
    """The three appends DAO made -- a table Description, a field Caption,
    a field Description -- produce this exact blob."""
    dao_blob = (
        b"MR2\0"
        + (46).to_bytes(4, "little") + b"\x80\x00"
        + (22).to_bytes(2, "little") + "Description".encode("utf-16le")
        + (14).to_bytes(2, "little") + "Caption".encode("utf-16le")
        + (64).to_bytes(4, "little") + b"\x00\x00" + b"\x06\x00\x00\x00\x00\x00"
        + b"\x34\x00\x00\x0a\x00\x00\x2c\x00" + "Table described by DAO".encode("utf-16le")
        + (100).to_bytes(4, "little") + b"\x01\x00" + b"\x0e\x00\x00\x00\x08\x00" + "Name".encode("utf-16le")
        + b"\x1c\x00\x00\x0a\x01\x00\x14\x00" + "Name shown".encode("utf-16le")
        + b"\x34\x00\x00\x0a\x00\x00\x2c\x00" + "Field described by DAO".encode("utf-16le")
    )
    db = AccessDatabase(TEMPLATE)
    table = db.create_table("Parent", [ColumnSpec("Id", "Long", autonumber=True), ColumnSpec("Name", "Text", size=50, compressed=False)], [IndexSpec("PK", ("Id",), primary=True)])
    assert table.properties() == {} and table.column_properties("Name") == {}
    table.set_properties({"Description": "Table described by DAO"})
    table.set_properties({"Caption": "Name shown"}, column="Name")
    table.set_properties({"Description": "Field described by DAO"}, column="Name")
    row = next(r for r in db.table("MSysObjects").rows() if r["Name"] == "Parent")
    assert row["LvProp"] == dao_blob
    assert table.properties() == {"Description": "Table described by DAO"}
    assert table.column_properties("Name") == {"Caption": "Name shown", "Description": "Field described by DAO"}
    check_indexes(db.table("MSysObjects"))

    # Other types, an explicit PropertyValue, and a replaced value keeping its type.
    table.set_properties({"Hidden": True, "Width": 1200, "Big": 70000, "Ratio": 0.5, "Stamp": dt.datetime(2026, 9, 2, 8, 0), "Raw": PropertyValue(DB_LONG, 1, b"\x2a\x00\x00\x00")})
    props = table.properties()
    assert props["Hidden"] is True and props["Width"] == 1200 and props["Big"] == 70000 and props["Ratio"] == 0.5
    assert props["Stamp"] == dt.datetime(2026, 9, 2, 8, 0) and props["Raw"] == 42
    blob = table.property_blob()
    assert blob.object_properties["Width"].type == DB_INTEGER and blob.object_properties["Big"].type == DB_LONG
    assert blob.object_properties["Hidden"].type == DB_BOOLEAN and blob.object_properties["Raw"].flags == 1
    table.set_properties({"Description": "Changed"})
    assert table.property_blob().object_properties["Description"].type == DB_TEXT and table.properties()["Description"] == "Changed"
    assert list(table.property_blob().object_properties) == ["Description", "Hidden", "Width", "Big", "Ratio", "Stamp", "Raw"]
    with pytest.raises(AccessError):
        table.set_properties({"Caption": "x"}, column="Nope")
    with pytest.raises(AccessError):
        table.set_properties({"Odd": object()})
    again = AccessDatabase(db.to_bytes()).table("Parent")
    assert again.column_properties("Name")["Caption"] == "Name shown"


def test_an_empty_blob_serializes_to_just_the_names_block() -> None:
    assert serialize_property_blob(PropertyBlob()) == b"MR2\0" + (6).to_bytes(4, "little") + b"\x80\x00"
    with pytest.raises(AccessError):
        parse_property_blob(b"KKD\0")


def test_a_database_property_appends_as_dao_appends_it(tmp_path: Path) -> None:
    """DAO's Properties.Append on a database written here put StartUpForm
    at the end of the name list and the object block and left the MSysDb
    row's stamps alone; the same property set here gives the same blob."""
    from pyopenvba.access import AccessDatabase

    db = AccessDatabase.create_new(tmp_path / "startup.accdb")
    before = next(row for row in db.table("MSysObjects").rows() if row["Name"] == "MSysDb")
    db.set_database_properties({"StartUpForm": "Calculator"})
    after = next(row for row in db.table("MSysObjects").rows() if row["Name"] == "MSysDb")
    assert after["LvProp"] == (Path(__file__).parent / "fixtures" / "msysdb_startupform.bin").read_bytes()
    assert after["DateUpdate"] == before["DateUpdate"]
    assert db.database_properties()["StartUpForm"] == "Calculator"

    db.set_database_properties({"StartUpForm": "Other", "AppTitle": "Orders"})
    assert db.database_properties()["StartUpForm"] == "Other"
    assert db.database_properties()["AppTitle"] == "Orders"

