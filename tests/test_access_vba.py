"""The VBA project writers: what a module costs the file.

Every structure asserted here was measured against Access's own
`VBComponents.Add`, `DoCmd.Rename` and `DoCmd.DeleteObject`; the live gate
(`test_live_access_vba_gate.py`) hands the result back to Access and runs
it.  These checks are the offline half, and they exist because three of
the rules -- the storage folder's name, the object-id step, and removing
the folder on delete -- are invisible from the file and were wrong for a
while without anything noticing.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._vba import (
    CLASS_BASE,
    MODULETYPE_CLASS,
    STALE_VERSION,
    module_blocks,
    module_offset_at,
    records,
)
from pyopenvba.access_read import AccessError, AccessReader
from pyopenvba.vba import decompress

TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
ADDER = (
    "Option Compare Database\n"
    "\n"
    "Public Function AdderGo() As Variant\n"
    "    AdderGo = 4242\n"
    "End Function"
)


@pytest.fixture
def db(tmp_path: Path) -> AccessDatabase:
    """The shipped template, which holds exactly one module."""
    out = tmp_path / "blank.accdb"
    shutil.copyfile(TEMPLATE, out)
    return AccessDatabase(out)


def dir_stream(db: AccessDatabase) -> bytes:
    return db._vba_dir()[1]  # pyright: ignore[reportPrivateUsage]


def storage_rows(db: AccessDatabase) -> list[dict[str, object]]:
    return [row for _rid, row in db.table("MSysAccessStorage").rows_with_ids()]


def stream_named(db: AccessDatabase, name: str) -> bytes:
    """One storage stream by name.  `\x03DirData` and `PropData` exist
    under more than one container, so those are scoped to `Modules`."""
    modules_id = db._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    scoped = name in ("\x03DirData", "PropData")
    payload = next(
        r["Lv"]
        for r in storage_rows(db)
        if str(r["Name"]) == name and (not scoped or r["ParentId"] == modules_id)
    )
    assert isinstance(payload, bytes)
    return payload


def folders(db: AccessDatabase) -> list[str]:
    modules_id = db._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    rows = [
        r
        for r in storage_rows(db)
        if r["ParentId"] == modules_id and r["Type"] == 1
    ]
    return [str(r["Name"]) for r in sorted(rows, key=lambda r: int(str(r["Id"])))]


# --- reading ------------------------------------------------------------------


def test_the_template_reads_as_one_standard_module(db: AccessDatabase) -> None:
    assert [(m.name, m.kind) for m in db.modules()] == [("Module1", "module")]
    assert db.module("MODULE1").name == "Module1"  # VBA compares names case-insensitively
    assert db.module("Module1").source.startswith("Option Compare Database")
    with pytest.raises(AccessError, match="no module named"):
        db.module("Nothing")


# --- create -------------------------------------------------------------------


def test_a_created_module_reads_back_through_every_reader(db: AccessDatabase, tmp_path: Path) -> None:
    db.create_module("Adder", ADDER)
    out = tmp_path / "created.accdb"
    db.save(out)

    reopened = AccessDatabase(out)
    assert [m.name for m in reopened.modules()] == ["Module1", "Adder"]
    assert reopened.module("Adder").source == ADDER.replace("\n", "\r\n")
    # and through the standalone reader, which finds modules its own way
    names = {module.name for module in AccessReader(out).iter_vba_modules()}
    assert {"Module1", "Adder"} <= names


def test_create_writes_every_place_a_module_lives(db: AccessDatabase) -> None:
    db.create_module("Adder", ADDER)
    stream = dir_stream(db)

    assert [name for name, _row, _kind in module_blocks(stream)] == ["Module1", "Adder"]
    assert b"A\x00d\x00d\x00e\x00r\x00" in stream_named(db, "\x03DirData")
    assert b"Adder\x00A\x00d\x00d\x00e\x00r\x00" in stream_named(db, "PROJECTwm")
    project = stream_named(db, "PROJECT").decode("latin-1")
    assert "Module=Adder" in project
    assert "Adder=38, 38, 1786, 1030, " in project

    catalog = {e.name: e for e in db.catalog() if e.type == -32761}
    assert "Adder" in catalog
    nav = {str(r["Name"]) for _rid, r in db.table("MSysNavPaneObjectIDs").rows_with_ids()}
    assert "Adder" in nav
    groups = {
        int(str(r["ObjectID"]))
        for _rid, r in db.table("MSysNavPaneGroupToObjects").rows_with_ids()
        if r["GroupID"] == 8
    }
    assert catalog["Adder"].id in groups


def test_a_created_module_carries_no_pcode(db: AccessDatabase) -> None:
    """The stream is the compressed source alone and MODULEOFFSET is zero,
    which is what lets VBA compile it from source."""
    module = db.create_module("Adder", ADDER)
    stream = dir_stream(db)
    at = module_offset_at(stream, "Adder")

    assert int.from_bytes(stream[at : at + 4], "little") == 0
    source = decompress(stream_named(db, module.stream_name)).decode("latin-1")
    assert source.startswith('Attribute VB_Name = "Adder"')
    assert b"\xfe\xca" not in stream_named(db, module.stream_name)


def test_create_marks_the_compiled_cache_stale(db: AccessDatabase) -> None:
    before = stream_named(db, "_VBA_PROJECT")
    db.create_module("Adder", ADDER)
    after = stream_named(db, "_VBA_PROJECT")

    assert int.from_bytes(before[2:4], "little") != STALE_VERSION
    assert int.from_bytes(after[2:4], "little") == STALE_VERSION
    assert after[:2] == before[:2] and len(after) == len(before)


def test_a_class_module_differs_in_exactly_three_places(db: AccessDatabase) -> None:
    module = db.create_module("Widget", "Option Compare Database", kind="class")
    stream = dir_stream(db)

    assert module.kind == "class" and module.is_class
    assert MODULETYPE_CLASS in {ident for _at, ident, _size, _p in records(stream)}
    assert "Class=Widget" in stream_named(db, "PROJECT").decode("latin-1")
    source = decompress(stream_named(db, module.stream_name)).decode("latin-1")
    assert f'Attribute VB_Base = "{CLASS_BASE}"' in source


def test_storage_folders_and_object_ids_follow_access(db: AccessDatabase) -> None:
    """`Modules` holds four rows that are not folders, so its folders start
    at `4`; object ids step by four."""
    assert folders(db) == ["0"]
    first = db.create_module("One", "Option Compare Database")
    second = db.create_module("Two", "Option Compare Database")
    assert folders(db) == ["0", "4", "5"]

    ids = {e.name: e.id for e in db.catalog() if e.type == -32761}
    assert ids[second.name] - ids[first.name] == 4


def test_two_creates_take_different_stream_names(db: AccessDatabase) -> None:
    first = db.create_module("One", "Option Compare Database")
    second = db.create_module("Two", "Option Compare Database")

    assert first.stream_name != second.stream_name
    assert len(first.stream_name) == 28 and first.stream_name.isupper()


@pytest.mark.parametrize(
    ("name", "kind", "message"),
    [
        ("Module1", "module", "already exists"),
        ("", "module", "1 to 64 characters"),
        ("x" * 65, "module", "1 to 64 characters"),
        ("Fine", "document", "must be 'module' or 'class'"),
    ],
)
def test_create_refuses_what_access_would(
    db: AccessDatabase, name: str, kind: str, message: str
) -> None:
    with pytest.raises(AccessError, match=message):
        db.create_module(name, "Option Compare Database", kind=kind)


# --- source -------------------------------------------------------------------


def test_setting_source_keeps_the_attribute_block(db: AccessDatabase) -> None:
    db.create_module("Widget", "Option Compare Database", kind="class")
    db.set_module_source("Widget", "Option Compare Database\n\nPublic Sub Go()\nEnd Sub")

    module = db.module("Widget")
    assert module.kind == "class"
    assert module.source.endswith("End Sub")
    source = decompress(stream_named(db, module.stream_name)).decode("latin-1")
    assert f'Attribute VB_Base = "{CLASS_BASE}"' in source


def test_setting_source_drops_any_compiled_region(db: AccessDatabase) -> None:
    """The template's `Module1` arrives with p-code; replacing its source
    leaves the stream source-only with MODULEOFFSET back at zero."""
    before = db.module("Module1")
    assert b"\xfe\xca" in stream_named(db, before.stream_name)

    db.set_module_source("Module1", "Option Compare Database\n\nPublic Sub Go()\nEnd Sub")

    stream = dir_stream(db)
    at = module_offset_at(stream, "Module1")
    assert int.from_bytes(stream[at : at + 4], "little") == 0
    assert b"\xfe\xca" not in stream_named(db, before.stream_name)
    assert db.module("Module1").source.endswith("End Sub")


# --- rename -------------------------------------------------------------------


def test_rename_moves_the_name_everywhere(db: AccessDatabase) -> None:
    db.create_module("Adder", ADDER)
    db.rename_module("Adder", "Summer")

    assert [m.name for m in db.modules()] == ["Module1", "Summer"]
    assert db.module("Summer").source == ADDER.replace("\n", "\r\n")
    assert b"S\x00u\x00m\x00m\x00e\x00r\x00" in stream_named(db, "\x03DirData")
    assert b"Summer\x00" in stream_named(db, "PROJECTwm")
    project = stream_named(db, "PROJECT").decode("latin-1")
    assert "Module=Summer" in project and "Module=Adder" not in project
    assert "Summer=38, 38, 1786, 1030, " in project
    assert {e.name for e in db.catalog() if e.type == -32761} == {"Module1", "Summer"}
    nav = {str(r["Name"]) for _rid, r in db.table("MSysNavPaneObjectIDs").rows_with_ids()}
    assert "Summer" in nav and "Adder" not in nav
    source = decompress(stream_named(db, db.module("Summer").stream_name)).decode("latin-1")
    assert source.startswith('Attribute VB_Name = "Summer"')


def test_rename_refuses_a_name_already_taken(db: AccessDatabase) -> None:
    db.create_module("Adder", ADDER)
    with pytest.raises(AccessError, match="already exists"):
        db.rename_module("Adder", "Module1")


# --- delete -------------------------------------------------------------------


def test_delete_removes_every_structure(db: AccessDatabase) -> None:
    module = db.create_module("Adder", ADDER)
    db.delete_module("Adder")

    assert [m.name for m in db.modules()] == ["Module1"]
    assert folders(db) == ["0"]
    assert not [r for r in storage_rows(db) if str(r["Name"]) == module.stream_name]
    assert b"A\x00d\x00d\x00e\x00r\x00" not in stream_named(db, "\x03DirData")
    assert b"Adder" not in stream_named(db, "PROJECTwm")
    project = stream_named(db, "PROJECT").decode("latin-1")
    assert "Adder" not in project
    assert {e.name for e in db.catalog() if e.type == -32761} == {"Module1"}
    nav = {str(r["Name"]) for _rid, r in db.table("MSysNavPaneObjectIDs").rows_with_ids()}
    assert "Adder" not in nav


def test_delete_frees_the_folder_for_the_next_module(db: AccessDatabase) -> None:
    """Access links a module to its folder by position and reuses a freed
    name; a delete that left the folder behind would make the next create
    pick a name Access will not look under."""
    db.create_module("First", "Option Compare Database")
    db.create_module("Second", "Option Compare Database")
    assert folders(db) == ["0", "4", "5"]

    db.delete_module("First")
    assert folders(db) == ["0", "5"]

    db.create_module("Third", "Option Compare Database")
    assert folders(db) == ["0", "5", "4"]
    assert [m.name for m in db.modules()] == ["Module1", "Second", "Third"]


def test_delete_refuses_an_unknown_module(db: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="no module named"):
        db.delete_module("Nothing")
