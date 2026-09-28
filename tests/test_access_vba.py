"""The VBA project writers: what a module costs the file.

Every structure asserted here was measured against Access's own
`VBComponents.Add`, `DoCmd.Rename` and `DoCmd.DeleteObject`; the live gate
(`test_live_access_vba_gate.py`) hands the result back to Access and runs
it.  These checks are the offline half, and they exist because three of
the rules -- the storage folder's name, the object id, and removing the
folder on delete -- are invisible from the file.  The first two were
measured wrong once, through a harness whose own modules took folders and
ids while Access worked, which is why the fixtures they are checked
against now come from Access driven over COM alone.
"""

from __future__ import annotations

import shutil
from collections.abc import Callable
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, ColumnSpec, MacroAction
from pyopenvba.access._storage import access_order, dir_data_entries, name_hash, next_folder
from pyopenvba.access._vba import (
    CRLF,
    CRLF as VBA_CRLF,
    CLASS_BASE,
    MODULETYPE_CLASS,
    STALE_VERSION,
    add_to_project,
    add_to_project_documents,
    folder_entry,
    module_blocks,
    module_offset_at,
    records,
    remove_from_project,
    rename_project,
)
from pyopenvba.access_read import AccessError, AccessReader
from pyopenvba.exceptions import NoVBAProjectError, UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import decompress

TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)
FIXTURES = Path(__file__).parent / "live_access_test"
#: The template after Access took every storage container past folder 9,
#: built by `live_access_test/_folders_past_nine.ps1`.
PAST_NINE = FIXTURES / "folders_past_nine.accdb"
#: The template after Access added, deleted and added modules over COM
#: alone, built by `live_access_test/_module_allocation.ps1`.
MODULES_ADDED = FIXTURES / "modules_added.accdb"
#: A database Access made and never gave code, and the same database once
#: Access gave it its first module, built by `live_access_test/_first_project.ps1`.
NO_PROJECT = FIXTURES / "no_project.accdb"
FIRST_MODULE = FIXTURES / "first_module.accdb"
FIRST_MODULE_REOPENED = FIXTURES / "first_module_reopened.accdb"
FIRST_FORM = FIXTURES / "first_form.accdb"
FIRST_MACRO = FIXTURES / "first_macro.accdb"
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
    # and through the standalone reader
    assert [module.name for module in AccessReader(out).iter_vba_modules()] == ["Module1", "Adder"]
    assert AccessReader(out).read_vba_module("Adder") == ADDER.replace("\n", "\r\n")


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


def test_what_access_adds_on_its_next_open_is_left_to_it(db: AccessDatabase) -> None:
    """Access files a new module in `Modules/PropData` and under a
    navigation-pane group the next time it opens the database, not when it
    makes the module: `first_module.accdb` has neither for its module, and
    opening it once more in Access, and nothing else, added both.  A
    delete leaves the group row, as Access's does."""
    for name, filed in ((FIRST_MODULE, False), (FIRST_MODULE_REOPENED, True)):
        access = AccessDatabase(name)
        module_id = next(e.id for e in access.catalog() if e.name == "Module1")
        modules_id = access._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
        lists = [r["Lv"] for r in storage_rows(access) if r["ParentId"] == modules_id and str(r["Name"]) == "PropData"]
        assert lists == ([bytes(4) + folder_entry("0")] if filed else []), name.name
        assert (module_id in group_members(access)) is filed, name.name

    before = stream_named(db, "PropData")
    db.create_module("Adder", ADDER)
    adder = next(e.id for e in db.catalog() if e.name == "Adder")
    assert stream_named(db, "PropData") == before
    assert adder not in group_members(db)

    module1 = next(e.id for e in db.catalog() if e.name == "Module1")
    assert module1 in group_members(db)
    db.delete_module("Module1")
    assert module1 in group_members(db)


def group_members(db: AccessDatabase) -> set[int]:
    return {int(str(r["ObjectID"])) for _rid, r in db.table("MSysNavPaneGroupToObjects").rows_with_ids()}


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


def module_places(db: AccessDatabase) -> dict[str, tuple[str, int]]:
    """Each module's storage folder, as `\\x03DirData` names it, and its
    object id."""
    folder_of = dict(dir_data_entries(stream_named(db, "\x03DirData")))
    return {e.name: (folder_of[e.name], e.id) for e in db.catalog() if e.type == -32761}


def test_storage_folders_and_object_ids_follow_access(db: AccessDatabase) -> None:
    """Access, driven over COM alone, gave the template's new modules the
    lowest free folders from `0` and the next object ids, and a module
    added after a middle one was deleted took the freed folder but a new
    id (`live_access_test/modules_added.accdb`).

    The rule this replaces started `Modules` at `4` and stepped ids by
    four: it was measured through pyvbaharness, whose three injected
    modules held the lowest folders and the next ids meanwhile."""
    db.create_module("One", "Option Compare Database")
    db.create_module("Two", "Option Compare Database")
    db.create_module("Three", "Option Compare Database")
    db.delete_module("Two")
    db.create_module("Four", "Option Compare Database")

    assert module_places(db) == module_places(AccessDatabase(MODULES_ADDED))


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



def test_dirdata_is_the_one_access_wrote(db: AccessDatabase) -> None:
    """The four bytes an entry ends with are the folder the module's
    stream lives in, not a terminator, and the entries come in the order
    Access writes them: replaying the sessions that built
    `modules_added.accdb` gives its `\\x03DirData` byte for byte, Four in
    the folder Two left and filed in front of Three."""
    for name in ("One", "Two", "Three"):
        db.create_module(name, "Option Compare Database")
    db.delete_module("Two")
    db.create_module("Four", "Option Compare Database")

    assert dir_data_entries(stream_named(db, "\x03DirData")) == [
        ("Module1", "0"),
        ("One", "1"),
        ("Four", "2"),
        ("Three", "3"),
    ]
    assert stream_named(db, "\x03DirData") == stream_named(AccessDatabase(MODULES_ADDED), "\x03DirData")


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
    assert folders(db) == ["0", "1", "2"]

    db.delete_module("First")
    assert folders(db) == ["0", "2"]

    db.create_module("Third", "Option Compare Database")
    assert folders(db) == ["0", "2", "1"]
    assert [m.name for m in db.modules()] == ["Module1", "Second", "Third"]


def test_folders_past_nine_are_named_in_decimal(db: AccessDatabase, tmp_path: Path) -> None:
    """The eleventh folder is `10`, not the character after `9`, which made
    `\\x03DirData` refuse the name (GitHub issue #34)."""
    names = [f"Extra{i}" for i in range(11)]
    for name in names:
        db.create_module(name, "Option Compare Database")

    assert folders(db) == [str(number) for number in range(12)]
    assert dir_data_entries(stream_named(db, "\x03DirData"))[-2:] == [("Extra9", "10"), ("Extra10", "11")]

    db.delete_module("Extra9")
    out = tmp_path / "written.accdb"
    db.save(out)
    assert [m.name for m in AccessDatabase(out).modules()] == ["Module1", *names[:9], "Extra10"]


def test_a_delete_takes_access_s_two_digit_folder_entry(tmp_path: Path) -> None:
    """The folder list's entry grows with the name, as Access's does, so a
    module in folder `10` is found in Access's own list and taken out."""
    copy = tmp_path / PAST_NINE.name
    shutil.copyfile(PAST_NINE, copy)
    db = AccessDatabase(copy)
    ten = next(name for name, folder in dir_data_entries(stream_named(db, "\x03DirData")) if folder == "10")
    assert folder_entry("10") in stream_named(db, "PropData")

    db.delete_module(ten)
    assert folder_entry("10") not in stream_named(db, "PropData")
    assert folder_entry("11") in stream_named(db, "PropData")


def access_containers() -> dict[str, tuple[set[str], list[tuple[str, str]], bytes | None]]:
    """Each container of the database Access filled past folder 9: its
    folders, its `\\x03DirData` entries and its `PropData`, if any."""
    return containers(AccessDatabase(PAST_NINE))


def containers(written: AccessDatabase) -> dict[str, tuple[set[str], list[tuple[str, str]], bytes | None]]:
    """Each object container of a database: its folders, its `\\x03DirData`
    entries and its `PropData`, if any."""
    rows = storage_rows(written)
    root = next(int(str(r["Id"])) for r in rows if str(r["Name"]) == "MSysAccessStorage_ROOT")
    out: dict[str, tuple[set[str], list[tuple[str, str]], bytes | None]] = {}
    for container in ("Modules", "Forms", "Reports", "Scripts"):
        held = next(int(str(r["Id"])) for r in rows if r["ParentId"] == root and str(r["Name"]) == container)
        children = [r for r in rows if r["ParentId"] == held]
        streams = {str(r["Name"]): r["Lv"] for r in children if r["Type"] != 1}
        listing = streams["\x03DirData"]
        property_list = streams.get("PropData")
        assert isinstance(listing, bytes)
        out[container] = (
            {str(r["Name"]) for r in children if r["Type"] == 1},
            dir_data_entries(listing),
            property_list if isinstance(property_list, bytes) else None,
        )
    return out


def test_access_numbers_every_container_s_folders_in_decimal() -> None:
    """Access took each container past 9 in the fixture: `10` follows `9`
    everywhere, `\\x03DirData` names the same folders, and the name the
    library gives the last object is the one Access gave it."""
    for container, (held, entries, _property_list) in access_containers().items():
        assert "10" in held, container
        assert {folder for _name, folder in entries} == held, container
        last = max(held, key=int)
        assert next_folder(held - {last}) == last, container


def test_the_folder_list_holds_access_s_own_entries() -> None:
    """Access's `Modules/PropData` lists every folder, one-digit and
    two-digit, in the bytes `folder_entry` writes and nothing else."""
    held, _entries, property_list = access_containers()["Modules"]
    assert property_list is not None
    for folder in held:
        assert folder_entry(folder) in property_list, folder
    assert len(property_list) == 4 + sum(len(folder_entry(folder)) for folder in held)


# --- the order of a container's lists ------------------------------------------
# Access keeps `\x03DirData` and `PropData` in a hash map's order, not the
# order objects were made in (`_storage.access_order`, read from
# MSACCESS.EXE).  Each list below is one Access wrote.


def folder_list(payload: bytes) -> list[str]:
    """The folders a container's `PropData` lists, in order."""
    out: list[str] = []
    at = 4
    while at < len(payload):
        body = payload[at + 2 : at + 2 + payload[at + 1]]
        out.append(body[1 : 1 + body[0]].decode("utf-16-le"))
        at += 2 + payload[at + 1]
    return out


def test_each_container_lists_its_objects_in_access_s_order(db: AccessDatabase) -> None:
    """Making the objects `folders_past_nine.accdb` holds, in the order
    Access made them, lists every container as Access did: `Mod6` in front
    of `Module1`, `Mod11` behind it, `Macro10` between `Macro1` and
    `Macro2`, and the forms and reports in the order they were made."""
    for number in range(2, 14):
        db.create_module(f"Mod{number}", "Option Compare Database")
    for number in range(1, 12):
        db.create_form(f"Form{number}")
    for number in range(1, 12):
        db.create_report(f"Report{number}")
    for number in range(1, 12):
        db.create_macro(f"Macro{number}", [MacroAction("Beep")])

    ours = containers(db)
    for container, (_held, entries, _property_list) in access_containers().items():
        assert ours[container][1] == entries, container
    assert [name for name, _folder in ours["Scripts"][1]][:4] == ["Macro1", "Macro10", "Macro2", "Macro11"]


#: What Access wrote for one change to a copy of `folders_past_nine.accdb`,
#: by `live_access_test/_container_order.ps1`: the container's
#: `\x03DirData` names, and for `Modules` its `PropData` folders.  The
#: module was renamed in the VBE, which leaves `PropData` as it was.
ACCESS_CHANGES: dict[str, tuple[str, list[str], list[str] | None]] = {
    "module deleted": (
        "Modules",
        ["Mod6", "Module1", "Mod11", "Mod2", "Mod12", "Mod3", "Mod13", "Mod4", "Mod5", "Mod8", "Mod9", "Mod10"],
        ["10", "0", "11", "1", "12", "2", "3", "4", "5", "7", "8", "9"],
    ),
    "module renamed": (
        "Modules",
        ["Mod6", "Module1", "Mod11", "Mod2", "Mod12", "Renamed3", "Mod13", "Mod4", "Mod5", "Mod7", "Mod8", "Mod9", "Mod10"],
        ["10", "0", "11", "1", "12", "2", "3", "4", "5", "6", "7", "8", "9"],
    ),
    "form deleted": (
        "Forms",
        ["Form1", "Form2", "Form3", "Form4", "Form6", "Form7", "Form8", "Form9", "Form10", "Form11"],
        None,
    ),
    "report renamed": (
        "Reports",
        ["Report1", "Report3", "Report4", "Report5", "Report6", "Report7", "Report8", "Report9", "Report10",
         "Report11", "Summary"],
        None,
    ),
}
CHANGES: dict[str, Callable[[AccessDatabase], None]] = {
    "module deleted": lambda db: db.delete_module("Mod7"),
    "module renamed": lambda db: db.rename_module("Mod3", "Renamed3"),
    "form deleted": lambda db: db.delete_form("Form5"),
    "report renamed": lambda db: db.rename_report("Report2", "Summary"),
}


@pytest.mark.parametrize("change", list(ACCESS_CHANGES))
def test_a_delete_or_rename_lists_the_container_as_access_does(change: str, tmp_path: Path) -> None:
    """Access loads a list afresh, erases the name and, for a rename,
    inserts the new one, so entries the change never named move too:
    `Mod6` and `Module1` swap places.  A form or report Access has never
    reopened has no `PropData`; Access would write one on opening the
    database, which is not the change's doing."""
    copy = tmp_path / PAST_NINE.name
    shutil.copyfile(PAST_NINE, copy)
    db = AccessDatabase(copy)
    CHANGES[change](db)

    container, names, folders = ACCESS_CHANGES[change]
    _held, entries, property_list = containers(db)[container]
    assert [name for name, _folder in entries] == names
    if folders is not None:
        assert property_list is not None and folder_list(property_list) == folders


SINGLES = list("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")


def test_access_order_follows_the_probes() -> None:
    """Orders Access wrote for macros named `0` to `9` and `A` to `Z`, made
    in one session forwards and backwards, then changed one session at a
    time.  Characters 32 apart hash alike, so `P` and `0` share a bucket,
    the newer in front; a reopened list keeps or reverses a pair by
    whether the map grew before or after loading it; a rename moves the
    entry; and a name past ASCII is hashed in cp1252."""
    forward = access_order([], add=tuple(SINGLES))
    assert forward == list("P0Q1R2S3T4U5V6W7X8Y9ABCDEFGHIJKLMNOZ")
    assert access_order([], add=tuple(reversed(SINGLES))) == list("Z9Y8X7W6V5U4T3S2R1Q0PONMLKJIHGFEDCBA")
    assert access_order(forward, add=("_",)) == list("P0Q1R2S34T5U6V7W8X9YABCDEFGHIJKLMNOZ_")

    steps: list[tuple[tuple[str, ...], tuple[str, ...], list[str]]] = [
        (("5",), (), list("P0Q1R2S34TU6V7W8X9YABCDEFGHIJKLMNOZ")),
        (("C",), ("CC",), [*"P0Q1R2S3T4UV6W7X8Y9ABDEFGHIJKLMNOZ", "CC"]),
        (("R",), ("r2",), [*"P0Q12S34TU6V7W8X9YABDEFGHIJKLMNOZ", "CC", "r2"]),
        (
            (),
            ("é", "š", "Café", "Œuvre", "Žaba", "Ärger", "Straße"),
            [*"P0Q12S3T4UV6W7", "Ärger", *"X8Y9ABDEFGH", "é", "I", "š", "J", "Straße",
             "Œuvre", *"KLMNOZ", "CC", "r2", "Café", "Žaba"],
        ),
        (
            (),
            ('a"b', "x y", "q~q", "o'o"),
            ["P", "0", "q~q", "Q", "1", "2", "S", "3", "4", "T", "U", "6", "V", "7", "W", "8", "X",
             "Ärger", "9", "Y", "A", 'a"b', "B", "D", "E", "F", "G", "H", "I", "é", "J", "š",
             "K", "Œuvre", "Straße", "L", "M", "N", "o'o", "O", "Z", "CC", "r2", "Café",
             "Žaba", "x y"],
        ),
        (
            ("CC",),
            (),
            ["P", "0", "q~q", "Q", "1", "2", "S", "3", "T", "4", "U", "V", "6", "W", "7", "Ärger", "X",
             "8", "Y", "9", "A", "B", 'a"b', "D", "E", "F", "G", "H", "é", "I", "š", "J",
             "Straße", "Œuvre", "K", "L", "M", "N", "O", "o'o", "Z", "r2", "Café", "Žaba",
             "x y"],
        ),
    ]
    stored = forward
    for removed, added, written in steps:
        stored = access_order(stored, remove=removed, add=added)
        assert stored == written, (removed, added)


def test_the_name_hash_counts_five_bits_a_character() -> None:
    """Each character that counts adds its low five bits, so case and the
    32 between `0` and `P` disappear; quotes and `~` are skipped, a space
    counts as nothing but still counts, and a leading `.` is dropped."""
    assert name_hash("0") == name_hash("P") == 17
    assert name_hash("Macro1") == name_hash("MACRO1") == name_hash("macro1")
    assert name_hash('a"b') == name_hash("a'b") == name_hash("a~b") == name_hash("ab")
    assert name_hash("x y") != name_hash("xy")
    assert name_hash(".Hidden") == name_hash("Hidden")
    assert name_hash("š") == name_hash("Š") != name_hash("S")


def test_delete_refuses_an_unknown_module(db: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="no module named"):
        db.delete_module("Nothing")


def vba_data_count(db: AccessDatabase) -> int:
    return int.from_bytes(stream_named(db, "AcessVBAData")[8:12], "little")


def test_acess_vba_data_counts_the_modules(db: AccessDatabase) -> None:
    """Access keeps a count of the project's modules in `AcessVBAData`:
    4, 1, 13 and 2 in the fixtures it wrote, a form's module counted.
    The writers here left it as they found it."""
    for name in ("modules_added.accdb", "first_module.accdb", "folders_past_nine.accdb", "form_with_code.accdb"):
        written = AccessDatabase(FIXTURES / name)
        assert vba_data_count(written) == len(written.modules()), name

    db.create_module("One", "Option Compare Database")
    db.create_module("Two", "Option Compare Database")
    assert vba_data_count(db) == 3
    db.delete_module("One")
    assert vba_data_count(db) == 2
    db.create_form("Summary")
    db.set_design_code("Summary", "Option Compare Database")
    assert vba_data_count(db) == 3
    db.delete_form("Summary")
    assert vba_data_count(db) == 2


# --- a database that has never held code --------------------------------------


@pytest.fixture
def bare(tmp_path: Path) -> AccessDatabase:
    """Access's database with no project, under the name of the file its
    first module went into, since a project is named after its file."""
    out = tmp_path / FIRST_MODULE.name
    shutil.copyfile(NO_PROJECT, out)
    return AccessDatabase(out)


def test_a_database_with_no_project_lists_nothing(bare: AccessDatabase, tmp_path: Path) -> None:
    assert not bare.has_vba_project()
    assert bare.modules() == []
    assert bare.module_names() == []
    assert bare.module_streams() == []
    assert bare.project_streams() == []
    assert bare.references() == []
    assert not bare.vba_is_protected()
    assert bare.pull_modules(tmp_path / "pulled") == []


def test_what_needs_the_project_refuses_naming_the_file(bare: AccessDatabase, tmp_path: Path) -> None:
    source = tmp_path / "source"
    source.mkdir()
    (source / "Module1.bas").write_text("Sub A()\nEnd Sub\n", encoding="utf-8")
    calls: list[Callable[[], object]] = [
        lambda: bare.vba_project(),
        lambda: bare.dir_stream(),
        lambda: bare.module("Module1"),
        lambda: bare.get_module("Module1"),
        lambda: bare.add_module("Module1"),
        lambda: bare.set_module("Module1", "Sub A()\r\nEnd Sub\r\n"),
        lambda: bare.rename_module("Module1", "Other"),
        lambda: bare.delete_module("Module1"),
        lambda: bare.push_modules(source),
        lambda: bare.add_reference("Excel"),
        lambda: bare.remove_reference("Excel"),
        lambda: bare.set_design_code("Summary", "Option Compare Database"),
    ]
    for call in calls:
        with pytest.raises(NoVBAProjectError, match="Access") as raised:
            call()
        assert repr(FIRST_MODULE.name) in str(raised.value)


def storage_tree(db: AccessDatabase) -> dict[str, bytes | None]:
    """Every storage row under its path, the compiled caches and the
    module streams' random names left out."""
    rows = storage_rows(db)
    by_id = {int(str(r["Id"])): r for r in rows}
    stream_names = {m.stream_name for m in db.modules()}
    out: dict[str, bytes | None] = {}
    for row in rows:
        parts: list[str] = []
        at: dict[str, object] = row
        while True:
            name = str(at["Name"])
            parts.append("<module>" if name in stream_names else name)
            parent = int(str(at["ParentId"]))
            if parent not in by_id or parent == int(str(at["Id"])):
                break
            at = by_id[parent]
        value = row.get("Lv")
        key = "/".join(reversed(parts))
        if not key.split("/")[-1].startswith("__SRP_"):
            out[key] = value if isinstance(value, bytes) else None
    return out


def test_a_first_project_is_the_one_access_makes(bare: AccessDatabase) -> None:
    """Access gave the database its first module; the library does the same
    in two steps.  What differs is what differs between two projects Access
    makes -- the ID, the protection records, PROJECTVERSION, the cookies,
    the stream names -- and what the library never writes: the compiled
    caches, and a module stream holding compiled code."""
    bare.add_vba_project()
    bare.add_module("Module1", "Public Function Module1Go() As Variant\n    Module1Go = 1\nEnd Function")
    access = AccessDatabase(FIRST_MODULE)

    ours, theirs = storage_tree(bare), storage_tree(access)
    assert set(ours) == set(theirs)
    assert group_members(bare) == group_members(access)
    for path in (
        "MSysAccessStorage_ROOT/PropData",
        "MSysAccessStorage_ROOT/VBA/AcessVBAData",
        "MSysAccessStorage_ROOT/VBA/VBAProject/PROJECTwm",
        "MSysAccessStorage_ROOT/Modules/\x03DirData",
        "MSysAccessStorage_ROOT/Modules/0/PropData",
    ):
        assert ours[path] == theirs[path], path

    def catalog(db: AccessDatabase) -> list[tuple[str, int, int, bytes | None]]:
        return sorted((e.name, e.id, e.type, e.owner) for e in db.catalog())

    assert catalog(bare) == catalog(access)
    # MSysDb gains HasOfflineLists and ProjVer as the project is made.
    assert database_properties(bare) == database_properties(access)

    def project_lines(db: AccessDatabase) -> list[str]:
        text = stream_named(db, "PROJECT").decode("latin-1")
        kept = ("ID=", "CMG=", "DPB=", "GC=", "Module1=")
        return [line for line in text.split("\r\n") if not line.startswith(kept)]

    assert project_lines(bare) == project_lines(access)
    assert bare.references() == access.references()
    ours_dir, theirs_dir = dir_stream(bare), dir_stream(access)
    per_project = {0x0009, 0x0013, 0x001A, 0x0032, 0x0031, 0x002C}

    def records_kept(stream: bytes) -> list[tuple[int, bytes]]:
        return [(ident, payload) for _at, ident, _size, payload in records(stream) if ident not in per_project]

    assert records_kept(ours_dir) == records_kept(theirs_dir)


def database_properties(db: AccessDatabase) -> bytes:
    """MSysDb's property blob, as stored."""
    blob = next(r["LvProp"] for r in db.table("MSysObjects").rows() if str(r["Name"]) == "MSysDb")
    assert isinstance(blob, bytes)
    return blob


def decrypt_project_data(hex_text: str) -> tuple[int, int, bytes]:
    """``(version, key, data)`` from a ``CMG``, ``DPB`` or ``GC`` value,
    decrypted as [MS-OVBA] 2.4.3.3 has it, independently of the writer."""
    data = bytes.fromhex(hex_text)
    seed, version_enc, key_enc = data[0], data[1], data[2]
    plain_prev, enc1, enc2 = seed ^ key_enc, key_enc, version_enc
    plain = bytearray()
    for byte_enc in data[3:]:
        byte = byte_enc ^ ((enc2 + plain_prev) & 0xFF)
        plain.append(byte)
        enc2, enc1, plain_prev = enc1, byte_enc, byte
    body = bytes(plain[(seed & 6) // 2 :])
    length = int.from_bytes(body[:4], "little")
    assert len(body) == 4 + length
    return seed ^ version_enc, seed ^ key_enc, body[4:]


def protection_records(db: AccessDatabase) -> dict[str, tuple[int, int, bytes]]:
    lines = stream_named(db, "PROJECT").decode("latin-1").split("\r\n")
    fields = dict(line.split("=", 1) for line in lines if line[:4] in ("ID=\"", "CMG=", "DPB=", "GC=\""))
    key = sum(fields["ID"].strip('"').encode("ascii")) & 0xFF
    records = {name: decrypt_project_data(fields[name].strip('"')) for name in ("CMG", "DPB", "GC")}
    assert {record[1] for record in records.values()} == {key}
    return records


def test_a_new_project_s_protection_records_are_keyed_to_its_id(bare: AccessDatabase) -> None:
    """Every project Access wrote keys `CMG`, `DPB` and `GC` to the byte
    sum of its `ID`, and they say not protected, no password, visible.  A
    new ID with the template's records made Access take the project for a
    protected one."""
    unprotected = {"CMG": (2, bytes(4)), "DPB": (2, bytes(1)), "GC": (2, b"\xff")}
    for name in (FIRST_MODULE, TEMPLATE, FIXTURES / "two_modules_one_page.accdb"):
        records = protection_records(AccessDatabase(name))
        assert {k: (v[0], v[2]) for k, v in records.items()} == unprotected, name.name

    bare.add_vba_project()
    records = protection_records(bare)
    assert {k: (v[0], v[2]) for k, v in records.items()} == unprotected


def project_part(db: AccessDatabase) -> dict[str, object]:
    """What a project made with a first object comes to, per-project values
    and compiled caches aside."""
    tree = storage_tree(db)
    kept = ("ID=", "CMG=", "DPB=", "GC=")
    text = stream_named(db, "PROJECT").decode("latin-1")
    per_project = {0x0009, 0x0013}
    return {
        "rows": sorted(path for path in tree if "/VBA" in path),
        "vba data": tree["MSysAccessStorage_ROOT/VBA/AcessVBAData"],
        "root properties": tree["MSysAccessStorage_ROOT/PropData"],
        "PROJECTwm": tree["MSysAccessStorage_ROOT/VBA/VBAProject/PROJECTwm"],
        "PROJECT": [line for line in text.split("\r\n") if not line.startswith(kept)],
        "dir": [(i, p) for _at, i, _size, p in records(dir_stream(db)) if i not in per_project],
    }


def make_form(db: AccessDatabase) -> object:
    return db.create_form("Form1")


def make_macro(db: AccessDatabase) -> object:
    return db.create_macro("Macro1", [MacroAction("Beep")])


@pytest.mark.parametrize(
    ("made_by_access", "make"),
    [(FIRST_FORM, make_form), (FIRST_MACRO, make_macro)],
    ids=["form", "macro"],
)
def test_a_first_form_or_macro_brings_the_project_access_makes(
    tmp_path: Path, made_by_access: Path, make: Callable[[AccessDatabase], object]
) -> None:
    """Access makes the VBA project with a database's first form, report or
    macro too, a project whose `PROJECT` has no `[Workspace]` section, and
    files the object in its container's `\\x03DirData` alone."""
    copy = tmp_path / made_by_access.name
    shutil.copyfile(NO_PROJECT, copy)
    db = AccessDatabase(copy)
    make(db)
    access = AccessDatabase(made_by_access)

    assert project_part(db) == project_part(access)
    assert "[Workspace]" not in stream_named(db, "PROJECT").decode("latin-1")
    container_rows = {
        path: value for path, value in storage_tree(access).items() if path.endswith(("\x03DirData", "/PropData"))
    }
    ours = storage_tree(db)
    for path, value in container_rows.items():
        assert ours.get(path) == value, path
    objects = {e.name: (e.id, e.type, e.owner) for e in access.catalog() if e.type in (-32768, -32766)}
    assert {e.name: (e.id, e.type, e.owner) for e in db.catalog() if e.type in (-32768, -32766)} == objects
    if made_by_access == FIRST_MACRO:
        # A first form also brings the database its theme, which is not written here.
        assert database_properties(db) == database_properties(access)


def test_form_code_opens_the_workspace_section_as_access_does() -> None:
    """Measured: code put behind Form1 in a project Access made with the
    form, which had no `[Workspace]` section, gave `PROJECT` a `DocClass`
    line under `ID` and a section of its own after a blank line."""
    before = (
        'ID="{17EAFA9D-E599-4840-A347-155B648B40F1}"\r\nName="form_first"\r\nHelpContextID="0"\r\n'
        'VersionCompatible32="393222000"\r\n\r\n[Host Extender Info]\r\n'
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n"
    )
    after = (
        'ID="{17EAFA9D-E599-4840-A347-155B648B40F1}"\r\nDocClass=Form_Form1/&H00000000\r\n'
        'Name="form_first"\r\nHelpContextID="0"\r\nVersionCompatible32="393222000"\r\n\r\n'
        "[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n"
        "\r\n[Workspace]\r\nForm_Form1=0, 0, 0, 0, C\r\n"
    )
    assert add_to_project_documents(before, "Form_Form1") == after


def test_a_container_s_last_object_takes_its_lists_with_it(tmp_path: Path) -> None:
    """Access removes a container's `\\x03DirData` and `PropData` rows
    with its last object rather than leave them empty."""
    copy = tmp_path / FIRST_MODULE_REOPENED.name
    shutil.copyfile(FIRST_MODULE_REOPENED, copy)
    db = AccessDatabase(copy)
    modules_id = db._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    db.delete_module("Module1")
    assert not [r for r in storage_rows(db) if r["ParentId"] == modules_id and r["Type"] != 1]

    db.create_form("Only")
    db.delete_form("Only")
    forms = next(int(str(r["Id"])) for r in storage_rows(db) if str(r["Name"]) == "Forms")
    assert not [r for r in storage_rows(db) if r["ParentId"] == forms]


def test_a_project_is_named_after_its_file_up_to_the_first_dot(tmp_path: Path) -> None:
    """As Access names one: `a-b.c.accdb` gives `a-b`, spaces and all
    kept; a database opened from bytes has no file to name it after."""
    named = tmp_path / "My db.v2.accdb"
    shutil.copyfile(NO_PROJECT, named)
    db = AccessDatabase(named)
    db.add_vba_project()
    assert 'Name="My db"' in stream_named(db, "PROJECT").decode("latin-1")
    assert b"\x04\x00\x05\x00\x00\x00My db" in dir_stream(db)

    unnamed = AccessDatabase(NO_PROJECT.read_bytes())
    unnamed.add_vba_project()
    assert 'Name="Database"' in stream_named(unnamed, "PROJECT").decode("latin-1")


def test_a_first_project_reads_back_and_takes_a_second(bare: AccessDatabase, tmp_path: Path) -> None:
    project = bare.add_vba_project()
    assert bare.has_vba_project()
    assert project.modules == []
    assert vba_data_count(bare) == 0
    bare.add_module("Adder", ADDER)
    bare.add_module("Other", "Option Compare Database")
    out = tmp_path / "written.accdb"
    bare.save(out)

    reopened = AccessDatabase(out)
    assert [m.name for m in reopened.modules()] == ["Adder", "Other"]
    assert reopened.module("Adder").source == ADDER.replace("\n", "\r\n")
    assert module_places(reopened) == {"Adder": ("0", -2147483638), "Other": ("1", -2147483637)}


def test_add_vba_project_refuses_what_it_cannot_make(db: AccessDatabase, tmp_path: Path) -> None:
    with pytest.raises(VBAProjectError, match="already has a VBA project"):
        db.add_vba_project()
    dao_made = tmp_path / "dao.accdb"
    shutil.copyfile(FIXTURES / "complex_columns.accdb", dao_made)
    with pytest.raises(UnsupportedFormatError, match="none of Access's own objects"):
        AccessDatabase(dao_made).add_vba_project()


# --- the project's references -------------------------------------------------


def test_the_template_points_at_the_two_libraries_access_ships(db: AccessDatabase) -> None:
    """VBA itself and Access are not in the file, so they are not here."""
    assert [(r.name, r.version) for r in db.references()] == [("stdole", (2, 0)), ("DAO", (12, 0))]


def test_a_libid_reads_out_in_its_parts(db: AccessDatabase) -> None:
    """The version is written in hex, so DAO 12.0 is stored as `c.0`."""
    dao = next(r for r in db.references() if r.name == "DAO")

    assert dao.guid == "{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}"
    assert dao.version == (12, 0)
    assert "c.0" in dao.libid
    assert dao.path.lower().endswith(".dll")
    assert "Access database engine" in dao.description


def test_a_reference_can_be_added(db: AccessDatabase, tmp_path: Path) -> None:
    made = db.add_reference(
        "Scripting", "420B2830-E718-11CF-893D-00A0C9054228", 1, 0,
        path="C:/Windows/System32/scrrun.dll", description="Microsoft Scripting Runtime",
    )

    assert made.name == "Scripting" and made.version == (1, 0)
    assert made.guid == "{420B2830-E718-11CF-893D-00A0C9054228}"
    out = tmp_path / "written.accdb"
    db.save(out)
    assert [r.name for r in AccessDatabase(out).references()] == ["stdole", "DAO", "Scripting"]


def test_a_reference_can_be_dropped(db: AccessDatabase) -> None:
    db.add_reference("Scripting", "420B2830-E718-11CF-893D-00A0C9054228", 1, 0)
    db.drop_reference("Scripting")

    assert [r.name for r in db.references()] == ["stdole", "DAO"]


def test_a_reference_already_there_is_refused(db: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="already references"):
        db.add_reference("DAO", "420B2830-E718-11CF-893D-00A0C9054228")


def test_dropping_one_that_is_not_there_is_refused(db: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="no reference named"):
        db.drop_reference("Nothing")


def test_adding_a_reference_marks_the_cache_stale(db: AccessDatabase) -> None:
    """VBA has to recompile before it will resolve the new names."""
    db.add_reference("Scripting", "420B2830-E718-11CF-893D-00A0C9054228", 1, 0)

    blob = stream_named(db, "_VBA_PROJECT")
    assert int.from_bytes(blob[2:4], "little") == STALE_VERSION


# --- a password-protected project ---------------------------------------------


def protect(db: AccessDatabase) -> None:
    """Give the project a DPB long enough to read as password-bearing,
    which is the same rule the other hosts use."""
    storage = db.table("MSysAccessStorage")
    for rid, row in list(storage.rows_with_ids()):
        payload = row.get("Lv")
        if str(row["Name"]) == "PROJECT" and isinstance(payload, bytes):
            text = payload.decode("latin-1")
            before = [line for line in text.split(CRLF) if line.startswith("DPB=")][0]
            storage.update_row(
                rid, {"Lv": text.replace(before, 'DPB="' + "AB" * 40 + '"').encode("latin-1")}
            )
            return
    raise AssertionError("no PROJECT stream to protect")


def test_an_ordinary_project_is_not_protected(db: AccessDatabase) -> None:
    assert not db.vba_is_protected()


def test_a_project_with_a_password_reads_as_protected(db: AccessDatabase) -> None:
    protect(db)
    assert db.vba_is_protected()


def test_saving_a_vba_change_into_a_protected_project_is_refused(
    db: AccessDatabase, tmp_path: Path
) -> None:
    protect(db)
    db.create_module("Adder", ADDER)

    with pytest.raises(AccessError, match="password-protected"):
        db.save(tmp_path / "refused.accdb")


def test_the_refusal_can_be_opted_out_of(db: AccessDatabase, tmp_path: Path) -> None:
    """The protection bytes are kept, so the result still wants the
    original password."""
    protect(db)
    db.create_module("Adder", ADDER)
    out = tmp_path / "anyway.accdb"

    db.save(out, allow_protected=True)

    reopened = AccessDatabase(out)
    assert reopened.vba_is_protected()
    assert "Adder" in {module.name for module in reopened.modules()}


def test_a_change_that_is_not_vba_saves_freely(db: AccessDatabase, tmp_path: Path) -> None:
    """The guard is about the VBA project, not the database."""
    protect(db)
    db.create_table("Notes", [ColumnSpec("Id", "Long")])

    db.save(tmp_path / "tables.accdb")  # no refusal


# --- the PROJECT stream's module block ----------------------------------------
# Every module is named on one line: `Module=`, `Class=`, or, for the code
# behind a form or report, `DocClass=<name>/<flags>`.  The editors reached
# the first two and left DocClass stale (GitHub issue #21).

_PROJECT = VBA_CRLF.join(
    [
        'ID="{X}"',
        "Module=Module1",
        "Class=Basket",
        "DocClass=Form_Calculator/&H00000000",
        'Name="Database"',
        "",
        "[Workspace]",
        "Module1=38, 38, 1786, 1030, ",
        "Form_Calculator=0, 0, 0, 0, C",
        "",
    ]
)


def test_renaming_reaches_a_document_module() -> None:
    renamed = rename_project(_PROJECT, "Form_Calculator", "Form_Invoice").split(VBA_CRLF)

    assert "DocClass=Form_Invoice/&H00000000" in renamed
    assert "Form_Invoice=0, 0, 0, 0, C" in renamed
    assert not [line for line in renamed if "Form_Calculator" in line]


def test_renaming_a_class_does_not_catch_a_document_module() -> None:
    """`^Class=` and `DocClass=` share a suffix, so the anchor is what
    keeps them apart."""
    renamed = rename_project(_PROJECT, "Basket", "Bag").split(VBA_CRLF)

    assert "Class=Bag" in renamed
    assert "DocClass=Form_Calculator/&H00000000" in renamed


def test_removing_a_module_drops_its_document_line() -> None:
    left = remove_from_project(_PROJECT, "Form_Calculator").split(VBA_CRLF)

    assert not [line for line in left if "Form_Calculator" in line]
    assert "Module=Module1" in left and "Class=Basket" in left


def test_a_project_with_no_module_block_still_takes_one() -> None:
    """Delete a project's last module and there is no line to sit after.
    Access opens the block right below the ID line."""
    empty = VBA_CRLF.join(['ID="{X}"', 'Name="Database"', "", "[Workspace]", ""])

    lines = add_to_project(empty, "Fresh", "module").split(VBA_CRLF)

    assert lines[:2] == ['ID="{X}"', "Module=Fresh"]
    assert "Fresh=38, 38, 1786, 1030, " in lines


def test_a_module_added_back_to_an_emptied_project_is_whole(
    db: AccessDatabase, tmp_path: Path
) -> None:
    """Deleting the only module empties the container's listing and the
    PROJECT block, and adding one back has to rebuild both."""
    db.delete_module("Module1")
    emptied = tmp_path / "emptied.accdb"
    db.save(emptied)

    again = AccessDatabase(emptied)
    assert again.modules() == []
    again.create_module("Fresh", "Option Compare Database")
    out = tmp_path / "refilled.accdb"
    again.save(out)

    after = AccessDatabase(out)
    assert [m.name for m in after.modules()] == ["Fresh"]
    assert "Fresh" in [e.name for e in after.catalog() if e.type == -32761]
    modules_id = after._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    listing = next(
        row["Lv"]
        for _rid, row in after.table("MSysAccessStorage").rows_with_ids()
        if row["ParentId"] == modules_id and str(row["Name"]) == chr(3) + "DirData"
    )
    assert isinstance(listing, bytes)
    assert [name for name, _folder in dir_data_entries(listing)] == ["Fresh"]
    # and it deletes again, which needs that listing entry
    after.delete_module("Fresh")
    assert after.modules() == []
