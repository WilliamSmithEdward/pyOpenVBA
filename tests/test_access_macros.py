"""Macros: the blob, and what a macro costs the file.

The fixture holds five macros Access itself created through
`LoadFromText`, chosen to cover an action with no arguments, two of the
same action, one with arguments, and a macro mixing all three.  The live
gate runs what this writes and reads the value back.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase, Macro, MacroAction
from pyopenvba.access._macros import (
    ACTION_IDS,
    HEADER,
    NAV_MACRO_TYPE,
    OBJECT_MACRO,
    build_macro,
    parse_macro,
)
from pyopenvba.access._storage import dir_data_entries
from pyopenvba.access_read import AccessError

FIXTURE = Path(__file__).parent / "live_access_test" / "macros.accdb"
TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)


@pytest.fixture
def written(tmp_path: Path) -> AccessDatabase:
    """Five macros Access made."""
    out = tmp_path / "macros.accdb"
    shutil.copyfile(FIXTURE, out)
    return AccessDatabase(out)


@pytest.fixture
def blank(tmp_path: Path) -> AccessDatabase:
    """The shipped template, which holds no macros at all."""
    out = tmp_path / "blank.accdb"
    shutil.copyfile(TEMPLATE, out)
    return AccessDatabase(out)


def scripts_rows(db: AccessDatabase) -> list[dict[str, object]]:
    scripts = db._scripts_id()  # pyright: ignore[reportPrivateUsage]
    return [
        row
        for _rid, row in db.table("MSysAccessStorage").rows_with_ids()
        if row["ParentId"] == scripts
    ]


# --- the blob -----------------------------------------------------------------


def test_the_fixture_reads_as_the_macros_access_made(written: AccessDatabase) -> None:
    assert [(m.name, [(a.name, a.arguments) for a in m.actions]) for m in written.macros()] == [
        ("M1Beep", [("Beep", ())]),
        ("M2Twice", [("Beep", ()), ("Beep", ())]),
        ("M3Msg", [("MsgBox", ("Hello", "Yes", "0"))]),
        ("M4Echo", [("Echo", ("No", "working"))]),
        ("M5Mixed", [("Beep", ()), ("MsgBox", ("Hi", "No", "1", "Title")), ("Beep", ())]),
    ]


def test_every_blob_access_wrote_rebuilds_byte_for_byte(written: AccessDatabase) -> None:
    """The strongest check the file alone can give: parse it, build it
    again, and require the same bytes."""
    scripts = written._scripts_id()  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in written.table("MSysAccessStorage").rows_with_ids()]
    folders = {
        int(str(r["Id"])) for r in rows if r["ParentId"] == scripts and r["Type"] == 1
    }
    blobs = [
        r["Lv"] for r in rows if r["ParentId"] in folders and str(r["Name"]) == "Blob"
    ]
    assert len(blobs) == 5
    for blob in blobs:
        assert isinstance(blob, bytes)
        assert build_macro(parse_macro(blob)) == blob


@pytest.mark.parametrize(
    "actions",
    [
        (MacroAction("Beep"),),
        (MacroAction("Beep"), MacroAction("Beep")),
        (MacroAction("MsgBox", ("Hello", "Yes", "0")),),
        (MacroAction("SetTempVar", ("name", "1 + 1")),),
        # a gap in the middle: the slot is left empty and reads back so
        (MacroAction("MsgBox", ("Hi", "", "1")),),
        (MacroAction("Beep"), MacroAction("Echo", ("No", "waiting")), MacroAction("Quit")),
    ],
)
def test_actions_round_trip(actions: tuple[MacroAction, ...]) -> None:
    assert parse_macro(build_macro(actions)) == actions


def test_a_blob_starts_with_its_header() -> None:
    blob = build_macro((MacroAction("Beep"),))
    assert blob.startswith(HEADER)
    assert len(blob) == 76  # what Access wrote for the same macro


def test_an_unknown_action_is_refused() -> None:
    with pytest.raises(AccessError, match="unknown macro action"):
        build_macro((MacroAction("NotAnAction"),))


def test_too_many_arguments_are_refused() -> None:
    with pytest.raises(AccessError, match="at most 10 arguments"):
        build_macro((MacroAction("MsgBox", tuple(str(i) for i in range(11))),))


def test_the_action_ids_are_the_ones_access_uses() -> None:
    """Three of these were confirmed against a second fixture where the
    correspondence was known independently."""
    assert ACTION_IDS["Beep"] == 4
    assert ACTION_IDS["Echo"] == 9
    assert ACTION_IDS["MsgBox"] == 22
    assert len(ACTION_IDS) == len(set(ACTION_IDS.values()))


# --- writing ------------------------------------------------------------------


def test_a_macro_can_be_created_in_a_database_with_none(blank: AccessDatabase) -> None:
    """The first macro brings the `Scripts` listing with it."""
    assert blank.macros() == []

    made = blank.create_macro("Ping", [MacroAction("Beep")])
    assert made == Macro("Ping", (MacroAction("Beep"),))
    assert [m.name for m in blank.macros()] == ["Ping"]
    assert dir_data_entries(
        next(r["Lv"] for r in scripts_rows(blank) if str(r["Name"]) == "\x03DirData")  # pyright: ignore[reportArgumentType]
    ) == [("Ping", "0")]


def test_a_created_macro_costs_what_access_charges(blank: AccessDatabase) -> None:
    blank.create_macro("Ping", [MacroAction("Beep")])

    catalog = [e for e in blank.catalog() if e.type == OBJECT_MACRO]
    assert [e.name for e in catalog] == ["Ping"]
    nav = {
        str(r["Name"]): int(str(r["Type"]))
        for _rid, r in blank.table("MSysNavPaneObjectIDs").rows_with_ids()
    }
    assert nav["Ping"] == NAV_MACRO_TYPE
    # a macro is not filed under a navigation-pane group, where a module is
    groups = {
        int(str(r["ObjectID"]))
        for _rid, r in blank.table("MSysNavPaneGroupToObjects").rows_with_ids()
    }
    assert catalog[0].id not in groups


def test_macro_object_ids_step_by_one(blank: AccessDatabase) -> None:
    """A module's step by four; the step is what an object reserves."""
    first = blank.create_macro("One", [MacroAction("Beep")])
    second = blank.create_macro("Two", [MacroAction("Beep")])
    ids = {e.name: e.id for e in blank.catalog() if e.type == OBJECT_MACRO}
    assert ids[second.name] - ids[first.name] == 1


def test_macro_folders_start_at_zero_and_run_on(blank: AccessDatabase) -> None:
    """`Scripts` starts empty, so the first macro takes `0`; the listing
    it creates is a non-folder row, so the second takes `1`."""
    blank.create_macro("One", [MacroAction("Beep")])
    blank.create_macro("Two", [MacroAction("Beep")])
    blank.create_macro("Three", [MacroAction("Beep")])
    listing = next(r["Lv"] for r in scripts_rows(blank) if str(r["Name"]) == "\x03DirData")
    assert isinstance(listing, bytes)
    assert dir_data_entries(listing) == [("One", "0"), ("Two", "1"), ("Three", "2")]


def test_a_created_macro_survives_a_save(blank: AccessDatabase, tmp_path: Path) -> None:
    blank.create_macro(
        "Probe", [MacroAction("SetTempVar", ("probe", "42")), MacroAction("Beep")]
    )
    out = tmp_path / "written.accdb"
    blank.save(out)

    reopened = AccessDatabase(out)
    assert reopened.macro("PROBE").actions == (
        MacroAction("SetTempVar", ("probe", "42")),
        MacroAction("Beep"),
    )


def test_a_macro_can_be_deleted(blank: AccessDatabase) -> None:
    blank.create_macro("One", [MacroAction("Beep")])
    blank.create_macro("Two", [MacroAction("Beep")])
    before = len(scripts_rows(blank))

    blank.delete_macro("One")

    assert [m.name for m in blank.macros()] == ["Two"]
    assert not [e for e in blank.catalog() if e.type == OBJECT_MACRO and e.name == "One"]
    # the folder and its Blob go with it
    assert len(scripts_rows(blank)) == before - 1
    assert "One" not in {
        str(r["Name"]) for _rid, r in blank.table("MSysNavPaneObjectIDs").rows_with_ids()
    }


@pytest.mark.parametrize(
    ("name", "message"),
    [("", "1 to 64 characters"), ("x" * 65, "1 to 64 characters")],
)
def test_create_refuses_an_impossible_name(
    blank: AccessDatabase, name: str, message: str
) -> None:
    with pytest.raises(AccessError, match=message):
        blank.create_macro(name, [MacroAction("Beep")])


def test_create_refuses_a_name_already_taken(blank: AccessDatabase) -> None:
    blank.create_macro("Ping", [MacroAction("Beep")])
    with pytest.raises(AccessError, match="already exists"):
        blank.create_macro("PING", [MacroAction("Beep")])


def test_an_unknown_macro_is_refused(blank: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="no macro named"):
        blank.macro("Nothing")
