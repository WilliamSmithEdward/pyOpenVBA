"""Forms and reports: the design blob, and what one costs the file.

The fixtures are a form Access made with a label and a text box, and a
report with its three sections.  The live gate opens what this writes in
Access's own designer.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._designs import (
    CONTROL_SLOTS,
    CONTROL_TYPES,
    PROPERTY_CODES,
    NAV_TYPES,
    OBJECT_TYPES,
    OPEN_CONTROL,
    build_design,
    parse_design,
    template,
    with_guid,
)
from pyopenvba.access._props import parse_property_blob
from pyopenvba.access._storage import dir_data_entries
from pyopenvba.access_read import AccessError

FIXTURES = Path(__file__).parent / "live_access_test"
FORM = FIXTURES / "form_with_controls.accdb"
REPORT = FIXTURES / "report.accdb"
TEMPLATE = (
    Path(__file__).parents[1]
    / "src"
    / "pyopenvba"
    / "_templates"
    / "blank_files"
    / "blank_database.accdb"
)


def opened(source: Path, tmp_path: Path, name: str) -> AccessDatabase:
    out = tmp_path / name
    shutil.copyfile(source, out)
    return AccessDatabase(out)


@pytest.fixture
def form_database(tmp_path: Path) -> AccessDatabase:
    return opened(FORM, tmp_path, "form.accdb")


@pytest.fixture
def report_database(tmp_path: Path) -> AccessDatabase:
    return opened(REPORT, tmp_path, "report.accdb")


@pytest.fixture
def blank(tmp_path: Path) -> AccessDatabase:
    return opened(TEMPLATE, tmp_path, "blank.accdb")


def design_blobs(db: AccessDatabase, kind: str) -> list[bytes]:
    container = db._design_container(kind)  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in db.table("MSysAccessStorage").rows_with_ids()]
    folders = {int(str(r["Id"])) for r in rows if r["ParentId"] == container and r["Type"] == 1}
    return [
        r["Lv"]
        for r in rows
        if r["ParentId"] in folders and str(r["Name"]) == "Blob" and isinstance(r["Lv"], bytes)
    ]


# --- the design blob ----------------------------------------------------------


def test_every_design_access_wrote_rebuilds_byte_for_byte(
    form_database: AccessDatabase, report_database: AccessDatabase
) -> None:
    """Parse it, build it again, require the same bytes -- the strongest
    check the file alone can give."""
    for db, kind in ((form_database, "form"), (report_database, "report")):
        blobs = design_blobs(db, kind)
        assert blobs
        for blob in blobs:
            header, objects, trailer = parse_design(blob)
            assert build_design(header, objects, trailer) == blob


def test_the_templates_rebuild_too() -> None:
    for kind in ("form", "report"):
        blob = template(kind, "blob")
        header, objects, trailer = parse_design(blob)
        assert build_design(header, objects, trailer) == blob


def test_a_form_reads_as_its_sections_and_controls(form_database: AccessDatabase) -> None:
    form = form_database.form("Rich")

    assert form.kind == "form"
    assert [s.name for s in form.sections] == ["Detail"]
    assert [(c.name, c.type_name) for c in form.controls] == [
        ("TheLabel", "Label"),
        ("TheBox", "TextBox"),
    ]
    # the design also carries the unnamed prototypes new controls are cut
    # from, and those are not controls
    assert len(form.objects) > len(form.controls) + len(form.sections) + 1


def test_a_report_reads_as_its_three_sections(report_database: AccessDatabase) -> None:
    report = report_database.report("Sheet")

    assert report.kind == "report"
    assert [s.name for s in report.sections] == [
        "PageHeaderSection",
        "Detail",
        "PageFooterSection",
    ]
    assert report.controls == ()


def test_the_control_markers_are_the_ones_access_writes(form_database: AccessDatabase) -> None:
    form = form_database.form("Rich")
    label = next(c for c in form.controls if c.name == "TheLabel")

    assert label.marker == OPEN_CONTROL
    assert label.type == 100 and CONTROL_TYPES[100] == "Label"
    assert label.records


def test_an_unknown_design_is_refused(form_database: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="no form named"):
        form_database.form("Nothing")
    with pytest.raises(AccessError, match="no report named"):
        form_database.report("Nothing")


def test_a_guid_can_be_replaced_without_disturbing_the_rest() -> None:
    blob = template("form", "blob")
    patched = with_guid(blob, bytes(range(16)))

    assert len(patched) == len(blob)
    assert bytes(range(16)) in patched
    _header, objects, _trailer = parse_design(patched)
    assert len(objects) == len(parse_design(blob)[1])


# --- creating -----------------------------------------------------------------


def test_a_form_can_be_created_in_a_database_with_none(blank: AccessDatabase) -> None:
    assert blank.forms() == []

    made = blank.create_form("Plain")

    assert made.name == "Plain" and made.kind == "form"
    assert [s.name for s in made.sections] == ["Detail"]
    assert [f.name for f in blank.forms()] == ["Plain"]


def test_a_report_can_be_created(blank: AccessDatabase) -> None:
    made = blank.create_report("Sheet")

    assert [s.name for s in made.sections] == [
        "PageHeaderSection",
        "Detail",
        "PageFooterSection",
    ]
    assert [r.name for r in blank.reports()] == ["Sheet"]


def test_creating_writes_the_catalog_rows_access_writes(blank: AccessDatabase) -> None:
    blank.create_form("Plain")
    blank.create_report("Sheet")

    catalog = {e.name: e for e in blank.catalog() if e.type in OBJECT_TYPES.values()}
    assert catalog["Plain"].type == OBJECT_TYPES["form"]
    assert catalog["Sheet"].type == OBJECT_TYPES["report"]
    nav = {
        str(r["Name"]): int(str(r["Type"]))
        for _rid, r in blank.table("MSysNavPaneObjectIDs").rows_with_ids()
    }
    assert nav["Plain"] == NAV_TYPES["form"]
    assert nav["Sheet"] == NAV_TYPES["report"]
    # ids step by one, as a macro's do and unlike a module's
    assert catalog["Sheet"].id - catalog["Plain"].id == 1


def test_each_design_gets_a_guid_of_its_own(blank: AccessDatabase) -> None:
    """The catalog row repeats the GUID the design carries, and two
    objects sharing one is not something Access writes."""
    blank.create_form("One")
    blank.create_form("Two")

    guids: list[bytes] = []
    for entry in blank.catalog():
        if entry.type != OBJECT_TYPES["form"]:
            continue
        row = next(
            r
            for _rid, r in blank.table("MSysObjects").rows_with_ids()
            if r["Id"] == entry.id
        )
        blob = row["LvProp"]
        assert isinstance(blob, bytes)
        guids.append(parse_property_blob(blob).object_properties["GUID"].raw)
    assert len(guids) == 2 and guids[0] != guids[1]
    for guid, name in zip(guids, ("One", "Two"), strict=True):
        assert guid in design_blobs(blank, "form")[0] or guid in b"".join(
            design_blobs(blank, "form")
        ), name


def test_creating_fills_the_container_streams(blank: AccessDatabase) -> None:
    blank.create_form("One")
    blank.create_form("Two")

    container = blank._design_container("form")  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in blank.table("MSysAccessStorage").rows_with_ids()]
    listing = next(r["Lv"] for r in rows if r["ParentId"] == container and str(r["Name"]) == "\x03DirData")
    assert isinstance(listing, bytes)
    assert dir_data_entries(listing) == [("One", "0"), ("Two", "1")]
    folders = {str(r["Name"]) for r in rows if r["ParentId"] == container and r["Type"] == 1}
    assert folders == {"0", "1"}
    # each folder carries the four streams a design needs
    for folder in folders:
        folder_id = next(
            int(str(r["Id"]))
            for r in rows
            if r["ParentId"] == container and str(r["Name"]) == folder and r["Type"] == 1
        )
        streams = {str(r["Name"]) for r in rows if r["ParentId"] == folder_id}
        assert streams == {"Blob", "TypeInfo", "BlobDelta", "PropData"}


@pytest.mark.parametrize("kind", ["form", "report"])
def test_creating_refuses_a_name_already_taken(blank: AccessDatabase, kind: str) -> None:
    make = blank.create_form if kind == "form" else blank.create_report
    make("Taken")
    with pytest.raises(AccessError, match="already exists"):
        make("TAKEN")


def test_creating_refuses_an_impossible_name(blank: AccessDatabase) -> None:
    with pytest.raises(AccessError, match="1 to 64 characters"):
        blank.create_form("x" * 65)


# --- deleting -----------------------------------------------------------------


def test_a_form_can_be_deleted(blank: AccessDatabase) -> None:
    blank.create_form("One")
    blank.create_form("Two")
    container = blank._design_container("form")  # pyright: ignore[reportPrivateUsage]
    before = len(
        [
            row
            for _rid, row in blank.table("MSysAccessStorage").rows_with_ids()
            if row["ParentId"] == container
        ]
    )

    blank.delete_form("One")

    assert [f.name for f in blank.forms()] == ["Two"]
    assert not [e for e in blank.catalog() if e.name == "One"]
    after = [
        row
        for _rid, row in blank.table("MSysAccessStorage").rows_with_ids()
        if row["ParentId"] == container
    ]
    assert len(after) == before - 1  # the folder goes with it


def test_a_report_can_be_deleted(blank: AccessDatabase) -> None:
    blank.create_report("Sheet")
    blank.delete_report("Sheet")

    assert blank.reports() == []
    assert "Sheet" not in {
        str(r["Name"]) for _rid, r in blank.table("MSysNavPaneObjectIDs").rows_with_ids()
    }


def test_a_design_created_after_a_delete_reuses_the_folder(blank: AccessDatabase) -> None:
    blank.create_form("One")
    blank.create_form("Two")
    blank.delete_form("One")
    blank.create_form("Three")

    container = blank._design_container("form")  # pyright: ignore[reportPrivateUsage]
    listing = next(
        r["Lv"]
        for _rid, r in blank.table("MSysAccessStorage").rows_with_ids()
        if r["ParentId"] == container and str(r["Name"]) == "\x03DirData"
    )
    assert isinstance(listing, bytes)
    assert dir_data_entries(listing) == [("Two", "1"), ("Three", "0")]


# --- adding a control ---------------------------------------------------------


def test_a_control_can_be_added_to_a_form(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    design = blank.add_control(
        "Built", "Label", "Title", left=240, top=240, width=2000, height=300, caption="Hello"
    )

    assert [(c.name, c.type_name) for c in design.controls] == [("Title", "Label")]
    label = design.controls[0]
    assert int.from_bytes(label.property_value(PROPERTY_CODES["Left"]) or b"", "little") == 240
    assert int.from_bytes(label.property_value(PROPERTY_CODES["Top"]) or b"", "little") == 240
    assert int.from_bytes(label.property_value(PROPERTY_CODES["Width"]) or b"", "little") == 2000
    assert int.from_bytes(label.property_value(PROPERTY_CODES["Height"]) or b"", "little") == 300
    caption = label.property_value(PROPERTY_CODES["Caption"])
    assert caption is not None and caption.decode("utf-16-le") == "Hello"


def test_the_markers_follow_how_many_controls_the_section_holds(
    blank: AccessDatabase,
) -> None:
    """One control is a single child, `0xFE`.  Two or more open a group:
    `0xFF` then `0xFD`.  Access writes both in one report -- a page header
    holding one control and a detail band holding two -- and refuses each
    in the other's place."""
    blank.create_form("Built")
    design = blank.add_control("Built", "Label", "One")
    assert [c.marker for c in design.controls] == [0xFE]

    design = blank.add_control("Built", "TextBox", "Two", caption="=1+1")
    assert [c.marker for c in design.controls] == [0xFF, 0xFD]
    assert [c.type_name for c in design.controls] == ["Label", "TextBox"]
    # and the sections are still where they were
    assert [s.name for s in design.sections] == ["Detail"]


def test_a_control_can_go_in_a_named_section(blank: AccessDatabase) -> None:
    blank.create_report("Sheet")
    blank.add_control("Sheet", "Label", "Banner", kind="report", section="PageHeaderSection")
    design = blank.add_control("Sheet", "Label", "Body", kind="report")

    assert [(c.name, c.marker) for c in design.controls] == [("Banner", 0xFE), ("Body", 0xFE)]
    # each sits inside its own section, in the order the sections come
    order = [o.name for o in design.objects if o.name]
    assert order.index("Banner") < order.index("Detail") < order.index("Body")


def test_an_unknown_section_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    with pytest.raises(AccessError, match="no 'PageFooterSection' section"):
        blank.add_control("Built", "Label", "One", section="PageFooterSection")


def test_a_control_survives_a_save(blank: AccessDatabase, tmp_path: Path) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "TextBox", "Box", left=100, top=200, caption="=2*21")
    out = tmp_path / "written.accdb"
    blank.save(out)

    control = AccessDatabase(out).form("Built").controls[0]
    assert control.name == "Box" and control.type_name == "TextBox"
    source = control.property_value(27)  # ControlSource
    assert source is not None and source.decode("utf-16-le") == "=2*21"


def test_a_control_can_be_added_to_a_report(blank: AccessDatabase) -> None:
    blank.create_report("Sheet")
    design = blank.add_control("Sheet", "Label", "Heading", kind="report", caption="Monthly")

    assert [(c.name, c.type_name) for c in design.controls] == [("Heading", "Label")]


def test_the_records_use_the_slots_access_uses(blank: AccessDatabase) -> None:
    """A record's id is its slot in the control type's own schema: a Label
    keeps its GUID at 234 and a TextBox at 250."""
    blank.create_form("Built")
    blank.add_control("Built", "Label", "One", caption="x")
    design = blank.add_control("Built", "TextBox", "Two")

    label, box = design.controls
    assert {r.id for r in label.records} >= {53, 96, 97, 98, 99, 220, 221, 234}
    assert {r.id for r in box.records} >= {55, 70, 96, 97, 98, 99, 220, 250}
    for control, kind in ((label, "Label"), (box, "TextBox")):
        for record in control.records:
            slot = CONTROL_SLOTS[kind]
            match = [s for s in slot.values() if s[0] == record.id]
            assert match, (kind, record.id)
            assert (record.code, record.value_type, record.width) == match[0][1:]


def test_a_control_type_without_measured_slots_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    with pytest.raises(AccessError, match="cannot be written yet"):
        blank.add_control("Built", "ComboBox", "Picker")


def test_a_name_already_on_the_design_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "Label", "One")
    with pytest.raises(AccessError, match="already has an object named"):
        blank.add_control("Built", "Label", "One")
