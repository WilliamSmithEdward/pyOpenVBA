"""Forms and reports: the design blob, and what one costs the file.

The fixtures are a form Access made with a label and a text box, a
report with its three sections, and a database in which Access built a
form and a report holding one of every control and every section, plus
copies of the form it then edited.  The live gate opens what this writes
in Access's own designer.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from pyopenvba.access import AccessDatabase
from pyopenvba.access._designs import (
    CONTROL_SLOTS,
    CONTROL_TYPES,
    NAV_TYPES,
    OBJECT_TYPES,
    OPEN_CONTROL,
    OPEN_SECTION,
    OPEN_SIBLING,
    READ_ONLY_TYPES,
    TYPE_CODES as CONTROL_CODES,
    PROPERTY_CODES,
    PROPERTY_SLOTS,
    property_code,
    build_design,
    pack_type_info,
    parse_design,
    prototype_records,
    template,
    type_info,
    type_info_entries,
    type_info_members,
    with_guid,
)
from pyopenvba.access._props import parse_property_blob

PROPERTY_CODES_BY_NUMBER = {code: name for name, code in PROPERTY_CODES.items()}
from pyopenvba.access._storage import dir_data_entries
from pyopenvba.vba import decompress
from pyopenvba.access_read import AccessError

FIXTURES = Path(__file__).parent / "live_access_test"
FORM = FIXTURES / "form_with_controls.accdb"
REPORT = FIXTURES / "report.accdb"
WITH_CODE = FIXTURES / "form_with_code.accdb"
EVERY = FIXTURES / "designs_every.accdb"
DIR_DATA = chr(3) + "DirData"
NEWLINE = chr(10)
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


@pytest.fixture
def every_database(tmp_path: Path) -> AccessDatabase:
    return opened(EVERY, tmp_path, "every.accdb")


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
    """A name the reader does not know at all is refused rather than
    guessed at."""
    blank.create_form("Built")
    with pytest.raises(AccessError, match="cannot be written yet"):
        blank.add_control("Built", "PivotTable", "Pivot")


@pytest.mark.parametrize("kind", READ_ONLY_TYPES)
def test_a_control_that_is_only_read_says_so(blank: AccessDatabase, kind: str) -> None:
    """A chart, a navigation control and an Edge browser parse and report
    their type, but each carries records this project cannot name -- one
    of them twice, at two ids -- so writing one is refused with the reason
    rather than attempted."""
    assert kind in CONTROL_TYPES.values()
    blank.create_form("Built")
    with pytest.raises(AccessError, match="read but not written"):
        blank.add_control("Built", kind, "Nope")


@pytest.mark.parametrize(
    "kind",
    [
        "Label",
        "TextBox",
        "CommandButton",
        "ToggleButton",
        "OptionButton",
        "CheckBox",
        "OptionGroup",
        "ListBox",
        "ComboBox",
        "Rectangle",
        "Line",
        "Image",
        "PageBreak",
        "BoundObjectFrame",
        "ObjectFrame",
        "Subform",
        "Tab",
        "CustomControl",
        "Attachment",
        "WebBrowser",
        "Chart",
        "EdgeBrowser",
    ],
)
def test_every_measured_control_type_can_be_written(blank: AccessDatabase, kind: str) -> None:
    blank.create_form("Built")
    design = blank.add_control("Built", kind, f"My{kind}", left=240, top=480, width=1400, height=300)

    control = next(o for o in design.objects if o.name == f"My{kind}")
    assert control.type is not None and CONTROL_TYPES[control.type] == kind
    ids = [r.id for r in control.records]
    assert ids == sorted(ids), "records have to go out in id order"
    assert {r.id for r in control.records} <= {s[0] for s in CONTROL_SLOTS[kind].values()}


@pytest.mark.parametrize("count", [1, 2, 3, 5, 9])
def test_the_group_marker_counts_the_controls_it_opens(
    blank: AccessDatabase, count: int
) -> None:
    """A form Access built with eleven controls carries `0xFF 11`.  Ours
    has to carry its own count: Access does not refuse a wrong one, it
    opens the form and shows only that many controls.
    """
    blank.create_form("Built")
    design = blank.form("Built")
    for i in range(count):
        design = blank.add_control("Built", "Label", f"C{i}", top=240 + i * 400)

    controls = [o for o in design.objects if o.name and o.name.startswith("C")]
    assert len(controls) == count
    if count == 1:
        assert controls[0].marker == OPEN_SECTION
    else:
        assert controls[0].marker == OPEN_CONTROL
        assert controls[0].code == count, "the opener names how many follow it"
        assert all(o.marker == OPEN_SIBLING for o in controls[1:])
        assert [o.code for o in controls[1:]] == [o.type for o in controls[1:]]


def test_a_page_break_carries_no_width(blank: AccessDatabase) -> None:
    """Access writes a page break with only its top, so writing a width
    would be inventing a slot it has no id for."""
    blank.create_form("Built")
    design = blank.add_control("Built", "PageBreak", "Break", top=1440, width=999)

    control = next(o for o in design.objects if o.name == "Break")
    codes = {r.code for r in control.records}
    assert PROPERTY_CODES["Top"] in codes
    assert PROPERTY_CODES["Width"] not in codes
    assert PROPERTY_CODES["Height"] not in codes


def test_the_tab_index_counts_the_controls_that_take_focus(blank: AccessDatabase) -> None:
    """Access omits the record for the first one and numbers the rest; a
    rectangle is not in the running."""
    blank.create_form("Built")
    blank.add_control("Built", "TextBox", "First")
    blank.add_control("Built", "Rectangle", "Box")
    design = blank.add_control("Built", "CommandButton", "Third")

    def tab_of(name: str) -> int | None:
        control = next(o for o in design.objects if o.name == name)
        found = [r for r in control.records if r.code == PROPERTY_CODES["TabIndex"]]
        return int.from_bytes(found[0].value, "little") if found else None

    assert tab_of("First") is None
    assert tab_of("Box") is None
    assert tab_of("Third") == 1


def test_a_name_already_on_the_design_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "Label", "One")
    with pytest.raises(AccessError, match="already has an object named"):
        blank.add_control("Built", "Label", "One")


# --- code behind a form -------------------------------------------------------


@pytest.fixture
def coded(tmp_path: Path) -> AccessDatabase:
    """A form Access gave a module of its own."""
    return opened(WITH_CODE, tmp_path, "coded.accdb")


def test_code_behind_a_form_reads_as_a_module(coded: AccessDatabase) -> None:
    """Access keeps it as a class module named after the form, and the
    ordinary module reader finds it."""
    modules = {m.name: m for m in coded.modules()}

    assert "Form_Coded" in modules
    assert modules["Form_Coded"].kind == "class"
    assert "Answer = 42" in modules["Form_Coded"].source


def test_code_behind_a_form_can_be_replaced(coded: AccessDatabase, tmp_path: Path) -> None:
    coded.set_module_source("Form_Coded", "Option Compare Database\n\nPublic Function Answer() As Variant\n    Answer = 4242\nEnd Function")
    out = tmp_path / "written.accdb"
    coded.save(out)

    assert "Answer = 4242" in AccessDatabase(out).module("Form_Coded").source


def test_a_form_module_is_not_an_object_under_modules(coded: AccessDatabase) -> None:
    """It belongs to the form, so it has no storage folder of its own and
    no catalog row -- only the dir stream, its own stream row, a
    `PROJECTwm` entry, and a `DocClass=` line rather than a `Class=` one."""
    modules_id = coded._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in coded.table("MSysAccessStorage").rows_with_ids()]
    listing = next(
        r["Lv"] for r in rows if r["ParentId"] == modules_id and str(r["Name"]) == DIR_DATA
    )
    assert isinstance(listing, bytes)
    assert [name for name, _folder in dir_data_entries(listing)] == ["Module1"]
    assert [e.name for e in coded.catalog() if e.type == -32761] == ["Module1"]
    project = next(r["Lv"] for r in rows if str(r["Name"]) == "PROJECT")
    assert isinstance(project, bytes)
    assert b"DocClass=Form_Coded/&H00000000" in project
    assert b"Class=Form_Coded" not in project.replace(b"DocClass=Form_Coded", b"")
    workspace = next(r["Lv"] for r in rows if str(r["Name"]) == "PROJECTwm")
    assert isinstance(workspace, bytes)
    assert b"Form_Coded" in workspace


def test_a_module_after_a_form_module_still_deletes(
    coded: AccessDatabase, tmp_path: Path
) -> None:
    """`delete_module` used to take the storage folder by position among
    `modules()` (GitHub issue #19).

    That list is dir-stream order and counts the module behind a form,
    which has no folder under `Modules`, so the two lists are different
    lengths.  Here they are three and two, and the module added last sat
    one place past the end.
    """
    added = tmp_path / "added.accdb"
    coded.create_module("MNewbie", "Option Compare Database")
    coded.save(added)

    again = AccessDatabase(added)
    assert [m.name for m in again.modules()] == ["Module1", "Form_Coded", "MNewbie"]
    again.delete_module("MNewbie")
    out = tmp_path / "deleted.accdb"
    again.save(out)

    after = AccessDatabase(out)
    assert [m.name for m in after.modules()] == ["Module1", "Form_Coded"]
    # Module1 keeps the folder it had; the deletion took MNewbie's.
    modules_id = after._vba_storage_ids()[0]  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in after.table("MSysAccessStorage").rows_with_ids()]
    listing = next(
        r["Lv"] for r in rows if r["ParentId"] == modules_id and str(r["Name"]) == DIR_DATA
    )
    assert isinstance(listing, bytes)
    assert dict(dir_data_entries(listing)) == {"Module1": "0"}


def test_deleting_the_code_behind_a_form_is_refused(coded: AccessDatabase) -> None:
    """It has no folder of its own, so there is nothing for
    `delete_module` to remove and the old positional lookup would have
    taken the next module's folder instead."""
    with pytest.raises(AccessError, match="code behind a form or report"):
        coded.delete_module("Form_Coded")

    assert "Form_Coded" in [m.name for m in coded.modules()]


def test_code_can_be_put_behind_a_form_that_has_none(blank: AccessDatabase) -> None:
    blank.create_form("Behind")
    assert not [m for m in blank.modules() if m.name.startswith("Form_")]

    module = blank.set_design_code(
        "Behind", "Option Compare Database" + NEWLINE + "Public Sub Go()" + NEWLINE + "End Sub"
    )

    assert module.name == "Form_Behind" and module.kind == "class"
    assert "Public Sub Go()" in module.source
    # it is the design's, not the Modules container's
    assert [e.name for e in blank.catalog() if e.type == -32761] == ["Module1"]


def test_the_design_and_its_module_share_a_clsid(blank: AccessDatabase) -> None:
    """`TypeInfo` carries it and the module's `VB_Base` repeats it; that
    pairing is what makes the form answer to the module."""
    import uuid

    blank.create_form("Behind")
    blank.set_design_code("Behind", "Option Compare Database")

    container = blank._design_container("form")  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in blank.table("MSysAccessStorage").rows_with_ids()]
    folder_id = next(
        int(str(r["Id"])) for r in rows if r["ParentId"] == container and r["Type"] == 1
    )
    info = next(r["Lv"] for r in rows if r["ParentId"] == folder_id and str(r["Name"]) == "TypeInfo")
    prop = next(r["Lv"] for r in rows if r["ParentId"] == folder_id and str(r["Name"]) == "PropData")
    assert isinstance(info, bytes) and isinstance(prop, bytes)
    clsid = str(uuid.UUID(bytes_le=info[16:32])).upper()
    # `source` leaves the attribute block out, so the stream is what to read
    stream = next(
        r["Lv"]
        for r in rows
        if str(r["Name"]) == blank.module("Form_Behind").stream_name
    )
    assert isinstance(stream, bytes)
    assert clsid in decompress(stream).decode("latin-1")
    assert prop[9] == 1  # the byte that says the design has a module


def test_a_project_line_marks_the_module_as_the_design_s(blank: AccessDatabase) -> None:
    blank.create_form("Behind")
    blank.set_design_code("Behind", "Option Compare Database")

    project = next(
        r["Lv"]
        for _rid, r in blank.table("MSysAccessStorage").rows_with_ids()
        if str(r["Name"]) == "PROJECT"
    )
    assert isinstance(project, bytes)
    assert b"DocClass=Form_Behind/&H00000000" in project


def test_code_behind_a_report_is_named_for_it(blank: AccessDatabase) -> None:
    blank.create_report("Sheet")
    module = blank.set_design_code("Sheet", "Option Compare Database", kind="report")

    assert module.name == "Report_Sheet"


def test_setting_code_twice_replaces_it(blank: AccessDatabase) -> None:
    blank.create_form("Behind")
    blank.set_design_code("Behind", "Option Compare Database" + NEWLINE + "Public Sub One()")
    module = blank.set_design_code(
        "Behind", "Option Compare Database" + NEWLINE + "Public Sub Two()"
    )

    assert "Public Sub Two()" in module.source
    assert "Public Sub One()" not in module.source
    assert len([m for m in blank.modules() if m.name == "Form_Behind"]) == 1


def test_no_code_names_two_properties(blank: AccessDatabase) -> None:
    """The table is read both ways -- name to code when writing, code to
    name when reading a design -- so a repeated code would make one of
    those a lie."""
    codes = list(PROPERTY_CODES.values())
    assert len(set(codes)) == len(codes)


def test_every_slot_names_a_property_the_table_knows(blank: AccessDatabase) -> None:
    for kind, slots in CONTROL_SLOTS.items():
        for name, (_id, code, _type, _width) in slots.items():
            assert property_code(name) == code, f"{kind}.{name} disagrees with the table"


def test_a_control_placed_by_hand_claims_no_layout(blank: AccessDatabase) -> None:
    """Access writes 596, 597 and 600 on a chart or Edge browser because
    its designer drops a new control into a layout.  A control given an
    explicit position must not carry them: with them, Access stacks the
    control under whatever else claims the same layout instead of leaving
    it where it was put."""
    blank.create_form("Modern")
    design = blank.add_control("Modern", "Chart", "Graph", left=240, top=240, width=5000)

    graph = next(o for o in design.objects if o.name == "Graph")
    codes = {r.code for r in graph.records}
    assert codes.isdisjoint({596, 597, 600})
    assert PROPERTY_CODES["Left"] in codes and PROPERTY_CODES["Top"] in codes


def test_a_navigation_control_says_why_it_is_only_read(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    with pytest.raises(AccessError, match="navigation control"):
        blank.add_control("Built", "NavigationControl", "Nav")


def test_a_page_belongs_to_a_tab_control(blank: AccessDatabase) -> None:
    """A tab control holds its pages as a group of its own, written right
    after it, and the section's own count does not include them."""
    blank.create_form("Tabbed")
    blank.add_control("Tabbed", "TextBox", "Box", top=240)
    blank.add_control("Tabbed", "Tab", "Tabs", top=800, width=4000, height=2000)
    blank.add_control("Tabbed", "Page", "First", parent="Tabs", caption="One")
    blank.add_control("Tabbed", "Page", "Second", parent="Tabs", caption="Two")
    design = blank.add_control("Tabbed", "CommandButton", "Go", top=3000)

    placed = [(o.name, o.marker, o.code) for o in design.objects if o.marker is not None and o.name]
    assert placed == [
        # A sibling of the three control-defaults objects ahead of it.
        ("Detail", OPEN_SIBLING, CONTROL_CODES["Detail"]),
        # Three controls on the section: the pages are not among them.
        ("Box", OPEN_CONTROL, 3),
        ("Tabs", OPEN_SIBLING, CONTROL_CODES["Tab"]),
        ("First", OPEN_CONTROL, 2),
        ("Second", OPEN_SIBLING, CONTROL_CODES["Page"]),
        ("Go", OPEN_SIBLING, CONTROL_CODES["CommandButton"]),
    ]


def test_a_page_without_a_tab_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    with pytest.raises(AccessError, match="needs a parent Tab"):
        blank.add_control("Built", "Page", "Loose")


def test_only_a_page_takes_a_parent(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "Tab", "Tabs")
    with pytest.raises(AccessError, match="needs a parent Tab"):
        blank.add_control("Built", "Label", "Stuck", parent="Tabs")


def test_a_parent_that_is_not_a_tab_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "Label", "Plain")
    with pytest.raises(AccessError, match="holds no Page"):
        blank.add_control("Built", "Page", "Nope", parent="Plain")


def test_a_parent_that_is_not_there_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Built")
    blank.add_control("Built", "Tab", "Tabs")
    with pytest.raises(AccessError, match="no control named"):
        blank.add_control("Built", "Page", "Nope", parent="Missing")


def test_a_control_added_after_a_tab_does_not_join_its_pages(
    blank: AccessDatabase,
) -> None:
    """The bug this guards against is silent: a control swallowed into the
    tab's group would still parse, and Access would show it on a page."""
    blank.create_form("Tabbed")
    blank.add_control("Tabbed", "Tab", "Tabs", width=4000, height=2000)
    blank.add_control("Tabbed", "Page", "Only", parent="Tabs")
    design = blank.add_control("Tabbed", "Label", "After", top=3000)

    names = [o.name for o in design.objects if o.marker is not None and o.name]
    assert names == ["Detail", "Tabs", "Only", "After"]
    by_name = {o.name: o for o in design.objects}
    assert by_name["Tabs"].code == 2, "the section holds the tab and the label"
    assert by_name["Only"].marker == OPEN_SECTION, "the tab's one page"
    assert by_name["After"].marker == OPEN_SIBLING


def test_a_property_the_control_already_has_is_replaced(blank: AccessDatabase) -> None:
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title", caption="before")

    design = blank.set_control_property("Styled", "Title", "Caption", "after")

    title = next(o for o in design.objects if o.name == "Title")
    assert title.property_value(PROPERTY_CODES["Caption"]) == "after".encode("utf-16-le")
    # One record, not two: the old one was replaced where it stood.
    assert [r.code for r in title.records].count(PROPERTY_CODES["Caption"]) == 1


def test_a_property_the_control_lacks_is_written_at_its_own_slot(
    blank: AccessDatabase,
) -> None:
    """A record's id is its slot in that control type's schema, so a
    property nothing has written yet still has one place it belongs."""
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title", caption="hello")

    design = blank.set_control_property("Styled", "Title", "FontSize", 18)

    title = next(o for o in design.objects if o.name == "Title")
    assert title.property_value(PROPERTY_CODES["FontSize"]) == (18).to_bytes(2, "little")
    ids = [r.id for r in title.records]
    assert ids == sorted(ids), "records have to stay in id order"
    assert PROPERTY_SLOTS["Label"]["FontSize"][0] in ids


@pytest.mark.parametrize(
    ("prop", "value", "expected"),
    [
        ("Caption", "hello", "hello".encode("utf-16-le")),
        ("FontName", "Consolas", "Consolas".encode("utf-16-le")),
        ("FontSize", 18, (18).to_bytes(2, "little")),
        ("FontWeight", 700, (700).to_bytes(2, "little")),
        ("ForeColor", 255, (255).to_bytes(4, "little")),
        ("BackColor", 65535, (65535).to_bytes(4, "little")),
        ("Left", 1440, (1440).to_bytes(2, "little")),
    ],
)
def test_each_kind_of_value_is_written_the_way_its_slot_says(
    blank: AccessDatabase, prop: str, value: object, expected: bytes
) -> None:
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title", caption="x")

    design = blank.set_control_property("Styled", "Title", prop, value)

    title = next(o for o in design.objects if o.name == "Title")
    assert title.property_value(PROPERTY_CODES[prop]) == expected


def test_the_design_has_properties_of_its_own(blank: AccessDatabase) -> None:
    blank.create_form("Styled")

    design = blank.set_design_property("Styled", "Caption", "My window")

    assert design.root.property_value(PROPERTY_CODES["Caption"]) == (
        "My window".encode("utf-16-le")
    )


def test_a_section_takes_properties_too(blank: AccessDatabase) -> None:
    blank.create_form("Styled")

    design = blank.set_control_property("Styled", "Detail", "Height", 2880)

    detail = next(o for o in design.objects if o.name == "Detail")
    assert detail.property_value(PROPERTY_CODES["Height"]) == (2880).to_bytes(2, "little")


def test_a_property_that_type_does_not_have_is_refused(blank: AccessDatabase) -> None:
    """The table holds what was measured off controls Access wrote, so a
    name that is not in it would be written at an id that means something
    else -- which is worse than refusing."""
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title")

    with pytest.raises(AccessError, match="has no 'ListRows' to set"):
        blank.set_control_property("Styled", "Title", "ListRows", 9)


def test_a_control_that_is_not_there_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Styled")
    with pytest.raises(AccessError, match="no object named 'Missing'"):
        blank.set_control_property("Styled", "Missing", "Caption", "x")


def test_the_wrong_kind_of_value_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title")

    with pytest.raises(AccessError, match="takes text"):
        blank.set_control_property("Styled", "Title", "Caption", 7)
    with pytest.raises(AccessError, match="takes a number"):
        blank.set_control_property("Styled", "Title", "FontSize", "big")


def test_a_number_too_large_for_its_slot_is_refused(blank: AccessDatabase) -> None:
    blank.create_form("Styled")
    blank.add_control("Styled", "Label", "Title")

    with pytest.raises(AccessError, match="does not fit"):
        blank.set_control_property("Styled", "Title", "FontSize", 70000)


def test_every_slot_names_a_code_the_table_knows() -> None:
    for kind, slots in PROPERTY_SLOTS.items():
        for name, slot in slots.items():
            assert property_code(name) == slot[1], f"{kind}.{name} disagrees"


def test_the_slots_a_new_control_gets_agree_with_the_schema() -> None:
    """Two tables describe the same records -- what a new control is given
    and where each property lives -- so an id in one that contradicts the
    other would put a record in the wrong place."""
    for kind, slots in CONTROL_SLOTS.items():
        for name, (ident, code, value_type, _width) in slots.items():
            schema = PROPERTY_SLOTS.get(kind, {}).get(name)
            if schema is None:
                continue
            assert schema[0] == ident, f"{kind}.{name}: id {schema[0]} against {ident}"
            assert schema[1] == code
            assert schema[2] == value_type


# --- the TypeInfo stream and the row source kind ------------------------------


def type_info_stream(db: AccessDatabase, kind: str, name: str) -> bytes:
    """The TypeInfo stream beside the design called `name`."""
    container = db._design_container(kind)  # pyright: ignore[reportPrivateUsage]
    rows = [row for _rid, row in db.table("MSysAccessStorage").rows_with_ids()]
    listing = next(
        row["Lv"] for row in rows if row["ParentId"] == container and str(row["Name"]) == DIR_DATA
    )
    assert isinstance(listing, bytes)
    folder = dict(dir_data_entries(listing))[name]
    folder_id = next(
        int(str(row["Id"]))
        for row in rows
        if row["ParentId"] == container and row["Type"] == 1 and str(row["Name"]) == folder
    )
    stream = next(
        row["Lv"] for row in rows if row["ParentId"] == folder_id and str(row["Name"]) == "TypeInfo"
    )
    assert isinstance(stream, bytes)
    return stream


def listed(db: AccessDatabase, kind: str, name: str) -> list[tuple[str, int]]:
    return [(e.name, e.ordinal) for e in type_info_entries(type_info_stream(db, kind, name))]


#: The designs Access built into the fixture: a form and a report with one
#: of every control, a form and a report with every section, a report with
#: two group levels, a form of controls held by other controls, and three
#: copies of the first form that Access then edited.
EVERY_DESIGNS = (
    ("form", "Every"),
    ("form", "Hdr"),
    ("form", "Nest"),
    ("form", "EveryDel"),
    ("form", "EveryRen"),
    ("form", "EveryAdd"),
    ("report", "EveryR"),
    ("report", "RHdr"),
    ("report", "TwoGroups"),
)


def test_type_info_lists_what_access_lists_with_the_ids_access_gives(
    form_database: AccessDatabase, report_database: AccessDatabase, every_database: AccessDatabase
) -> None:
    """For every design Access wrote: carrying its stream through the design
    changes nothing, and each member's id comes out as Access wrote it --
    sections per kind, controls per kind, a label attached to a control or
    a button inside an option group by its own class, an ActiveX control
    with its 36-byte tail."""
    designs = [
        (form_database, "form", form_database.forms()[0].name),
        (report_database, "report", report_database.reports()[0].name),
        *((every_database, kind, name) for kind, name in EVERY_DESIGNS),
    ]
    checked = 0
    for db, kind, name in designs:
        design = db.form(name) if kind == "form" else db.report(name)
        engine = type_info_stream(db, kind, name)
        assert type_info(kind, design.objects, engine) == engine, name
        members = {member.name: member for member in type_info_members(kind, design.objects)}
        entries = type_info_entries(engine)
        assert {entry.name for entry in entries} == set(members), name
        for entry in entries:
            member = members[entry.name]
            assert (entry.ident, entry.tail) == (member.ident, member.tail), (name, entry)
            checked += 1
    assert checked > 150


def test_type_info_is_carried_forward_the_way_access_carries_it(
    every_database: AccessDatabase,
) -> None:
    """Three copies of one form, each edited in Access: controls deleted,
    controls renamed, controls added.  From the stream the copies started
    with, the same edits here give the streams Access wrote: a removed
    member drops out and the rest keep their ordinals, a renamed one moves
    to the end with its ordinal, a new one takes the ordinal after the
    highest present."""
    db = every_database
    after_add = type_info_entries(type_info_stream(db, "form", "EveryAdd"))
    before = pack_type_info(
        "form", bytes(16), [e for e in after_add if e.name not in ("New1", "New2")]
    )

    deleted = type_info(
        "form", db.form("EveryDel").objects, before
    )
    assert type_info_entries(deleted) == type_info_entries(type_info_stream(db, "form", "EveryDel"))
    assert [e.name for e in type_info_entries(deleted)][:6] == ["Lbl", "Rect", "Ln", "Img", "Chk", "Grp"]
    assert [e.ordinal for e in type_info_entries(deleted)][:6] == [0, 1, 2, 3, 6, 7]

    renamed = type_info(
        "form", db.form("EveryRen").objects, before, renamed={"Rect": "Rect2", "Txt": "Txt2"}
    )
    assert type_info_entries(renamed) == type_info_entries(type_info_stream(db, "form", "EveryRen"))
    assert [(e.name, e.ordinal) for e in type_info_entries(renamed)][-2:] == [("Rect2", 1), ("Txt2", 12)]

    added = type_info("form", db.form("EveryAdd").objects, before)
    assert type_info_entries(added) == after_add
    assert [(e.name, e.ordinal) for e in after_add][-2:] == [("New1", 34), ("New2", 35)]


def test_type_info_follows_the_controls_added_removed_and_renamed(blank: AccessDatabase) -> None:
    form = blank.add_form("Members")
    form.add_control("TextBox", "Qty", top=240)
    form.add_control("CommandButton", "Go", top=700)
    entries = type_info_entries(type_info_stream(blank, "form", "Members"))
    assert [(e.name, e.ordinal, e.ident) for e in entries] == [
        ("Detail", 0, 0x1898),
        ("Qty", 1, 0x126D),
        ("Go", 2, 0x0B68),
    ]

    form.remove_control("Qty")
    assert listed(blank, "form", "Members") == [("Detail", 0), ("Go", 2)]

    form.add_control("ListBox", "Lines", top=1200)
    assert listed(blank, "form", "Members") == [("Detail", 0), ("Go", 2), ("Lines", 3)]

    form.control("Go").set_property("Name", "Run")
    assert listed(blank, "form", "Members") == [("Detail", 0), ("Lines", 3), ("Run", 2)]
    assert [c.name for c in form.controls] == ["Run", "Lines"]


def test_type_info_on_a_report_uses_the_report_ids(blank: AccessDatabase) -> None:
    report = blank.add_report("Sheet")
    report.add_control("TextBox", "Amount", top=240)
    report.add_control("Label", "Banner", section="PageHeaderSection", caption="Banner")
    entries = type_info_entries(type_info_stream(blank, "report", "Sheet"))
    assert [(e.name, e.ordinal, e.ident) for e in entries] == [
        ("Detail", 0, 0x1998),
        ("PageHeaderSection", 1, 0x1F9B),
        ("PageFooterSection", 2, 0x1F9C),
        ("Amount", 3, 0x1B6D),
        ("Banner", 4, 0x1B64),
    ]


def test_every_section_type_reads_as_a_section(every_database: AccessDatabase) -> None:
    form = every_database.form("Hdr")
    assert [s.name for s in form.sections] == [
        "FormHeader",
        "PageHeaderSection",
        "Detail",
        "PageFooterSection",
        "FormFooter",
    ]
    assert [s.kind for s in form.sections] == [
        "HeaderSection",
        "PageHeaderSection",
        "Detail",
        "PageFooterSection",
        "FooterSection",
    ]
    assert [c.name for c in form.controls] == [
        "InHeader",
        "InPageHeader",
        "InDetail",
        "InPageFooter",
        "InFooter",
    ]
    report = every_database.report("RHdr")
    assert [s.name for s in report.sections] == [
        "ReportHeader",
        "PageHeaderSection",
        "GroupHeader0",
        "Detail",
        "GroupFooter1",
        "PageFooterSection",
        "ReportFooter",
    ]
    assert len(report.controls) == 7
    two = every_database.report("TwoGroups")
    assert [s.kind for s in two.sections].count("GroupHeaderSection") == 2


def test_a_type_without_a_measured_id_is_refused(every_database: AccessDatabase) -> None:
    """The navigation control and its button have ids on a form and are not
    allowed on a report, where nothing was measured for them."""
    form = every_database.form("Every")
    assert type_info_members("form", form.objects)
    with pytest.raises(AccessError, match="no TypeInfo id has been measured"):
        type_info_members("report", form.objects)


# --- control defaults ---------------------------------------------------------


def top_level(blob: bytes) -> list[tuple[int | None, int | None, int | None, str | None]]:
    """(marker, type, code, name) of the objects directly under the design:
    the control-defaults objects, then the sections."""
    _header, objects, _trailer = parse_design(blob)
    return [(o.marker, o.type, o.code, o.name) for o in objects[1:] if not o.name or o.is_section]


def test_the_first_control_of_a_type_brings_its_defaults_object(blank: AccessDatabase) -> None:
    """Access keeps one nameless object per control type ahead of the
    sections, in type order, and marks that run as one group; a bare
    template has none and its Detail stands alone."""
    assert top_level(template("form", "blob")) == [(OPEN_SECTION, 152, 152, "Detail")]

    form = blank.add_form("Defaults")
    form.add_control("TextBox", "B", top=240)
    blob = design_blobs(blank, "form")[0]
    assert top_level(blob) == [(OPEN_CONTROL, 109, 2, None), (OPEN_SIBLING, 152, 152, "Detail")]

    form.add_control("Label", "A", top=700, caption="A")
    form.add_control("TextBox", "C", top=1100)
    form.add_control("Tab", "T", top=1500, width=3000, height=1200)
    form.add_control("Page", "P", parent="T")
    blob = design_blobs(blank, "form")[0]
    assert top_level(blob) == [
        (OPEN_CONTROL, 100, 5, None),
        (OPEN_SIBLING, 109, 109, None),
        (OPEN_SIBLING, 123, 123, None),
        (OPEN_SIBLING, 124, 124, None),
        (OPEN_SIBLING, 152, 152, "Detail"),
    ]
    _header, objects, _trailer = parse_design(blob)
    defaults = [o for o in objects[1:] if not o.name and not o.is_section]
    assert [len(o.records) for o in defaults] == [len(prototype_records("form", o.type or 0)) for o in defaults]
    assert all(o.records == prototype_records("form", o.type or 0) for o in defaults)

    form.remove_control("C")
    form.remove_control("B")
    assert [t for _m, t, _c, _n in top_level(design_blobs(blank, "form")[0])] == [100, 109, 123, 124, 152]


def test_a_report_has_its_own_defaults_and_marks_its_sections_with_them(blank: AccessDatabase) -> None:
    report = blank.add_report("Sheet")
    report.add_control("TextBox", "Amount", top=240)
    report.add_control("Label", "Banner", section="PageHeaderSection", caption="B")
    blob = design_blobs(blank, "report")[0]
    assert top_level(blob) == [
        (OPEN_CONTROL, 100, 5, None),
        (OPEN_SIBLING, 109, 109, None),
        (OPEN_SIBLING, 155, 155, "PageHeaderSection"),
        (OPEN_SIBLING, 152, 152, "Detail"),
        (OPEN_SIBLING, 156, 156, "PageFooterSection"),
    ]
    assert prototype_records("report", 109) != prototype_records("form", 109)
    assert prototype_records("report", 100) == prototype_records("form", 100)


def test_the_defaults_are_the_ones_access_wrote(every_database: AccessDatabase) -> None:
    """Every form Access built carries the objects the template holds, for
    exactly the types it uses."""
    checked = 0
    for blob in design_blobs(every_database, "form"):
        _header, objects, _trailer = parse_design(blob)
        defaults = {o.type: o.records for o in objects[1:] if not o.name and not o.is_section}
        used = {o.type for o in objects[1:] if o.name and not o.is_section}
        # a page break is the one control type without a defaults object
        assert set(defaults) >= used - {118}
        assert all(records == prototype_records("form", t or 0) for t, records in defaults.items())
        checked += len(defaults)
    assert checked > 100


def test_a_type_without_captured_defaults_is_refused() -> None:
    with pytest.raises(AccessError, match="no control defaults were captured"):
        prototype_records("report", 128)


def test_a_colour_or_font_brings_the_records_access_writes_beside_it(blank: AccessDatabase) -> None:
    """Set in Access, a colour comes with its theme index at -1, a font
    with the theme font index at -1 and its family byte, and a command
    button's fill with the gradient off; a type without the companion's
    slot takes the property alone.  Measured one property at a time."""
    form = blank.add_form("Styled")
    form.add_control("Label", "Title", top=240, caption="x")
    form.add_control("CommandButton", "Go", top=700, width=1800, height=480, caption="Go")
    form.add_control("Line", "Rule", top=1300, width=3000, height=0)

    title = form.control("Title")
    title.set_property("FontName", "Consolas")
    title.set_property("ForeColor", 255)
    title.set_property("BackColor", 65535)
    named = {PROPERTY_CODES_BY_NUMBER.get(r.code, r.code): r.value for r in form.control("Title").object.records}
    assert named["ThemeFontIndex"] == bytes([255, 255, 255, 255])
    assert named["TextFontFamily"] == bytes([0x31])
    assert named["ForeThemeColorIndex"] == bytes([255, 255, 255, 255])
    assert named["BackThemeColorIndex"] == bytes([255, 255, 255, 255])
    assert "Gradient" not in named

    button = form.control("Go")
    button.set_property("UseTheme", False)
    for colour, value in (("Back", 0x79_4E_1F), ("Fore", 0xFFFFFF), ("Hover", 0xB6_75_2E), ("Pressed", 0x54_36_15), ("HoverFore", 0xFFFFFF), ("PressedFore", 0xFFFFFF), ("Border", 0x79_4E_1F)):
        button.set_property(f"{colour}Color", value)
    ids = {r.id: r for r in form.control("Go").object.records}
    assert {304, 314, 318, 319, 320, 326, 327, 331, 332, 335, 336, 339, 340, 343, 344} <= set(ids)
    assert ids[314].value == bytes([0]) and ids[318].value == bytes(4)
    assert ids[319].value == bytes([0x1F, 0x4E, 0x79, 0]) and ids[320].value == bytes([255, 255, 255, 255])
    assert ids[332].value == bytes([255, 255, 255, 255]) and ids[344].value == bytes([255, 255, 255, 255])

    rule = form.control("Rule")
    rule.set_property("BorderColor", 255)
    codes = {PROPERTY_CODES_BY_NUMBER.get(r.code, r.code) for r in form.control("Rule").object.records}
    assert "BorderThemeColorIndex" in codes and "BackThemeColorIndex" not in codes


def test_the_row_source_type_writes_its_companion_kind(blank: AccessDatabase) -> None:
    """The text is what the property sheet shows; Access acts on a one-byte
    record beside it: 1 for a value list, 10 for a field list, none for a
    table or query."""
    form = blank.add_form("Lists")
    box = form.add_control("ListBox", "Lines", top=240, height=900)

    def kind_record() -> bytes | None:
        return form.control("Lines").property_value(92)

    assert kind_record() is None
    box.set_property("RowSourceType", "Value List")
    assert kind_record() == b"\x01" and form.control("Lines").get("RowSourceType") == "Value List"
    box.set_property("RowSourceType", "Field List")
    assert kind_record() == b"\x0a"
    box.set_property("RowSourceType", "Table/Query")
    assert kind_record() is None



# --- renaming and deleting a design ------------------------------------------
# A design's name lives in four places: the container's listing, the
# catalog row, the navigation-pane row, and the module behind it, which
# Access binds by name.  PROJECT names that module only as a `DocClass=`
# line, and a stale one makes Access report the whole project as corrupt
# (GitHub issue #21).


def project_lines(database: AccessDatabase) -> list[str]:
    raw = next(
        row["Lv"]
        for _rid, row in database.table("MSysAccessStorage").rows_with_ids()
        if str(row["Name"]) == "PROJECT"
    )
    assert isinstance(raw, bytes)
    return raw.decode("cp1252").split(chr(13) + chr(10))


def test_renaming_a_form_carries_its_module(coded: AccessDatabase, tmp_path: Path) -> None:
    coded.rename_form("Coded", "Invoice")
    out = tmp_path / "renamed.accdb"
    coded.save(out)

    after = AccessDatabase(out)
    assert [form.name for form in after.forms()] == ["Invoice"]
    assert [m.name for m in after.modules()] == ["Module1", "Form_Invoice"]
    assert "Answer = 42" in after.module("Form_Invoice").source
    assert "DocClass=Form_Invoice/&H00000000" in project_lines(after)
    assert not [line for line in project_lines(after) if "Form_Coded" in line]


def test_renaming_a_form_reaches_the_catalog_and_the_nav_pane(
    coded: AccessDatabase, tmp_path: Path
) -> None:
    coded.rename_form("Coded", "Invoice")
    out = tmp_path / "catalog.accdb"
    coded.save(out)

    after = AccessDatabase(out)
    assert "Invoice" in [e.name for e in after.catalog() if e.type == -32768]
    assert "Coded" not in [e.name for e in after.catalog() if e.type == -32768]
    nav = [str(row["Name"]) for row in after.table("MSysNavPaneObjectIDs").rows()]
    assert "Invoice" in nav and "Coded" not in nav
    container = after._design_container("form")  # pyright: ignore[reportPrivateUsage]
    listing = next(
        row["Lv"]
        for _rid, row in after.table("MSysAccessStorage").rows_with_ids()
        if row["ParentId"] == container and str(row["Name"]) == DIR_DATA
    )
    assert isinstance(listing, bytes)
    assert [name for name, _folder in dir_data_entries(listing)] == ["Invoice"]


def test_renaming_a_form_that_has_no_code(blank: AccessDatabase, tmp_path: Path) -> None:
    """A design Access never opened a code window for has three places,
    not four, and the rename must not go looking for a module."""
    blank.create_form("Plain")
    blank.rename_form("Plain", "Renamed")
    out = tmp_path / "plain.accdb"
    blank.save(out)

    after = AccessDatabase(out)
    assert [form.name for form in after.forms()] == ["Renamed"]
    assert not [m for m in after.modules() if m.name.startswith("Form_")]


def test_renaming_a_report_works_the_same_way(blank: AccessDatabase, tmp_path: Path) -> None:
    blank.create_report("Monthly")
    blank.set_design_code(
        "Monthly", "Option Compare Database" + NEWLINE + "Public Sub Go()" + NEWLINE + "End Sub",
        kind="report",
    )
    blank.rename_report("Monthly", "Quarterly")
    out = tmp_path / "report.accdb"
    blank.save(out)

    after = AccessDatabase(out)
    assert [report.name for report in after.reports()] == ["Quarterly"]
    assert "Report_Quarterly" in [m.name for m in after.modules()]
    assert "DocClass=Report_Quarterly/&H00000000" in project_lines(after)


def test_renaming_a_design_refuses_a_name_already_taken(blank: AccessDatabase) -> None:
    blank.create_form("One")
    blank.create_form("Two")
    with pytest.raises(AccessError, match="already exists"):
        blank.rename_form("One", "Two")


def test_deleting_a_form_takes_its_module_with_it(
    coded: AccessDatabase, tmp_path: Path
) -> None:
    """The module used to survive the design, leaving PROJECT naming a
    DocClass whose form was gone."""
    coded.delete_form("Coded")
    out = tmp_path / "deleted.accdb"
    coded.save(out)

    after = AccessDatabase(out)
    assert after.forms() == []
    assert [m.name for m in after.modules()] == ["Module1"]
    assert not [line for line in project_lines(after) if "Form_Coded" in line]
    workspace = next(
        row["Lv"]
        for _rid, row in after.table("MSysAccessStorage").rows_with_ids()
        if str(row["Name"]) == "PROJECTwm"
    )
    assert isinstance(workspace, bytes)
    assert b"Form_Coded" not in workspace


def test_deleting_a_form_without_code_leaves_the_project_alone(
    blank: AccessDatabase, tmp_path: Path
) -> None:
    blank.create_form("Plain")
    before = project_lines(blank)
    blank.delete_form("Plain")
    out = tmp_path / "nocode.accdb"
    blank.save(out)

    after = AccessDatabase(out)
    assert after.forms() == []
    assert project_lines(after) == before
