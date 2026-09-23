"""The supported Python API shares shape state with VBA and persists every control part."""

from __future__ import annotations

import re
import json
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE = Path(__file__).parent / "fixtures" / "shapes" / "excel_shapes.xlsm"


def opened() -> ExcelApplication:
    return ExcelApplication.open(FIXTURE, with_vba=False)


def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def test_read_controls_and_captions_without_importing_private_modules() -> None:
    sheet = opened().sheet(1)
    assert len(sheet.shapes()) == 9
    assert sheet.shape("button1").text == "Press"
    assert sheet.shape("Oval").auto_shape_type == 9
    check = sheet.shape("Check1").control
    drop = sheet.shape("Drop1").control
    assert check is not None and check.linked_cell == "$H$1"
    assert check.value == -4146
    assert drop is not None and drop.list_range == "$J$1:$J$3"
    assert drop.value == 0


def test_checkbox_values_match_live_excel_measurements() -> None:
    from pyopenvba.shapes._xlsx import read_controls

    measured = json.loads((FIXTURE.parent / "checkbox_values.json").read_text(encoding="utf-8"))
    for state in measured:
        xml = '<controls><control shapeId="1" r:id="rId1"/></controls>'
        control = read_controls(xml, {"rId1": state["properties"]})[1]
        assert control.value == int(state["reported"])


def test_control_bindings_round_trip_and_clear(tmp_path: Path) -> None:
    app = opened()
    sheet = app.sheet(1)
    sheet.update_control("Check1", linked_cell="$K$2")
    sheet.update_control("Drop1", linked_cell="$K$3", list_range="$J$2:$J$3")
    out = tmp_path / "bindings.xlsm"
    app.save(out)
    reopened = ExcelApplication.open(out, with_vba=False).sheet(1)
    check = reopened.shape("Check1").control
    drop = reopened.shape("Drop1").control
    assert check is not None and check.linked_cell == "$K$2"
    assert drop is not None and (drop.linked_cell, drop.list_range) == ("$K$3", "$J$2:$J$3")
    saved = parts(out)
    assert b"<x:FmlaLink>$K$2</x:FmlaLink>" in saved["xl/drawings/vmlDrawing1.vml"]
    assert b'noThreeD="1"' in saved["xl/ctrlProps/ctrlProp2.xml"]
    sheet.update_control("Check1", linked_cell="")
    sheet.update_control("Drop1", linked_cell="", list_range="")
    app.save(out)
    cleared = ExcelApplication.open(out, with_vba=False).sheet(1)
    check = cleared.shape("Check1").control
    drop = cleared.shape("Drop1").control
    assert check is not None and check.linked_cell == ""
    assert drop is not None and (drop.linked_cell, drop.list_range) == ("", "")
    assert b"<x:FmlaLink>" not in parts(out)["xl/drawings/vmlDrawing1.vml"]


@pytest.mark.parametrize("name,link,source", [
    ("Rect", "$K$1", None), ("Check1", "$K$1", "$J$1:$J$3"), ("Drop1", "$K$1", "Missing!A1:A2"),
    ("Drop1", "$K$1", "A1\x00"), ("Button1", "$K$1", None),
])
def test_invalid_control_bindings_are_atomic(name: str, link: str, source: str | None) -> None:
    sheet = opened().sheet(1)
    before = sheet.shape(name)
    with pytest.raises((ValueError, VBAUnsupportedError)):
        sheet.update_control(name, linked_cell=link, list_range=source)
    assert sheet.shape(name) == before


def test_snapshots_do_not_mutate_the_workbook_or_nested_controls(tmp_path: Path) -> None:
    app = opened()
    sheet = app.sheet(1)
    snapshot = sheet.shape("Check1")
    snapshot.name = "Not an edit"
    assert snapshot.control is not None
    snapshot.control.linked_cell = "Z99"
    group = sheet.shape("Group1")
    group.children[0].name = "Not an edit either"
    assert sheet.shape("Check1").control != snapshot.control
    assert sheet.shape("Group1").children[0].name != group.children[0].name
    out = tmp_path / "untouched.xlsm"
    app.save(out)
    assert out.read_bytes() == FIXTURE.read_bytes()


@pytest.mark.parametrize("kind", ["shape", "textBox", "formControl"])
def test_create_edit_save_reopen_and_vba_share_one_shape(kind: str, tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    creator = {"shape": sheet.add_shape, "textBox": sheet.add_textbox, "formControl": sheet.add_button}[kind]
    made = creator(name="Start", left=12, top=30, width=90, height=40, text="First", macro="Clicked")
    assert made.kind == kind
    assert app.evaluate('ActiveSheet.Shapes("Start").OnAction') == "Clicked"
    first = tmp_path / "first.xlsm"
    app.save(first)
    sheet.update_shape("start", new_name=r"Run & \1", left=144, top=88, width=123, height=45,
                       text="Next & <end>", macro="Other")
    assert app.evaluate('ActiveSheet.Shapes(1).Left') == 144
    second = tmp_path / "second.xlsm"
    app.save(second)
    again = ExcelApplication.open(second, with_vba=False).sheet(1).shape(r"Run & \1")
    assert (again.left, again.top, again.width, again.height) == (144, 88, 123, 45)
    assert (again.text, again.macro) == ("Next & <end>", "Other")
    # Revert values after a save: persistence must compare against the
    # latest saved state, including the VML caption.
    sheet.update_shape(r"Run & \1", text="First", macro="")
    app.save(second)
    final = ExcelApplication.open(second, with_vba=False).sheet(1).shape(r"Run & \1")
    assert final.text == "First" and final.macro == ""


def test_control_edits_preserve_sibling_controls_and_unrelated_parts(tmp_path: Path) -> None:
    app = opened()
    app.sheet(1).update_shape("Button1", new_name="Run", left=25, top=50, width=105, height=33,
                              text="Changed\nCaption", macro="")
    out = tmp_path / "edited.xlsm"
    app.save(out)
    before, after = parts(FIXTURE), parts(out)
    changed = {name for name in before if before[name] != after[name]}
    assert changed == {"xl/worksheets/sheet1.xml", "xl/drawings/drawing1.xml", "xl/drawings/vmlDrawing1.vml"}
    vml_before = before["xl/drawings/vmlDrawing1.vml"].decode()
    vml_after = after["xl/drawings/vmlDrawing1.vml"].decode()
    for name in ("Check1", "Drop1"):
        pattern = rf'<v:shape id="{name}".*?</v:shape>'
        old, new = re.search(pattern, vml_before, re.S), re.search(pattern, vml_after, re.S)
        assert old is not None and new is not None
        assert old.group() == new.group()
    assert b"Changed" in after["xl/drawings/drawing1.xml"]
    assert b"Caption" in after["xl/drawings/drawing1.xml"]
    sheet = ExcelApplication.open(out, with_vba=False).sheet(1)
    assert sheet.shape("Run").text == "Changed\nCaption"
    assert sheet.shape("Run").macro == ""
    assert sheet.shape("Check1").control == app.sheet(1).shape("Check1").control


def test_delete_control_removes_all_records_and_preserves_others(tmp_path: Path) -> None:
    app = opened()
    sheet = app.sheet(1)
    control = sheet.shape("Check1").control
    assert control is not None
    sheet.remove_shape("Check1")
    out = tmp_path / "deleted.xlsm"
    app.save(out)
    data = parts(out)
    assert control.part_name not in data
    assert b'Check1' not in data["xl/worksheets/sheet1.xml"]
    assert b'_x0000_s1026' not in data["xl/drawings/vmlDrawing1.vml"]
    rels = ET.fromstring(data["xl/worksheets/_rels/sheet1.xml.rels"])
    assert all(entry.get("Id") != control.relationship for entry in rels)
    assert control.part_name.encode() not in data["[Content_Types].xml"]
    again = ExcelApplication.open(out, with_vba=False).sheet(1)
    assert len(again.shapes()) == 8
    assert again.shape("Button1").text == "Press"
    for name in ("xl/ctrlProps/ctrlProp1.xml", "xl/ctrlProps/ctrlProp3.xml", "xl/vbaProject.bin"):
        assert data[name] == parts(FIXTURE)[name]


def test_delete_last_control_then_create_another_and_save_again(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_button(name="Old", text="Old")
    out = tmp_path / "cycles.xlsm"
    app.save(out)
    sheet.remove_shape("Old")
    app.save(out)
    data = parts(out)
    assert not any(name.startswith("xl/ctrlProps/") or name.endswith(".vml") for name in data)
    assert b"legacyDrawing" not in data["xl/worksheets/sheet1.xml"]
    assert b"<controls>" not in data["xl/worksheets/sheet1.xml"]
    for name, body in data.items():
        if name.endswith((".xml", ".rels")):
            ET.fromstring(body)
    sheet.add_button(name="New", text="New")
    app.save(out)
    assert ExcelApplication.open(out, with_vba=False).sheet(1).shape("New").text == "New"


def test_replacing_a_deleted_control_before_saving_removes_its_old_parts(tmp_path: Path) -> None:
    app = opened()
    sheet = app.sheet(1)
    sheet.remove_shape("Drop1")
    sheet.add_button(name="Replacement", text="New")
    out = tmp_path / "replacement.xlsm"
    app.save(out)
    assert b"Drop1" not in parts(out)["xl/worksheets/sheet1.xml"]
    assert len(ExcelApplication.open(out, with_vba=False).sheet(1).shapes()) == 9


@pytest.mark.parametrize("bad", [-1.0, float("nan"), float("inf")])
def test_bad_geometry_refuses_before_mutating_state(bad: float) -> None:
    app = opened()
    sheet = app.sheet(1)
    before = sheet.shapes()
    with pytest.raises(ValueError):
        sheet.update_shape("Rect", new_name="Changed", left=bad)
    with pytest.raises(ValueError):
        sheet.add_button(name="Changed", width=bad)
    assert sheet.shapes() == before
    assert app.workbook.saved


def test_duplicate_names_unknown_types_and_missing_shapes_fail_cleanly() -> None:
    sheet = opened().sheet(1)
    with pytest.raises(ValueError):
        sheet.add_shape(name="RECT")
    with pytest.raises(ValueError):
        sheet.update_shape("Rect", new_name="Oval")
    with pytest.raises(VBAUnsupportedError):
        sheet.add_shape(9999)
    with pytest.raises(KeyError):
        sheet.remove_shape("Missing")
    with pytest.raises(ValueError):
        sheet.update_shape("Rect", new_name="Changed", text="bad\x00text")
    assert sheet.shape("Rect").text == "Hello"
    assert len(sheet.shapes()) == 9


def test_automatic_names_do_not_collide_with_user_names() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_shape(name="Rectangle 2")
    assert sheet.add_shape().name == "Rectangle 2 (2)"


def test_deleting_controls_keeps_other_vml_shapes(tmp_path: Path) -> None:
    data = parts(FIXTURE)
    vml_name = "xl/drawings/vmlDrawing1.vml"
    note = b'<v:shape id="Note" o:spid="_x0000_s4096"><x:ClientData ObjectType="Note"/></v:shape>'
    data[vml_name] = data[vml_name].replace(b"</xml>", note + b"</xml>")
    source = tmp_path / "with_note.xlsm"
    with zipfile.ZipFile(source, "w") as package:
        for name, body in data.items():
            package.writestr(name, body)
    app = ExcelApplication.open(source, with_vba=False)
    for name in ("Button1", "Check1", "Drop1"):
        app.sheet(1).remove_shape(name)
    out = tmp_path / "note_survives.xlsm"
    app.save(out)
    after = parts(out)
    assert note in after[vml_name]
    assert b"legacyDrawing" in after["xl/worksheets/sheet1.xml"]
    assert not any(name.startswith("xl/ctrlProps/") for name in after)


def test_multiline_button_text_survives_creation(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).add_button(name="Lines", text="First\nSecond & third")
    out = tmp_path / "lines.xlsm"
    app.save(out)
    assert ExcelApplication.open(out, with_vba=False).sheet(1).shape("Lines").text == "First\nSecond & third"


def test_setting_macro_on_a_drawing_without_the_attribute(tmp_path: Path) -> None:
    data = parts(FIXTURE)
    drawing = "xl/drawings/drawing1.xml"
    data[drawing] = re.sub(rb' macro="[^"]*"', b"", data[drawing])
    source = tmp_path / "no_macro_attributes.xlsm"
    with zipfile.ZipFile(source, "w") as package:
        for name, body in data.items():
            package.writestr(name, body)
    app = ExcelApplication.open(source, with_vba=False)
    app.sheet(1).update_shape("Rect", macro="Clicked")
    out = tmp_path / "assigned.xlsm"
    app.save(out)
    assert ExcelApplication.open(out, with_vba=False).sheet(1).shape("Rect").macro == "Clicked"


def test_deleting_a_control_keeps_a_property_part_referenced_elsewhere(tmp_path: Path) -> None:
    data = parts(FIXTURE)
    extra = (b'<Relationship Id="sharedControl" Type="https://example.test/shared" '
             b'Target="/xl/ctrlProps/ctrlProp2.xml"/>')
    data["_rels/.rels"] = data["_rels/.rels"].replace(b"</Relationships>", extra + b"</Relationships>")
    source = tmp_path / "shared_control_part.xlsm"
    with zipfile.ZipFile(source, "w") as package:
        for name, body in data.items():
            package.writestr(name, body)
    app = ExcelApplication.open(source, with_vba=False)
    app.sheet(1).remove_shape("Check1")
    out = tmp_path / "retained_part.xlsm"
    app.save(out)
    after = parts(out)
    assert after["xl/ctrlProps/ctrlProp2.xml"] == data["xl/ctrlProps/ctrlProp2.xml"]
    assert b"/xl/ctrlProps/ctrlProp2.xml" in after["[Content_Types].xml"]
