"""Reading a worksheet's shapes, against what Excel reported.

``scripts/measure_shapes.py`` had Excel build one of every kind of
shape a sheet can hold, asked Excel what it then said about each, and
kept both the workbook and the answers.  This holds the reader to
those answers without needing Office.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.shapes import Shape
from pyopenvba.shapes._xlsx import SheetGrid, grid_of, read_controls, read_drawing, written

FIXTURES = Path(__file__).parent / "fixtures" / "shapes"
WORKBOOK = FIXTURES / "excel_shapes.xlsm"


def package() -> zipfile.ZipFile:
    return zipfile.ZipFile(WORKBOOK)


def drawing() -> str:
    return package().read("xl/drawings/drawing1.xml").decode("utf-8")


def sheet() -> str:
    return package().read("xl/worksheets/sheet1.xml").decode("utf-8")


def measured() -> list[dict[str, str]]:
    body = json.loads((FIXTURES / "excel_shapes.json").read_text(encoding="utf-8"))
    return list(body["shapes"])


def grid() -> SheetGrid:
    """The sheet's own columns and rows, which a control's place needs."""
    return grid_of(sheet())


def control_parts() -> dict[str, str]:
    """Each control's own part, by the relationship the sheet names."""
    package_file = package()
    rels = package_file.read("xl/worksheets/_rels/sheet1.xml.rels").decode("utf-8")
    out: dict[str, str] = {}
    for found in re.finditer(r'<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
        target = found.group(2)
        if "ctrlProp" not in target:
            continue
        part = "xl/" + target.lstrip("./").replace("../", "")
        out[found.group(1)] = package_file.read(part).decode("utf-8")
    return out


def by_name() -> dict[str, Shape]:
    return {one.name: one for one in read_drawing(drawing(), grid())}


def test_every_shape_excel_reported_is_read() -> None:
    assert sorted(by_name()) == sorted(one["name"] for one in measured())


@pytest.mark.parametrize("want", measured(), ids=lambda one: str(one["name"]))
def test_the_shape_is_where_excel_said(want: dict[str, str]) -> None:
    shape = by_name()[want["name"]]
    where = (shape.left, shape.top, shape.width, shape.height)
    assert where == pytest.approx(
        (float(want["left"]), float(want["top"]), float(want["width"]), float(want["height"])),
        abs=0.75,
    )


def test_the_kinds_are_what_excel_calls_them() -> None:
    kinds = {name: shape.mso_type for name, shape in by_name().items()}
    assert kinds == {one["name"]: int(one["type"]) for one in measured()}


def test_a_shapes_text_comes_back() -> None:
    shapes = by_name()
    assert shapes["Rect"].text == "Hello"
    assert shapes["Box"].text == "Two\nlines"
    assert shapes["Oval"].text == ""


def test_the_macro_on_a_drawing_shape_is_read() -> None:
    assert by_name()["Rect"].macro == "Clicked"
    assert by_name()["Oval"].macro == ""


def test_a_form_controls_macro_lives_on_the_sheet() -> None:
    """A control keeps its macro in the sheet, not in the drawing."""
    from pyopenvba.shapes._xlsx import control_macro

    button = by_name()["Button1"]
    assert button.kind == "formControl"
    assert button.macro == ""
    assert control_macro(sheet(), button.shape_id) == "[0]!Clicked"


def test_a_control_says_what_it_is_wired_to() -> None:
    controls = read_controls(sheet(), control_parts())
    check = by_name()["Check1"]
    assert controls[check.shape_id].linked_cell == "$H$1"
    drop = by_name()["Drop1"]
    assert controls[drop.shape_id].list_range == "$J$1:$J$3"


def test_a_group_holds_its_members() -> None:
    group = by_name()["Group1"]
    assert group.kind == "group"
    assert [one.name for one in group.children] == ["Grouped1", "Grouped2"]


def test_the_geometry_is_the_preset_excel_wrote() -> None:
    shapes = by_name()
    assert shapes["Rect"].geometry == "rect"
    assert shapes["Rounded"].geometry == "roundRect"
    assert shapes["Oval"].geometry == "ellipse"
    assert shapes["Line1"].geometry == "line"


def test_writing_back_an_untouched_drawing_changes_nothing() -> None:
    """A drawing nobody edited has to survive a save exactly."""
    original = drawing()
    assert written(read_drawing(original), original) == original


def test_moving_one_shape_leaves_the_others_alone() -> None:
    original = drawing()
    shapes = read_drawing(original)
    moved = next(one for one in shapes if one.name == "Oval")
    moved.left, moved.top = 300.0, 400.0
    out = written(shapes, original)
    assert out != original
    again = {one.name: one for one in read_drawing(out)}
    assert (again["Oval"].left, again["Oval"].top) == (300.0, 400.0)
    assert (again["Rect"].left, again["Rect"].top) == (10.0, 20.0)
    assert again["Rect"].macro == "Clicked"


def test_the_macro_can_be_set_and_cleared() -> None:
    original = drawing()
    shapes = read_drawing(original)
    oval = next(one for one in shapes if one.name == "Oval")
    rect = next(one for one in shapes if one.name == "Rect")
    oval.macro = "Other"
    rect.macro = ""
    again = {one.name: one for one in read_drawing(written(shapes, original))}
    assert again["Oval"].macro == "Other"
    assert again["Rect"].macro == ""


def test_a_new_shape_is_written_and_read_back() -> None:
    original = drawing()
    shapes = read_drawing(original)
    shapes.append(
        Shape(
            name="Added",
            kind="shape",
            geometry="ellipse",
            left=11.0,
            top=22.0,
            width=33.0,
            height=44.0,
            text="New",
            macro="Clicked",
            shape_id=99,
        )
    )
    again = {one.name: one for one in read_drawing(written(shapes, original))}
    added = again["Added"]
    assert (added.left, added.top, added.width, added.height) == (11.0, 22.0, 33.0, 44.0)
    assert (added.text, added.macro, added.geometry) == ("New", "Clicked", "ellipse")


def test_a_deleted_shape_is_gone() -> None:
    original = drawing()
    shapes = [one for one in read_drawing(original) if one.name != "Rounded"]
    assert "Rounded" not in {one.name for one in read_drawing(written(shapes, original))}
