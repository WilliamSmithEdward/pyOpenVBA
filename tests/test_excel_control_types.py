"""Issue #25: replay Excel-authored controls without losing values or types.

Fixtures are measurement data from pyOfficeEditor, copied as requested in
the issue. The libraries' implementations remain independent.
"""
from __future__ import annotations

import json
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.shapes import Shape
from pyopenvba.shapes._values import ControlInfo
from pyopenvba.shapes._xlsx import control_properties, control_vml, read_controls

FOLDER = Path(__file__).parent / "fixtures" / "shapes"


def test_all_control_values_match_excel_measurements() -> None:
    answers = json.loads((FOLDER / "all_controls_answers.json").read_text(encoding="utf-8"))
    app = ExcelApplication.open(FOLDER / "all_controls.xlsm", with_vba=False)
    for name, answer in answers.items():
        if isinstance(answer["value"], int):
            control = app.sheet(1).shape(name).control
            assert control is not None and control.value == answer["value"], name


def test_regenerated_properties_preserve_all_measured_control_types_and_values() -> None:
    app = ExcelApplication.open(FOLDER / "all_controls.xlsm", with_vba=False)
    for shape in app.sheet(1).shapes():
        if shape.control is None:
            continue
        properties = control_properties(shape)
        root = ET.fromstring(properties)
        assert root.attrib["objectType"] == shape.control.kind, shape.name
        if shape.control.kind in {"CheckBox", "Radio"} and shape.control.value == -4146:
            assert "checked" not in root.attrib
        copied = read_controls('<control shapeId="1" r:id="rId1"/>', {"rId1": properties})[1]
        assert copied.value == shape.control.value, shape.name
        if copied.kind in {"Spin", "Scroll"}:
            assert (copied.minimum, copied.maximum, copied.increment, copied.page_change) == (
                shape.control.minimum, shape.control.maximum, shape.control.increment, shape.control.page_change,
            )


@pytest.mark.parametrize("kind,maximum", [("Spin", 30000), ("Spinner", 30000),
                                         ("Scroll", 100), ("ScrollBar", 100)])
def test_new_numeric_controls_have_measured_defaults(kind: str, maximum: int) -> None:
    shape = Shape(name="Numeric", kind="formControl", shape_id=1025, control=ControlInfo(kind=kind, value=7))
    root = ET.fromstring(control_properties(shape))
    assert root.attrib["objectType"] == ("Spin" if maximum == 30000 else "Scroll")
    assert {name: root.attrib[name] for name in ("min", "max", "inc", "page", "dx", "val")} == {
        "min": "0", "max": str(maximum), "inc": "1", "page": "10", "dx": "31", "val": "7",
    }
    vml = control_vml(shape)
    assert f"<x:Max>{maximum}</x:Max>" in vml
    assert "<x:Val>7</x:Val>" in vml


def test_custom_numeric_bounds_survive_regeneration() -> None:
    properties = '<formControlPr objectType="Spin" min="5" max="45" inc="2" page="8" dx="24" val="13"/>'
    control = read_controls('<control shapeId="1" r:id="rId1"/>', {"rId1": properties})[1]
    shape = Shape(name="Custom", kind="formControl", control=control)
    root = ET.fromstring(control_properties(shape))
    assert {key: root.attrib[key] for key in ("min", "max", "inc", "page", "dx", "val")} == {
        "min": "5", "max": "45", "inc": "2", "page": "8", "dx": "24", "val": "13",
    }


@pytest.mark.parametrize("number,kind", [(0, "Button"), (1, "CheckBox"), (2, "Drop"),
                                       (4, "GBox"), (5, "Label"), (6, "List"),
                                       (7, "Radio"), (8, "Scroll"), (9, "Spin")])
def test_python_creates_form_control(tmp_path: Path, number: int, kind: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    made = sheet.add_form_control(number, name="Created", text="Caption", left=12, top=20)
    assert made.control is not None and made.control.kind == kind
    made.control.value = 999
    path = tmp_path / "control.xlsm"
    app.save(path)
    copy = ExcelApplication.open(path, with_vba=False).sheet(1).shape("Created")
    assert copy.control is not None and copy.control.kind == kind and copy.control.value != 999
    assert (copy.left, copy.top, copy.text) == (12, 20, "Caption")


def test_python_control_creation_validates_before_mutation() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_form_control(1, name="Taken")
    with pytest.raises(ValueError):
        sheet.add_form_control(3)
    with pytest.raises(ValueError):
        sheet.add_form_control(2, name="taken")
    with pytest.raises(ValueError):
        sheet.add_form_control(6, text="bad\x00")
    assert len(sheet.shapes()) == 1
