"""Dropdown/list selection semantics measured in Excel."""
from __future__ import annotations

import json
from pathlib import Path
from typing import cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.shapes._xlsx import read_controls

FOLDER = Path(__file__).parent / "fixtures" / "shapes"
RECORDS = json.loads((FOLDER / "list_controls.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("record", RECORDS, ids=[f'{one["kind"]}-{one["name"]}' for one in RECORDS])
def test_list_control_measurements(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = ('Public Function Report() As String\nDim sh As Object, n As Long\n'
              + str(record["body"]) + 'End Function\n')
    app.add_module(source, name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("record", RECORDS, ids=[f'{one["kind"]}-{one["name"]}' for one in RECORDS])
def test_saved_excel_selection_reader(record: dict[str, object]) -> None:
    parts = record["parts"]
    assert isinstance(parts, dict)
    properties = cast("dict[str, str]", parts)["xl/ctrlProps/ctrlProp1.xml"]
    controls = read_controls('<control shapeId="1" r:id="rId1"/>', {"rId1": properties})
    assert controls[1].value == int(str(record["reported"]).split("|")[1])


def test_python_selection_round_trip_and_clear(tmp_path: Path) -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.update_control("Drop1", linked_cell="$H$2")
    for value in (2, 0, 3):
        control = sheet.set_control_value("Drop1", value).control
        assert control is not None and control.value == value
        assert sheet.value("H2") == value
        out = tmp_path / "selected.xlsm"
        app.save(out)
        reopened = ExcelApplication.open(out, with_vba=False).sheet(1)
        control = reopened.shape("Drop1").control
        assert control is not None and control.value == value
        assert reopened.value("H2") == value


def test_invalid_selection_leaves_control_and_linked_cell_unchanged() -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.update_control("Drop1", linked_cell="$H$2")
    sheet.set_control_value("Drop1", 2)
    with pytest.raises(ValueError):
        sheet.set_control_value("Drop1", 4)
    assert sheet.value("H2") == 2
    control = sheet.shape("Drop1").control
    assert control is not None and control.value == 2


def test_multi_selection_reader_recognizes_excel_attribute() -> None:
    properties = '<formControlPr objectType="List" seltype="multi" multiSel="3, 1" sel="0" val="0"/>'
    control = read_controls('<control shapeId="1" r:id="rId1"/>', {"rId1": properties})[1]
    assert control.selection_mode == "multi"
    assert sorted(control.selected_indices) == [1, 3]
