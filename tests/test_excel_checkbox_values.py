"""Replay control behavior recorded from live Excel."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FOLDER = Path(__file__).parent / "fixtures" / "shapes"
RECORDS = json.loads((FOLDER / "checkbox_links.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("record", RECORDS, ids=[item["name"] for item in RECORDS])
def test_measured_checkbox_links(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    source = (
        'Public Function Report() As String\nDim sh As Object, n As Long\n'
        'Set sh = ActiveSheet.Shapes.AddFormControl(1, 0, 0, 90, 20)\n'
        'sh.ControlFormat.LinkedCell = "$H$1"\nOn Error Resume Next\n'
        + record["statement"] + '\nn = Err.Number\nOn Error GoTo 0\n'
        'Report = CStr(n) & "|" & CStr(sh.ControlFormat.Value) & "|" & _\n'
        'TypeName(Range("H1").Value) & "|" & CStr(Range("H1").Value)\nEnd Function\n'
    )
    app.add_module(source, name="Probe")
    assert app.run("Report") == record["reported"]


@pytest.mark.parametrize("value", [-4146, 1, 2])
def test_python_checkbox_values_persist(tmp_path: Path, value: int) -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.set_control_value("Check1", 1)
    sheet.set_control_value("Check1", value)
    path = tmp_path / "checkbox.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path, with_vba=False)
    control = reopened.sheet(1).shape("Check1").control
    assert control is not None and control.value == value
    expected = "Error 2042" if value == 2 else str(value == 1)
    assert reopened.evaluate('CStr(Range("H1").Value)') == expected


def test_invalid_value_is_atomic_and_unsupported_controls_fail() -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    original = sheet.shape("Check1")
    with pytest.raises(ValueError):
        sheet.set_control_value("Check1", 3)
    assert sheet.shape("Check1") == original
    with pytest.raises(VBAUnsupportedError):
        sheet.set_control_value("Button1", 1)


def test_formula_link_refreshes_python_snapshot() -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.set_value("H1", "=G1>0")
    sheet.set_value("G1", 1)
    control = sheet.shape("Check1").control
    assert control is not None and control.value == 1
    sheet.set_value("G1", 0)
    control = sheet.shape("Check1").control
    assert control is not None and control.value == -4146


def test_formula_link_updates_when_saved_without_reading_snapshot(tmp_path: Path) -> None:
    app = ExcelApplication.open(FOLDER / "excel_shapes.xlsm", with_vba=False)
    sheet = app.sheet(1)
    sheet.set_value("H1", "=G1>0")
    sheet.set_value("G1", 1)
    path = tmp_path / "calculated.xlsm"
    app.save(path)
    control = ExcelApplication.open(path, with_vba=False).sheet(1).shape("Check1").control
    assert control is not None and control.value == 1


def test_cross_sheet_and_shared_cell_bindings() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Setup()
Dim sh As Object
Worksheets.Add.Name = "Data"
Set sh = Worksheets(2).Shapes.AddFormControl(1, 0, 0, 90, 20)
sh.Name = "First"
sh.ControlFormat.LinkedCell = "Data!H1"
Set sh = Worksheets(2).Shapes.AddFormControl(1, 0, 40, 90, 20)
sh.Name = "Second"
sh.ControlFormat.LinkedCell = "Data!H1"
End Sub''', name="SetupModule")
    app.run("Setup")
    sheet = app.sheet(2)
    sheet.set_control_value("First", 1)
    control = sheet.shape("Second").control
    assert control is not None and control.value == 1
    app.sheet("Data").set_value("H1", False)
    for name in ("First", "Second"):
        control = sheet.shape(name).control
        assert control is not None and control.value == -4146
