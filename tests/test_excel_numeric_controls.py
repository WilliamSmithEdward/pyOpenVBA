"""Replay native Excel spinner and scroll-bar behavior."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/numeric_controls.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[f'{r["kind"]}-{r["name"]}' for r in RECORDS])
def test_measured_numeric_control(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, n As Long, p As String\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_numeric_values_and_bounds_persist(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
Dim sh As Object
Set sh = ActiveSheet.Shapes.AddFormControl(8, 0, 0, 90, 30)
sh.Name = "Scroll"
sh.ControlFormat.LinkedCell = "$H$1"
sh.ControlFormat.Min = 5
sh.ControlFormat.Max = 45
sh.ControlFormat.SmallChange = 2
sh.ControlFormat.LargeChange = 8
End Sub''', name="Builder")
    app.run("Build")
    app.sheet(1).set_control_value("Scroll", 13)
    path = tmp_path / "numeric.xlsm"
    app.save(path)
    sheet = ExcelApplication.open(path, with_vba=False).sheet(1)
    control = sheet.shape("Scroll").control
    assert control is not None
    assert (control.value, control.minimum, control.maximum, control.increment, control.page_change) == (13, 5, 45, 2, 8)
    assert sheet.value("H1") == 13
