"""Replay Excel's multi/extended list binding and mode behavior."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/list_rebinding.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[f'{r["mode"]}-{r["name"]}' for r in RECORDS])
def test_measured_list_rebinding(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, n As Long, i As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_python_source_and_mode_changes_preserve_valid_selections(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
ActiveSheet.Shapes.AddFormControl(6, 0, 0, 90, 60).Name = "Choices"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    sheet.set_control_items("Choices", ["a", "b", "c"])
    sheet.set_control_selection_mode("Choices", 2)
    sheet.set_control_selection("Choices", [1, 3])
    sheet.set_control_selection_mode("Choices", 3)
    sheet.update_control("Choices", list_range="$J$1:$J$2")
    path = tmp_path / "rebound.xlsm"
    app.save(path)
    control = ExcelApplication.open(path, with_vba=False).sheet(1).shape("Choices").control
    assert control is not None
    assert control.selected_indices == [1]
    assert control.selection_mode == "extended"
    assert control.list_range == "$J$1:$J$2"
