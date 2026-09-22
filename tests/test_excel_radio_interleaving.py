"""Measured non-adjacent radio groups and multi-box deletion."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/radio_interleaving.json").read_text())


@pytest.mark.parametrize("record", RECORDS,
                         ids=[f'{r["boxes"]}-{r["order"]}-{r["operation"]}' for r in RECORDS])
def test_measured_radio_interleaving(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, box1 As Object, box2 As Object\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_moved_radio_membership_is_reconstructed_on_open() -> None:
    # Native Excel: moving Second out of its box keeps its old group during
    # editing, but reopening puts Second and Third in the unboxed group.
    app = ExcelApplication.open(Path(__file__).parent / "fixtures/shapes/radios_moved.xlsm", with_vba=False)
    app.add_module('''Public Function Links() As String
Links = ActiveSheet.Shapes("First").ControlFormat.LinkedCell & "|" & _
ActiveSheet.Shapes("Second").ControlFormat.LinkedCell & "|" & _
ActiveSheet.Shapes("Third").ControlFormat.LinkedCell
End Function''', name="Inspector")
    assert app.run("Links") == "$H$1||"
    sheet = app.sheet(1)
    sheet.set_control_value("Third", 1)
    control = sheet.shape("Second").control
    assert control is not None and control.value == -4146
