"""Replay measured Forms option-button group behavior."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/radio_controls.json").read_text())
DELETIONS = json.loads((Path(__file__).parent / "fixtures/shapes/radio_deletion.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[f'{r["boxed"]}-{r["name"]}' for r in RECORDS])
def test_measured_radio_control(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim a As Object, b As Object, c As Object, n As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_radio_groups_persist(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_form_control(4, width=150, height=100)
    sheet.add_form_control(7, name="First", left=10, top=10, width=90, height=20)
    sheet.add_form_control(7, name="Second", left=10, top=40, width=90, height=20)
    sheet.add_form_control(7, name="Outside", left=200, top=10, width=90, height=20)
    sheet.update_control("Second", linked_cell="$H$1")
    sheet.set_control_value("Second", 1)
    sheet.set_control_value("Outside", 1)
    path = tmp_path / "radios.xlsm"
    app.save(path)
    sheet = ExcelApplication.open(path, with_vba=False).sheet(1)
    sheet.set_value("H1", 1)
    assert [s.control.value for s in sheet.shapes() if s.control and s.control.kind == "Radio"] == [1, -4146, 1]
    sheet.set_control_value("Second", 0)
    assert sheet.value("H1") == 1


@pytest.mark.parametrize("record", DELETIONS,
                         ids=[f'{r["boxed"]}-{r["selected"]}-{r["deleted"]}' for r in DELETIONS])
def test_measured_radio_deletion(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, i As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_late_identical_box_preserves_a_group_without_unboxed_radios() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.add_form_control(4, name="Box", width=150, height=100)
    sheet.add_form_control(7, name="Radio", left=10, top=10, width=90, height=20)
    sheet.update_control("Radio", linked_cell="$H$1")
    sheet.add_form_control(4, width=150, height=100)
    sheet.set_control_value("Radio", 1)
    assert sheet.value("H1") == 1
    assert len(sheet.shapes()) == 3
