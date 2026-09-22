"""Native Excel named control binding behavior."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/shapes/control_names.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_measured_control_names(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim sh As Object, n As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_local_named_control_bindings_persist(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
ActiveWorkbook.Names.Add "Target", "=Sheet1!$H$1"
ActiveWorkbook.Names.Add "Sheet1!Target", "=Sheet1!$H$2"
ActiveWorkbook.Names.Add "Choices", "=Sheet1!$J$1:$J$3"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    sheet.add_form_control(6, name="ChoicesControl")
    sheet.update_control("ChoicesControl", linked_cell="target", list_range="=Choices")
    sheet.set_control_value("ChoicesControl", 2)
    path = tmp_path / "named.xlsm"
    app.save(path)
    sheet = ExcelApplication.open(path, with_vba=False).sheet(1)
    sheet.set_control_value("ChoicesControl", 3)
    assert sheet.value("H1") is None
    assert sheet.value("H2") == 3
    assert len(sheet.control_items("ChoicesControl")) == 3


def test_unsupported_named_formulas_and_cycles_do_not_mutate() -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Sub Build()
ActiveWorkbook.Names.Add "Dynamic", "=XLOOKUP(1,Sheet1!$A$1:$A$2,Sheet1!$B$1:$B$2)"
ActiveWorkbook.Names.Add "LoopOne", "=LoopTwo"
ActiveWorkbook.Names.Add "LoopTwo", "=LoopOne"
End Sub''', name="Builder")
    app.run("Build")
    sheet = app.sheet(1)
    sheet.add_form_control(6, name="Choices")
    before = sheet.shape("Choices")
    for name in ("Dynamic", "LoopOne", "[Other.xlsx]Sheet1!A1"):
        with pytest.raises(VBAUnsupportedError):
            sheet.update_control("Choices", linked_cell="$H$1", list_range=name)
        assert sheet.shape("Choices") == before
