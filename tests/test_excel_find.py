"""Range.Find cases measured in native Excel."""
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORDS = json.loads((Path(__file__).parent / "fixtures/range_find.json").read_text())


@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_measured_find(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Report() As String\nDim hit As Object, n As Long\n'
                   + str(record["body"]) + 'End Function\n', name="Probe")
    assert app.run("Report") == record["reported"]


def test_find_whole_sheet_does_not_materialize_cells() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Function Report() As String
Dim hit As Object
Range("XFD1048576").Value = "Last"
Set hit = Cells.Find("Last", LookAt:=xlWhole)
Report = hit.Address
Set hit = Cells.Find("", After:=hit)
Report = Report & "|" & hit.Address
End Function''', name="Probe")
    assert app.run("Report") == "$XFD$1048576|$A$1"
    book = app.application.active_book
    assert book is not None
    assert len(book.sheets_[0].cells_) == 1


def test_find_saved_settings_belong_to_application() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Public Function Report() As String
Dim hit As Object
Set hit = Range("A1").Find("none", LookAt:=xlWhole)
Workbooks.Add
Range("A1").Value = "alphabet"
Set hit = Range("A1").Find("alpha")
Report = CStr(hit Is Nothing)
End Function''', name="Probe")
    assert app.run("Report") == "True"


def test_find_after_reopen(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    report = '''Public Function Report() As String
Dim hit As Object
Set hit = Range("A1:C3").Find("50%", LookIn:=xlValues)
Report = hit.Address
End Function'''
    app.add_module(report, name="Probe")
    app.sheet(1).set_value("C3", 0.5)
    app.add_module('Sub FormatCell()\nRange("C3").NumberFormat = "0%"\nEnd Sub', name="Setup")
    app.run("FormatCell")
    path = tmp_path / "find.xlsm"
    app.save(path)
    reopened = ExcelApplication.open(path)
    reopened.add_module(report, name="Probe")
    assert reopened.run("Report") == "$C$3"


@pytest.mark.parametrize("expression", [
    'Range("A1,B2").Find("a")',
    'Range("A1").Find("a", SearchFormat:=True)',
    'Range("A1").Find("a", MatchByte:=True)',
    'Range("A1").Find("a", LookIn:=-4144)',
])
def test_unimplemented_find_options_are_explicit(expression: str) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Sub Probe()\nDim hit As Object\nSet hit = {expression}\nEnd Sub', name="ProbeModule")
    with pytest.raises(VBAUnsupportedError):
        app.run("Probe")
