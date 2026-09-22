"""Multi-workbook copying and lifecycle semantics."""
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORDS = json.loads((Path(__file__).parent / "fixtures/workbook_copy.json").read_text())

@pytest.mark.parametrize("record", RECORDS, ids=[r["name"] for r in RECORDS])
def test_native_workbook_copy(record: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim src As Object, dst As Object, sh As Object, cell As Object, nm As Object\n'
                   + str(record["body"]) + 'End Function', name="Probe")
    if record["name"] == "range_qualified":
        with pytest.raises(VBAUnsupportedError, match="external-link"):
            app.run("Report")
    else:
        assert app.run("Report") == record["reported"]


def test_workbook_names_are_not_reused_and_thisworkbook_stays_bound() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Report() As String
Dim first As Object, second As Object
Set first = Workbooks.Add
first.Close SaveChanges:=False
Set second = Workbooks.Add
Report = ThisWorkbook.Name & "|" & second.Name & "|" & ActiveWorkbook.Name & "|" & CStr(Workbooks.Count)
End Function''', name="Probe")
    assert app.run("Report") == "Book1|Book3|Book3|2"


def test_copy_into_reopened_workbook_and_save_independently(tmp_path: Path) -> None:
    app = ExcelApplication()
    source = app.add_workbook()
    source_sheet = app.sheet(1)
    source_sheet.set_value("A1", 7)
    source_sheet.set_value("B1", "=SUM(A1,1)")
    source_sheet.add_named_range("LocalAmount", "=$A$1")
    destination = app.add_workbook()
    target_path = tmp_path / "destination.xlsx"
    app.save(target_path, workbook=destination)
    destination.Close(SaveChanges=False)
    destination = app.open_workbook(target_path)
    assert app.open_workbook(target_path) is destination
    copied = source_sheet.copy(after=app.sheet(1, workbook=destination))
    assert copied.name == "Sheet1 (2)"
    assert copied.value("B1") == 8
    copied.set_value("A1", 19)
    assert source_sheet.value("A1") == 7
    assert copied.value("B1") == 20
    source_sheet.copy_range("A1:B1", app.sheet(1, workbook=destination), "C3")
    assert app.sheet(1, workbook=destination).value("D3") == 8
    app.save(target_path, workbook=destination)
    app.save(target_path, workbook=destination)
    source_path = tmp_path / "source.xlsx"
    app.save(source_path, workbook=source)
    opened = ExcelApplication.open(target_path, with_vba=False)
    assert [sheet.name for sheet in opened.sheets()] == ["Sheet1", "Sheet1 (2)"]
    assert opened.sheet(2).value("A1") == 19
    assert opened.sheet(2).value("LocalAmount") == 19
    assert ExcelApplication.open(source_path, with_vba=False).sheet(1).value("A1") == 7


def test_sheet_copy_preserves_blank_cells_and_merges() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Build()
Range("A1:B2").Merge
Range("A1").Value = 7
Range("D1").Font.Bold = True
End Sub''', name="Builder")
    app.run("Build")
    copied = app.sheet(1).copy()
    assert copied.value("D1") is None
    app.add_module('''Function Report() As String
Report = Range("B2").MergeArea.Address & "|" & CStr(Range("D1").Font.Bold)
End Function''', name="Probe")
    assert app.run("Report") == "$A$1:$B$2|True"
