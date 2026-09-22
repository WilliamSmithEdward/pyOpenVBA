"""VBA worksheet relocation, retained objects, and workbook state."""
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError, VBARuntimeError

RECORDS = json.loads((Path(__file__).parent / 'fixtures/sheet_move.json').read_text())


@pytest.mark.parametrize('record', RECORDS, ids=[r['name'] for r in RECORDS])
def test_native_move(record: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Function Report() As String\nDim src As Object, dst As Object, bk As Object, sh As Object, nm As Object\n' + record['body'] + 'End Function', name='Probe')
    if record['name'] == 'names':
        with pytest.raises(VBAUnsupportedError, match='external-link'):
            app.run('Report')
    else:
        assert app.run('Report') == record['reported']


def test_vba_move_invalidates_held_objects_and_transfers_local_names() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Report() As String
Dim src As Object, dst As Object, sh As Object, held As Object, nm As Object, names As Object
Set src = ThisWorkbook
Set sh = src.Worksheets(1)
Set held = sh.Range("A1")
held.Value = 7
Set nm = sh.Names.Add("LocalAmount", "=Sheet1!$A$1")
Set names = sh.Names
sh.Range("B1").Formula = "=LocalAmount+1"
src.Worksheets.Add After:=sh
Set dst = Workbooks.Add
sh.Move Before:=dst.Worksheets(1)
Report = CStr(sh Is dst.Worksheets(1))
On Error Resume Next
held.Value = 20
Report = Report & "|" & CStr(Err.Number)
Err.Clear
Report = Report & sh.Name
Report = Report & "|" & CStr(Err.Number)
Err.Clear
Report = Report & nm.Name
Report = Report & "|" & CStr(Err.Number)
Err.Clear
Report = Report & CStr(names.Count)
Report = Report & "|" & CStr(Err.Number)
On Error GoTo 0
Set sh = dst.Worksheets(1)
sh.Range("A1").Value = 20
Report = Report & "|" & CStr(sh.Range("B1").Value) & "|" & sh.Names("LocalAmount").Name
End Function''', name='Probe')
    assert app.run('Report') == "False|424|424|424|424|21|'Sheet1 (2)'!LocalAmount"


@pytest.mark.parametrize('failure', ['references', 'both', 'different_app'])
def test_move_failure_is_atomic(failure: str) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    source = app.sheet(1)
    book.add_sheet('Other')
    app.sheet(2).set_value('A1', '=Sheet1!A1')
    destination = app.add_workbook()
    target = app.sheet(1)
    before = list(book.sheets_)
    if failure == 'references':
        with pytest.raises(VBAUnsupportedError, match='external-link'):
            source.move(after=target)
    elif failure == 'both':
        with pytest.raises(VBARuntimeError):
            source.move(before=target, after=target)
    else:
        other_app = ExcelApplication()
        other_app.add_workbook()
        with pytest.raises(VBARuntimeError):
            source.move(after=other_app.sheet(1))
    assert book.sheets_ == before
    assert len(destination.sheets_) == 1


def test_move_reorders_then_transfers_existing_packages(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    book.add_sheet('Other')
    source = app.sheet(1)
    source.set_value('A1', 7)
    source.add_named_range('LocalAmount', '=$A$1')
    source_path = tmp_path / 'source.xlsx'
    app.save(source_path)
    source.move(after=app.sheet(2))
    app.save(source_path)
    reopened = ExcelApplication.open(source_path, with_vba=False)
    assert [s.name for s in reopened.sheets()] == ['Other', 'Sheet1']
    assert reopened.sheet(2).value('LocalAmount') == 7
    destination = app.add_workbook()
    destination_path = tmp_path / 'destination.xlsx'
    app.save(destination_path)
    moved = source.move(before=app.sheet(1))
    assert moved is not source
    app.save(source_path, workbook=book)
    app.save(destination_path, workbook=destination)
    assert len(ExcelApplication.open(source_path, with_vba=False).sheets()) == 1
    reopened = ExcelApplication.open(destination_path, with_vba=False)
    assert reopened.sheet(1).value('LocalAmount') == 7


def test_reserved_names_follow_reorder_and_block_unsupported_transfer(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    book.add_sheet('Other')
    source = app.sheet(1)
    path = tmp_path / 'reserved.xlsx'
    app.save(path)
    xml = book.package.read('xl/workbook.xml').decode()
    xml = xml.replace('</sheets>', '</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</definedName></definedNames>')
    book.package.write('xl/workbook.xml', xml.encode())
    source.move(after=app.sheet(2))
    # The in-memory order differs from the package until the next save.
    destination = app.add_workbook()
    with pytest.raises(VBAUnsupportedError, match='reserved'):
        source.move(after=app.sheet(1))
    assert len(destination.sheets_) == 1
    app.save(path, workbook=book)
    from xml.etree import ElementTree as ET
    root = ET.fromstring(book.package.read('xl/workbook.xml'))
    reserved = next(node for node in root.iter() if node.attrib.get('name') == '_xlnm.Print_Area')
    assert reserved.attrib['localSheetId'] == '1'
