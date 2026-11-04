"""What becomes of a name local to a sheet when the sheet is deleted, which depends on what else reads a name of that spelling.

    python scripts/measure_deleted_local_names.py

Excel makes each book, saves it, opens it and deletes the sheet Gone; the model does the same to its own save of the book. The record has what each reads in the
session and on opening its save, and what the save holds.

Writes tests/fixtures/deleted_local_names.json. tests/test_excel_deleted_local_names.py replays it.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import copy_saved  # noqa: E402
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/deleted_local_names.json'
BOOK = ROOT / 'tests/fixtures/deleted_active_tab.xlsx'
#: What stands beside =Gone!Mine, which each book has on Reader!A1 unless it clears it.
BOOKS = {
    'qualified only': '',
    'bare on another sheet': 'r.Range("B1").Formula = "=Mine"',
    'bare on the deleted sheet': 'g.Range("D1").Formula = "=Mine"',
    'qualified twice': 'r.Range("B1").Formula = "=Gone!Mine+1"',
    'qualified on the deleted sheet too': 'g.Range("D1").Formula = "=Gone!Mine"',
    'qualified on the deleted sheet alone': 'r.Range("A1").ClearContents: g.Range("D1").Formula = "=Gone!Mine"',
    'bare, in arithmetic': 'r.Range("B1").Formula = "=Mine+1"',
    'bare, inside SUM': 'r.Range("B1").Formula = "=SUM(Mine)"',
    'bare with no qualified': 'r.Range("A1").ClearContents: r.Range("B1").Formula = "=Mine"',
    'read by a workbook name': 'wb.Names.Add "Other", "=Gone!Mine"',
    'read by a workbook name alone': 'r.Range("A1").ClearContents: wb.Names.Add "Other", "=Gone!Mine"',
    'read by a local name on the sheet': 'r.Range("A1").ClearContents: g.Names.Add "Other", "=Gone!Mine"',
    'an area': 'g.Names("Mine").RefersTo = "=Gone!$A$1:$A$3"',
    'a constant': 'g.Names("Mine").RefersTo = "=7"',
    'a name of another sheet too': 'wb.Worksheets("Keep").Names.Add "Mine", "=Keep!$A$1"',
    'a workbook name of the same spelling': 'wb.Names.Add "Mine", "=Keep!$A$1"',
    'nothing reads it': 'r.Range("A1").ClearContents',
    'hidden': 'g.Names("Mine").Visible = False',
    'with a comment': 'g.Names("Mine").Comment = "says"',
}
MAKE = r'''Public Function Make(ByVal path As String) As String
    Dim wb As Object, r As Object, g As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 3: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    Set r = wb.Worksheets(1): r.Name = "Reader"
    Set g = wb.Worksheets(2): g.Name = "Gone"
    wb.Worksheets(3).Name = "Keep": wb.Worksheets(3).Range("A1").Value = 5
    g.Range("A1").Value = 1: g.Range("A2").Value = 2: g.Range("A3").Value = 3
    g.Names.Add "Mine", "=Gone!$A$1"
    r.Range("A1").Formula = "=Gone!Mine"
    {extra}
    wb.SaveAs path, 51
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
READER = r'''Public Function Reads(wb As Object) As String
    Dim one As Object, c As String, n As Long
    On Error Resume Next
    For n = 1 To 2
        c = c & "@@" & wb.Worksheets("Reader").Cells(1, n).Formula & "=" & wb.Worksheets("Reader").Cells(1, n).Text
    Next
    For Each one In wb.Names
        If InStr(one.Name, "_xlfn") = 0 Then c = c & "@@" & one.Name & "=" & one.RefersTo & "=" & one.Visible & "=" & one.Comment
    Next
    Reads = c
End Function
'''
EXCEL = READER + r'''Public Function Run(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    wb.Worksheets("Gone").Delete
    s = Reads(wb)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(out)
    s = s & "##" & Reads(wb)
    wb.Close False
    Application.DisplayAlerts = True
    Run = s
End Function
Public Function Opened(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    If wb Is Nothing Then
        Opened = "refused " & Err.Number
    Else
        Opened = Reads(wb)
        wb.Close False
    End If
    Application.DisplayAlerts = True
End Function
'''


VIEWS = r'''Public Function MakeView(ByVal path As String, ByVal active As Long) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 4: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    wb.Worksheets(active).Activate
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    MakeView = "ok"
End Function
Public Function DeleteOne(ByVal path As String, ByVal out As String, ByVal gone As Long) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    wb.Worksheets(gone).Delete
    DeleteOne = wb.ActiveSheet.Name
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
End Function
Public Function ActiveOf(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    ActiveOf = Err.Number & " " & wb.ActiveSheet.Name
    If Not wb Is Nothing Then wb.Close False
    Application.DisplayAlerts = True
End Function
'''


def tabs(path: Path) -> dict[str, int]:
    """The first and active tabs a file's workbookView holds, 0 if it says nothing."""
    with zipfile.ZipFile(path) as package:
        tag = re.search(r'<workbookView\b[^>]*/>', package.read('xl/workbook.xml').decode('utf-8'))
    assert tag is not None
    found = dict(re.findall(r'\b(firstSheet|activeTab)="(\d+)"', tag.group(0)))
    return {'first': int(found.get('firstSheet', 0)), 'active': int(found.get('activeTab', 0))}


def measure_views(excel: ExcelSession, base: Path) -> list[dict[str, object]]:
    """Each sheet active and each deleted: the tabs in Excel's file, and in the model's, and what Excel then has active."""
    rows: list[dict[str, object]] = []
    for active in (1, 2, 3, 4):
        book = base / f'view{active}.xlsx'
        made = excel.run_vba(VIEWS, 'MakeView', (str(book), active), timeout=60.0, line_numbers=False)
        assert made.ok and made.value == 'ok', made
        if active == 1:
            copy_saved(book, BOOK)
        for gone in (1, 2, 4):
            own = base / f'viewown{active}{gone}.xlsx'
            done = excel.run_vba(VIEWS, 'DeleteOne', (str(book), str(own), gone), timeout=60.0, line_numbers=False)
            assert done.ok, done
            app = ExcelApplication.open(book, with_vba=False, sandbox=False)
            app.add_module(f'Public Function E() As String\n    Application.DisplayAlerts = False\n    ActiveWorkbook.Worksheets({gone}).Delete\n    E = ActiveWorkbook.ActiveSheet.Name\nEnd Function\n', name='M')
            session = str(app.run('M.E'))
            saved = base / f'viewmodel{active}{gone}.xlsx'
            app.save(saved)
            opened = str(excel.run_vba(VIEWS, 'ActiveOf', (str(saved),), timeout=60.0, line_numbers=False).value)
            own_opened = str(excel.run_vba(VIEWS, 'ActiveOf', (str(own),), timeout=60.0, line_numbers=False).value)
            row = {'active': active, 'gone': gone, 'start': tabs(book), 'excel': tabs(own), 'excel_session': str(done.value), 'excel_opens': own_opened, 'model': tabs(saved), 'model_session': session, 'model_opens': opened}
            rows.append(row)
            print(active, gone, 'same' if row['excel'] == row['model'] and own_opened == opened and session == done.value else f'DIFFERS {row}')
    return rows


def file_says(path: Path) -> dict[str, list[str]]:
    with zipfile.ZipFile(path) as package:
        return {'names': re.findall(r'<definedName\b[^>]*>[^<]*</definedName>', package.read('xl/workbook.xml').decode('utf-8')),
                'formulas': re.findall(r'<f[^>]*>([^<]*)</f>', package.read('xl/worksheets/sheet1.xml').decode('utf-8'))}


def plain(text: str, *books: str) -> str:
    """Reads with the workbook's own name, which differs by file, as one word."""
    for book in books:
        text = text.replace(book, 'BOOK')
    return text


def in_model(start: Path, out: Path) -> tuple[str, str]:
    app = ExcelApplication.open(start, with_vba=False, sandbox=False)
    app.add_module(READER + 'Public Function Go() As String\n    Application.DisplayAlerts = False\n    ActiveWorkbook.Worksheets("Gone").Delete\n    Go = Reads(ActiveWorkbook)\nEnd Function\n', name='E')
    try:
        session = str(app.run('E.Go'))
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return '', f'{type(refused).__name__} {str(refused)[:140]}'
    return session, ''


def main() -> None:
    rows: dict[str, dict[str, object]] = {}
    views: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            views = measure_views(excel, base)
            for number, (name, extra) in enumerate(BOOKS.items()):
                book = base / f'b{number}.xlsx'
                source = MAKE.replace('{extra}', extra)
                made = excel.run_vba(source, 'Make', (str(book),), timeout=60.0, line_numbers=False)
                assert made.ok and made.value == 'ok', (name, made)
                own = base / f'o{number}.xlsx'
                done = excel.run_vba(EXCEL, 'Run', (str(book), str(own)), timeout=60.0, line_numbers=False)
                assert done.ok, (name, done)
                session, reopened = plain(str(done.value), book.name, own.name).split('##')
                row: dict[str, object] = {'extra': extra, 'excel_session': session, 'excel_reopened': reopened, 'excel_file': file_says(own)}
                saved = base / f'm{number}.xlsx'
                answer, refused = in_model(book, saved)
                if refused:
                    row['model_refuses'] = refused
                else:
                    row['model_session'] = plain(answer, book.name)
                    row['model_reopened'] = plain(str(excel.run_vba(EXCEL, 'Opened', (str(saved),), timeout=60.0, line_numbers=False).value), saved.name)
                    row['model_file'] = file_says(saved)
                rows[name] = row
                same = (row.get('model_session') == session, row.get('model_reopened') == reopened, row.get('model_file') == row['excel_file'])
                print(name.ljust(38), same, refused[:80])
                if not all(same) and not refused:
                    print('   excel ', session[:150], '##', reopened[:90], row['excel_file'])
                    print('   model ', str(row['model_session'])[:150], '##', str(row['model_reopened'])[:90], row['model_file'])
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reader': READER, 'books': rows, 'views': views}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
