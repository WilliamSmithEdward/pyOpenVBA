"""What a formula that read a sheet becomes when the sheet is deleted, and what it then comes to.

    python scripts/measure_sheet_delete_refs.py

Excel makes a workbook whose formulas, names, validation and conditional formats read the sheets Gone and "My Gone" every way they can, deletes a sheet,
and reads each back, in the session and from its save. The model does the same to the same file (tests/fixtures/sheet_delete_refs.xlsx).

Writes tests/fixtures/sheet_delete_refs.json. tests/test_excel_sheet_delete_refs.py replays it.
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

OUT = ROOT / 'tests/fixtures/sheet_delete_refs.json'
BOOK = ROOT / 'tests/fixtures/sheet_delete_refs.xlsx'
#: The same with validation, a conditional format and a hyperlink reading the sheet too.
OBJECTS = ROOT / 'tests/fixtures/sheet_delete_refs_objects.xlsx'

#: The formulas of the sheet Reader, one to a row of column A.
FORMULAS = [
    '=Gone!A1',
    '=Gone!A1*2',
    '=Gone!$A$1',
    '=Gone!$A1+Gone!B$2',
    '=SUM(Gone!A1:B2)',
    '=SUM(Gone!$A$1:$B$2)',
    '=Gone!A:A',
    '=SUM(Gone!1:2)',
    "='My Gone'!A1",
    "=SUM('My Gone'!A1:B2)",
    '=Gone!A1+\'My Gone\'!A1',
    '=SUM(Gone:Middle!A1)',
    '=SUM(First:Gone!A1)',
    '=SUM(First:Last!A1)',
    '=SUM(First:Gone!A1:B2)',
    '=IFERROR(Gone!A1,"caught")',
    '=ISREF(Gone!A1)',
    '=ISERROR(Gone!A1)',
    '=IF(1,2,Gone!A1)',
    '=COUNT(Gone!A1:A3,1)',
    '=ROW(Gone!A5)',
    '=COLUMNS(Gone!A1:C1)',
    '=INDIRECT("Gone!A1")',
    '=OFFSET(Gone!A1,1,0)',
    '=INDEX(Gone!A1:A3,2)',
    '=VLOOKUP(1,Gone!A1:B3,2,FALSE)',
    '=GoneName',
    '=GoneName*2',
    '=GoneArea',
    '=SUM(GoneArea)',
    '=GoneConstant',
    '=LocalName',
    '=Gone!LocalName',
    '=Gone!A1&"x"',
    '="Gone!A1"',
    '=SUM(T[x])',
    '=#REF!',
    '=#REF!A1',
    '=SUM(#REF!A1:B2)',
    '=NA()+Gone!A1',
]
MAKE = r'''Public Function Make(ByVal path As String, ByVal objects As Boolean) As String
    Dim wb As Workbook, s As Worksheet, g As Worksheet, n As Long, f As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 6: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    wb.Worksheets(1).Name = "Reader": wb.Worksheets(2).Name = "First": wb.Worksheets(3).Name = "Gone"
    wb.Worksheets(4).Name = "My Gone": wb.Worksheets(5).Name = "Middle": wb.Worksheets(6).Name = "Last"
    For Each g In wb.Worksheets
        If g.Name <> "Reader" Then
            g.Range("A1").Value = 1: g.Range("A2").Value = 2: g.Range("A3").Value = 3
            g.Range("B1").Value = 4: g.Range("B2").Value = 5: g.Range("B3").Value = 6
        End If
    Next
    wb.Worksheets("Gone").Range("K1").Value = "x": wb.Worksheets("Gone").Range("K2").Value = 1
    wb.Worksheets("Gone").ListObjects.Add(xlSrcRange, wb.Worksheets("Gone").Range("K1:K2"), , xlYes).Name = "T"
    wb.Names.Add "GoneName", "=Gone!$A$1"
    wb.Names.Add "GoneArea", "=Gone!$A$1:$B$2"
    wb.Names.Add "GoneConstant", "=7"
    wb.Names.Add "GoneMixed", "=Gone!$A$1+First!$A$1"
    wb.Worksheets("Gone").Names.Add "LocalName", "=Gone!$B$1"
    wb.Worksheets("First").Names.Add "OtherLocal", "=Gone!$B$2"
    Set s = wb.Worksheets("Reader")
    n = 1
    Dim problems As String
    On Error Resume Next
    For Each f In Formulas()
        s.Cells(n, 1).Formula = f
        If Err.Number Then problems = problems & "|" & f: Err.Clear
        n = n + 1
    Next
    On Error GoTo 0
    If objects Then
        s.Range("C1").Validation.Add xlValidateList, , , "=Gone!$A$1:$A$3"
        s.Range("C2").Validation.Add xlValidateWholeNumber, , xlGreater, "=Gone!$A$1"
        s.Range("C3").FormatConditions.Add(xlExpression, , "=Gone!$A$1>0").Interior.Color = 255
        s.Hyperlinks.Add s.Range("C4"), "", "Gone!A1", , "to gone"
    End If
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok" & problems
End Function
'''
READER = r'''Public Function ReadAll(wb As Object, ByVal count As Long) As String
    Dim s As Object, n As Long, out As String, v As Variant, one As Object
    Set s = wb.Worksheets("Reader")
    On Error Resume Next
    For n = 1 To count
        v = s.Cells(n, 1).Value2
        If IsError(v) Then v = "error " & CStr(v)
        If IsArray(v) Then v = "array"
        out = out & "@@" & s.Cells(n, 1).Formula & "=" & TypeName(v) & "=" & CStr(v) & "=" & s.Cells(n, 1).Text
        If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    Next
    For Each one In wb.Names
        out = out & "@@name " & one.Name & "=" & one.RefersTo & "=" & one.Visible
        If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    Next
    out = out & "@@sheets " & wb.Worksheets.Count
    out = out & "@@v1 " & s.Range("C1").Validation.Formula1
    If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    out = out & "@@v2 " & s.Range("C2").Validation.Formula1
    If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    out = out & "@@cf " & s.Range("C3").FormatConditions(1).Formula1
    If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    out = out & "@@link " & s.Hyperlinks(1).SubAddress
    If Err.Number Then out = out & "=E" & Err.Number: Err.Clear
    ReadAll = out
End Function
'''
EXCEL = READER + r'''Public Function Run(ByVal path As String, ByVal out As String, ByVal count As Long) As String
    Dim wb As Workbook, ws As Worksheet, session As String, again As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path): Set ws = wb.Worksheets("Reader")
    {edit}
    session = ReadAll(wb, count)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(out)
    again = ReadAll(wb, count)
    wb.Close False
    Application.DisplayAlerts = True
    Run = session & "##" & again
End Function
Public Function Opened(ByVal path As String, ByVal count As Long) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    If wb Is Nothing Then
        Opened = "refused " & Err.Number
    Else
        On Error GoTo 0
        Opened = ReadAll(wb, count)
        wb.Close False
    End If
    Application.DisplayAlerts = True
End Function
'''
EDITS = {
    'delete Gone': 'wb.Worksheets("Gone").Delete',
    'delete My Gone': 'wb.Worksheets("My Gone").Delete',
    'delete both': 'wb.Worksheets("Gone").Delete: wb.Worksheets("My Gone").Delete',
    'delete the end of a span': 'wb.Worksheets("Last").Delete',
    'delete the start of a span': 'wb.Worksheets("First").Delete',
    'delete inside a span': 'wb.Worksheets("Middle").Delete',
    'delete, then a same-named sheet': 'wb.Worksheets("Gone").Delete: wb.Worksheets.Add().Name = "Gone"',
    'delete, then write a precedent': 'wb.Worksheets("Gone").Delete: wb.Worksheets("First").Range("A1").Value = 50',
    'delete, then rewrite a formula': 'wb.Worksheets("Gone").Delete: On Error Resume Next: ws.Range("A2").Formula = ws.Range("A2").Formula: ws.Range("Z1").Value = Err.Number: Err.Clear: On Error GoTo 0',
    'delete, then insert a row on Reader': 'wb.Worksheets("Gone").Delete: ws.Rows(1).Insert',
    'delete, then rename a sheet': 'wb.Worksheets("Gone").Delete: wb.Worksheets("First").Name = "Primero"',
    'objects, delete Gone': 'wb.Worksheets("Gone").Delete',
}


def formulas_vba() -> str:
    """The list as a VBA function, in pieces, since a line holds only so much."""
    lines = ['Private Function Formulas() As Variant', f'    Dim a(0 To {len(FORMULAS) - 1}) As String']
    for number, formula in enumerate(FORMULAS):
        spelled = formula.replace('"', '""')
        lines.append(f'    a({number}) = "{spelled}"')
    return '\n'.join(lines + ['    Formulas = a', 'End Function', ''])


def part_text(path: Path) -> dict[str, list[str]]:
    """What the file says of the formulas: the Reader sheet's f and v, the defined names."""
    with zipfile.ZipFile(path) as package:
        workbook = package.read('xl/workbook.xml').decode('utf-8')
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    cells = [f'{ref}|{kind}|{formula}|{value}' for ref, kind, formula, value in re.findall(
        r'<c r="([A-Z]+\d+)"(?: s="\d+")?(?: t="(\w+)")?[^>]*><f[^>]*?>(.*?)</f><v>(.*?)</v></c>', sheet)]
    names = re.findall(r'<definedName\b[^>]*>.*?</definedName>', workbook)
    other = re.findall(r'<(?:formula1|formula)>.*?</(?:formula1|formula)>', sheet) + re.findall(r'<hyperlink\b[^>]*/>', sheet)
    return {'cells': cells, 'names': names, 'other': other}


def in_model(start: Path, out: Path, edit: str) -> dict[str, object]:
    app = ExcelApplication.open(start, with_vba=False)
    app.add_module(READER + f'Public Function Run(ByVal count As Long) As String\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("Reader")\n    Application.DisplayAlerts = False\n    {edit}\n    Run = ReadAll(wb, count)\nEnd Function\n', name='E')
    try:
        session = str(app.run('E.Run', len(FORMULAS)))
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return {'refuses': f'{type(refused).__name__} {str(refused)[:160]}'}
    return {'session': session.split('@@')}


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(formulas_vba() + MAKE, 'Make', (str(base / 'book.xlsx'), False), timeout=60.0, line_numbers=False)
            assert made.ok and str(made.value).startswith('ok'), made
            print('made', made.value)
            copy_saved(base / 'book.xlsx', BOOK)
            made = excel.run_vba(formulas_vba() + MAKE, 'Make', (str(base / 'objects.xlsx'), True), timeout=60.0, line_numbers=False)
            assert made.ok and str(made.value).startswith('ok'), made
            copy_saved(base / 'objects.xlsx', OBJECTS)
            rows: dict[str, dict[str, object]] = {}
            for number, (name, edit) in enumerate(EDITS.items()):
                own = base / f'own{number}.xlsx'
                start = OBJECTS if name.startswith('objects') else BOOK
                done = excel.run_vba(EXCEL.format(edit=edit), 'Run', (str(start), str(own), len(FORMULAS)), timeout=60.0, line_numbers=False)
                assert done.ok, (name, done)
                session, again = str(done.value).split('##')
                row: dict[str, object] = {'edit': edit, 'excel_session': session.split('@@'), 'excel_reopened': again.split('@@'), 'excel_file': part_text(own)}
                saved = base / f'model{number}.xlsx'
                model = in_model(start, saved, edit)
                row['model'] = model
                if 'session' in model:
                    opened = excel.run_vba(EXCEL.format(edit=edit), 'Opened', (str(saved), len(FORMULAS)), timeout=60.0, line_numbers=False)
                    row['model_reopened'] = str(opened.value).split('@@')
                    row['model_file'] = part_text(saved)
                rows[name] = row
                print(name, model.get('refuses', 'model ran'), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'formulas': FORMULAS, 'make': MAKE, 'reader': READER, 'edits': rows}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
