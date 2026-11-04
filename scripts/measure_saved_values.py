"""What a save writes for the value of a formula on a sheet nobody wrote to.

    python scripts/measure_saved_values.py

Excel makes a workbook whose formulas read other sheets every way a formula can, and saves it (tests/fixtures/saved_values.xlsx). For each edit,
Excel makes the edit to its own file and saves, and the model makes it and saves. The record has each sheet's cells as the file holds them, and what
Excel reads from each, on opening and after a value changes.

Writes tests/fixtures/saved_values.json. tests/test_excel_saved_values.py replays it.
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

OUT = ROOT / 'tests/fixtures/saved_values.json'
BOOK = ROOT / 'tests/fixtures/saved_values.xlsx'

MAKE = r'''Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, ext As Workbook, s As Worksheet, chart As ChartObject
    Application.DisplayAlerts = False
    Set ext = Workbooks.Add
    ext.Worksheets(1).Name = "Sheet1": ext.Worksheets(1).Range("A1").Value = 7
    ext.SaveAs folder & "\ext.xlsx", xlOpenXMLWorkbook
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 6: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    wb.Worksheets(1).Name = "One": wb.Worksheets(2).Name = "Two": wb.Worksheets(3).Name = "Three"
    wb.Worksheets(4).Name = "Vol": wb.Worksheets(5).Name = "Link": wb.Worksheets(6).Name = "Untouched"
    With wb.Worksheets("One")
        .Range("A1").Value = 1: .Range("A2").Value = 2: .Range("A3").Value = 3: .Range("A4").Value = 4
        .Range("B1").Formula = "=A1*2"
        .Range("C3").Formula = "=SUM(A1:A4)"
        .Range("K1").Value = "x": .Range("L1").Value = "y": .Range("K2").Value = 1: .Range("K3").Value = 2
        .Range("L2").Value = 1: .Range("L3").Value = 2
        .ListObjects.Add(xlSrcRange, .Range("K1:L3"), , xlYes).Name = "T"
    End With
    wb.Names.Add "MyName", "=One!$A$1"
    Set s = wb.Worksheets("Two")
    s.Range("A1").Formula = "=One!C3"
    s.Range("B1").Formula = "=One!A1&""x"""
    s.Range("B2").Formula = "=One!A1>2"
    s.Range("B3").Formula = "=1/(One!A1-1)"
    s.Range("B4").Formula = "=IF(One!A1>1,""big"",1)"
    s.Range("B5").Formula = "=IF(One!A1>1,"""",1)"
    s.Range("C1:C3").Formula = "=One!$A$1*ROW()"
    s.Range("D1").FormulaArray = "=SUM(One!A1:A4*2)"
    s.Range("E1:E3").FormulaArray = "=One!A1:A3*2"
    s.Range("F1").Formula = "=MyName"
    s.Range("G1").Formula = "=SUM(T[x])"
    s.Range("H1").Formula = "=SUM(One:One!A1:A4)"
    s.Range("I1").Formula = "=SUBTOTAL(109,One!A1:A4)"
    s.Range("J1").Formula = "=One!A1/3"
    s.Range("J2").Formula = "=One!A1+45000": s.Range("J2").NumberFormat = "yyyy-mm-dd"
    s.Range("Z1").Value = 1: s.Range("Z2").Value = 2: s.Range("Z3").Value = 3
    Set chart = s.ChartObjects.Add(300, 100, 200, 120)
    chart.Chart.ChartType = xlColumnClustered
    chart.Chart.SetSourceData wb.Worksheets("One").Range("A1:A4")
    With wb.Worksheets("Three")
        .Range("A1").Formula = "=Two!A1*2"
        .Range("B1").Formula = "=A1+1"
        .Range("C1").Formula = "=Two!F1"
    End With
    With wb.Worksheets("Vol")
        .Range("A1").Formula = "=RAND()"
        .Range("B1").Formula = "=NOW()"
        .Range("C1").Formula = "=OFFSET(One!A1,1,0)"
        .Range("D1").Formula = "=INDIRECT(""One!A4"")"
        .Range("E1").Formula = "=TODAY()"
        .Range("F1").Formula = "=A1*0+1"
    End With
    With wb.Worksheets("Link")
        .Range("A1").Formula = "=[ext.xlsx]Sheet1!$A$1"
        .Range("B1").Formula = "=A1+One!A1"
    End With
    With wb.Worksheets("Untouched")
        .Range("A1").Value = 5
        .Range("B1").Formula = "=A1*2"
    End With
    ext.Close False
    wb.SaveAs folder & "\book.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
READER = r'''Public Function CellsOf(wb As Object) As String
    Dim s As Object, r As Long, c As Long, out As String, v As Variant
    For Each s In wb.Worksheets
        If s.Name <> "Vol" Then
            out = out & "#" & s.Name
            For r = 1 To 6
                For c = 1 To 13
                    v = s.Cells(r, c).Value2
                    If IsError(v) Then v = "error"
                    If Not IsEmpty(v) Then out = out & "|" & r & "," & c & "=" & TypeName(v) & "=" & v
                Next
            Next
        End If
    Next
    CellsOf = out
End Function
'''
READS = READER + r'''Public Function Reads(ByVal path As String) As String
    Dim wb As Workbook, first As String
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    On Error Resume Next
    Set wb = Workbooks.Open(path, UpdateLinks:=0)
    If wb Is Nothing Then
        Reads = "refused " & Err.Number
    Else
        On Error GoTo 0
        first = CellsOf(wb)
        wb.Worksheets("One").Range("A3").Value = 100
        Reads = first & "@@" & CellsOf(wb)
        wb.Close False
    End If
    Application.AskToUpdateLinks = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Function
'''
EDITOR = r'''Public Function Edited(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0): Set ws = wb.Worksheets("One")
    {edit}
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.CalculateBeforeSave = True
    Application.Calculation = xlCalculationAutomatic
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
'''
MANUAL = 'Application.Calculation = xlCalculationManual: '
EDITS = {
    'nothing that reads it': 'wb.Worksheets("Untouched").Range("Z9").Value = 1',
    'a value read everywhere': 'ws.Range("A1").Value = 2',
    'a value read by a sum': 'ws.Range("A4").Value = 50',
    'a value giving an error': 'ws.Range("A1").Value = 1: ws.Range("A1").Value = 1',
    'clear a value': 'ws.Range("A1").ClearContents',
    'text in a value': 'ws.Range("A1").Value = "abc"',
    'a formula over a value': 'ws.Range("A1").Formula = "=2+2"',
    'a table value': 'ws.Range("K2").Value = 40',
    'hide a row': 'ws.Rows(2).Hidden = True',
    'repoint a name': 'wb.Names("MyName").RefersTo = "=One!$A$2"',
    'insert a row': 'ws.Rows(1).Insert',
    'delete a row': 'ws.Rows(1).Delete',
    'delete a sheet nothing reads': 'wb.Worksheets("Untouched").Delete',
    'delete the sheet of the sum': 'wb.Worksheets("Two").Delete',
    'calculate': 'Application.Calculate',
    'the middle of a chain of sheets': 'wb.Worksheets("Two").Range("A1").Formula = "=7"',
    'the sheet with the linked value': 'wb.Worksheets("Link").Range("Z9").Value = 1',
    'manual, saving calculates': MANUAL + 'Application.CalculateBeforeSave = True: ws.Range("A4").Value = 50',
    'manual, saving does not calculate': MANUAL + 'Application.CalculateBeforeSave = False: ws.Range("A4").Value = 50',
}
_CELL = re.compile(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', re.S)


def sheet_parts(package: zipfile.ZipFile) -> dict[str, str]:
    """Each sheet's name against its part."""
    relationships = package.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    targets = {}
    for tag in re.findall(r'<Relationship\b[^>]*/>', relationships):
        targets[re.search(r'\bId="([^"]*)"', tag).group(1)] = re.search(r'\bTarget="([^"]*)"', tag).group(1)
    out = {}
    for tag in re.findall(r'<sheet\b[^>]*/>', package.read('xl/workbook.xml').decode('utf-8')):
        out[re.search(r'\bname="([^"]*)"', tag).group(1)] = 'xl/' + targets[re.search(r'\br:id="([^"]*)"', tag).group(1)]
    return out


def formula_cells(path: Path) -> dict[str, dict[str, str]]:
    """Each sheet's formula cells as the file has them: the f tag's attributes, the t, and the v."""
    with zipfile.ZipFile(path) as package:
        out: dict[str, dict[str, str]] = {}
        for name, part in sheet_parts(package).items():
            cells: dict[str, str] = {}
            for attributes, inner in _CELL.findall(package.read(part).decode('utf-8')):
                if not inner or '<f' not in inner:
                    continue
                reference = re.search(r'\br="([^"]*)"', attributes).group(1)
                kind = re.search(r'\bt="([^"]*)"', attributes)
                head = re.search(r'<f\b([^>]*?)/?>', inner).group(1).strip()
                value = re.search(r'<v>(.*?)</v>', inner, re.S)
                cells[reference] = f"{kind.group(1) if kind else ''}|{value.group(1) if value else ''}|{head}"
            out[name] = cells
        charts = [re.findall(r'<c:v>(.*?)</c:v>', package.read(name).decode('utf-8')) for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)]
        out['#charts'] = {str(number): ' '.join(values) for number, values in enumerate(charts)}
        calculation = re.search(r'<calcPr\b[^>]*/>', package.read('xl/workbook.xml').decode('utf-8'))
        out['#calcPr'] = {'tag': calculation.group(0) if calculation else ''}
        return out


def edited_in_model(out: Path, edit: str) -> tuple[str, str]:
    """What the model refused, or that it saved; and each sheet's values as the session reads them."""
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(READER + f'Public Function Edit() As String\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("One")\n    {edit}\n    Edit = CellsOf(wb)\nEnd Function\n', name='E')
    try:
        session = str(app.run('E.Edit'))
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return f'{type(refused).__name__} {str(refused)[:140]}', ''
    return '', session


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (folder,), timeout=120.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(base / 'book.xlsx', BOOK)
            rows: dict[str, dict[str, object]] = {}
            for number, (name, edit) in enumerate(EDITS.items()):
                own = base / f'own{number}.xlsx'
                row: dict[str, object] = {'edit': edit}
                done = excel.run_vba(READS + EDITOR.format(edit=edit), 'Edited', (str(base / 'book.xlsx'), str(own)), timeout=60.0, line_numbers=False)
                if done.ok and done.value == 'ok':
                    row['excel_cells'] = formula_cells(own)
                    row['excel_reads'] = str(excel.run_vba(READS, 'Reads', (str(own),), timeout=60.0, line_numbers=False).value)
                else:
                    row['excel_refuses'] = f'{done.outcome} {str(done.message)[:120]}'
                saved = base / f'model{number}.xlsx'
                refused, session = edited_in_model(saved, edit)
                if refused:
                    row['model_refuses'] = refused
                else:
                    row['session'] = session
                    row['model_cells'] = formula_cells(saved)
                    row['model_reads'] = str(excel.run_vba(READS, 'Reads', (str(saved),), timeout=60.0, line_numbers=False).value)
                rows[name] = row
                print(name, 'excel' if 'excel_cells' in row else row.get('excel_refuses'), row.get('model_refuses', 'model'), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reader': READER, 'start': formula_cells(BOOK), 'edits': rows}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
