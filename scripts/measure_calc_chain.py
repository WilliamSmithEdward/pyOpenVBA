"""When Excel refuses a file for its calculation chain, and whether a file without one calculates the same.

    python scripts/measure_calc_chain.py

Excel makes a workbook with formulas of every kind and saves it (tests/fixtures/calc_chain.xlsx). For each edit, Excel makes the edit to its own
file and saves, and the model makes it and saves. Excel reads every cell of both, then changes a value and reads them again, and tries the model's
save with the chain the file came with put back: what it does with that decides whether the chain holds. Excel also opens files with chains written
by hand, and the same book in each format that has a chain, once with alerts on.

Writes tests/fixtures/calc_chain.json. tests/test_excel_calc_chain.py replays it.
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

OUT = ROOT / 'tests/fixtures/calc_chain.json'
BOOK = ROOT / 'tests/fixtures/calc_chain.xlsx'
#: The model refuses to rename a sheet in a workbook with an array, a spill or a data table, so the renames start from a book without them.
PLAIN = ROOT / 'tests/fixtures/calc_chain_plain.xlsx'
CHAIN = 'xl/calcChain.xml'
TYPES = '[Content_Types].xml'
RELATIONSHIPS = 'xl/_rels/workbook.xml.rels'

MAKE = r'''Public Function Make(ByVal stem As String, ByVal kinds As Boolean) As String
    Dim wb As Workbook, s As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 5: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    wb.Worksheets(1).Name = "One": wb.Worksheets(2).Name = "Two": wb.Worksheets(3).Name = "Three": wb.Worksheets(4).Name = "Forms": wb.Worksheets(5).Name = "Blank"
    With wb.Worksheets(1)
        .Range("A1").Value = 1: .Range("A2").Value = 2: .Range("A3").Value = 3: .Range("A4").Value = 4
        .Range("B1").Formula = "=A1+1"
        .Range("B2").Formula = "=A2+B1"
        .Range("C3").Formula = "=SUM(A1:A4)"
        .Range("D2:D4").Formula = "=A2*2"
    End With
    wb.Worksheets(2).Range("A1").Formula = "=One!C3"
    wb.Worksheets(2).Range("C1").Formula = "=4+4"
    wb.Worksheets(3).Range("A1").Formula = "=1+1"
    Set s = wb.Worksheets(4)
    s.Range("A1").Value = 1: s.Range("A2").Value = 2: s.Range("A3").Value = 3
    If kinds Then
        s.Range("B1").FormulaArray = "=SUM(A1:A3*2)"
        s.Range("C1:C3").FormulaArray = "=A1:A3*2"
        s.Range("D1").Formula2 = "=SEQUENCE(3)"
    End If
    s.Range("E1").Formula = "=E2+1": s.Range("E2").Formula = "=E3+1": s.Range("E3").Value = 1
    s.Range("F1:F3").Formula = "=ROW()*2"
    s.Range("G1").Value = 4
    s.Range("I2").Formula = "=G1*2": s.Range("H3").Value = 1: s.Range("H4").Value = 2
    If kinds Then s.Range("H2:I4").Table , s.Range("G1")
    wb.SaveAs stem & ".xlsx", xlOpenXMLWorkbook
    wb.SaveAs stem & ".xlsm", xlOpenXMLWorkbookMacroEnabled
    wb.SaveAs stem & ".xlsb", xlExcel12
    wb.SaveAs stem & ".xltx", xlOpenXMLTemplate
    wb.SaveAs stem & ".xltm", xlOpenXMLTemplateMacroEnabled
    wb.SaveAs stem & ".xlam", xlOpenXMLAddIn
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
#: The same reader runs in Excel and in the model.
READER = r'''Public Function CellsOf(wb As Object) As String
    Dim s As Object, r As Long, c As Long, out As String, v As Variant
    For Each s In wb.Worksheets
        out = out & "#" & s.Name
        For r = 1 To 14
            For c = 1 To 12
                v = s.Cells(r, c).Value2
                If IsError(v) Then v = "error"
                If Not IsEmpty(v) Then out = out & "|" & r & "," & c & "=" & s.Cells(r, c).Formula & "=" & v
            Next
        Next
    Next
    CellsOf = out
End Function
'''
READS = READER + r'''Public Function Reads(ByVal path As String) As String
    Dim wb As Workbook, first As String
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    If wb Is Nothing Then
        Reads = "refused " & Err.Number
    Else
        On Error GoTo 0
        first = CellsOf(wb)
        wb.Worksheets(1).Range("A3").Value = 100
        Reads = first & "@@" & CellsOf(wb)
        wb.Close False
    End If
    Application.DisplayAlerts = True
End Function
Public Function Tryit(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    Tryit = CStr(Err.Number)
    If Not wb Is Nothing Then wb.Close False
    Application.DisplayAlerts = True
End Function
Public Function WithAlerts(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = True
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    WithAlerts = Err.Number & " " & TypeName(wb)
    If Not wb Is Nothing Then wb.Close False
End Function
'''
EDITOR = r'''Public Function Edited(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path): Set ws = wb.Worksheets(1)
    {edit}
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
'''
EDITS = {
    'unrelated value': 'ws.Range("Z100").Value = 1',
    'value over a formula': 'ws.Range("B2").Value = 5',
    'clear a formula': 'ws.Range("B2").ClearContents',
    'clear all of formula': 'ws.Range("C3").Clear',
    'new formula': 'ws.Range("F1").Formula = "=A1"',
    'formula over a formula': 'ws.Range("B2").Formula = "=A1"',
    'formula over a value': 'ws.Range("A4").Formula = "=A1"',
    'row insert above': 'ws.Rows(1).Insert',
    'row insert below': 'ws.Rows(50).Insert',
    'row delete of a formula': 'ws.Rows(1).Delete',
    'row delete above': 'ws.Rows(3).Delete',
    'column insert before': 'ws.Columns(1).Insert',
    'column delete': 'ws.Columns(3).Delete',
    'range insert shift': 'ws.Range("B1").Insert Shift:=xlDown',
    'range delete shift': 'ws.Range("B1").Delete Shift:=xlUp',
    'cut a formula': 'ws.Range("C3").Cut ws.Range("H8")',
    'copy over a formula': 'ws.Range("A1").Copy ws.Range("C3")',
    'sort formulas': 'ws.Range("A1:D4").Sort Key1:=ws.Range("A1"), Order1:=xlDescending',
    'delete a sheet': 'Application.DisplayAlerts = False: wb.Worksheets("Two").Delete',
    'delete last sheet': 'Application.DisplayAlerts = False: wb.Worksheets("Three").Delete',
    'add a sheet': 'wb.Worksheets.Add',
    'move a sheet': 'wb.Worksheets("Three").Move Before:=wb.Worksheets(1)',
    'rename a sheet': 'ws.Name = "Renamed"',
    'copy a sheet': 'ws.Copy After:=ws',
    'merge over a formula': 'ws.Range("B1:B2").Merge',
    'calculate': 'Application.Calculate',
    # Beyond the first 26: what the chain's sheets and flags turn on.
    'delete a sheet without formulas': 'Application.DisplayAlerts = False: wb.Worksheets("Blank").Delete',
    'move and rename a sheet with formulas': 'wb.Worksheets("Two").Name = "Second": wb.Worksheets("Second").Move Before:=wb.Worksheets(1)',
    'rename a sheet and clear a formula': 'ws.Name = "Renamed": ws.Range("B2").ClearContents',
    'move a sheet and clear a formula': 'wb.Worksheets("Two").Move Before:=wb.Worksheets(1): wb.Worksheets("One").Range("B2").ClearContents',
    'delete then add a sheet': 'Application.DisplayAlerts = False: wb.Worksheets("Three").Delete: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)',
    'plain over a one-cell array': 'wb.Worksheets("Forms").Range("B1").Formula = "=1+1"',
    'clear a one-cell array': 'wb.Worksheets("Forms").Range("B1").ClearContents',
    'array over a plain formula': 'wb.Worksheets("Forms").Range("F1").FormulaArray = "=1+1"',
    'clear a whole array': 'wb.Worksheets("Forms").Range("C1:C3").ClearContents',
    'clear a spill': 'wb.Worksheets("Forms").Range("D1").ClearContents',
    'formula over a chain child': 'wb.Worksheets("Forms").Range("E1").Formula = "=1+1"',
    'clear a chain child': 'wb.Worksheets("Forms").Range("E1").ClearContents',
    'clear a shared child': 'wb.Worksheets("Forms").Range("F2").ClearContents',
    'clear the data table corner': 'wb.Worksheets("Forms").Range("I2").ClearContents',
}
#: Sheet ids: One 1, Two 2, Three 3, Forms 4, Blank 5. Both One and Two have formulas where the entries below say.
HAND = {
    'every i': '<c r="C1" i="2"/><c r="A1" i="2"/><c r="B2" i="1"/><c r="B1" i="1"/>',
    'i omitted, same sheet': '<c r="C1" i="2"/><c r="A1"/><c r="B2" i="1"/><c r="B1"/>',
    'omitted, cell only on the sheet before the last': '<c r="B1" i="1"/><c r="C1"/>',
    'omitted, cell on the inherited sheet': '<c r="A1" i="2"/><c r="C1"/>',
    'first omitted, formula on the first sheet': '<c r="B1"/>',
    'first omitted, formula on the second only': '<c r="A1"/>',
    'i zero': '<c r="B1" i="0"/>',
    'i is no sheet': '<c r="B1" i="9"/>',
    'i is a sheet without the formula': '<c r="B1" i="5"/>',
    'cell with no formula': '<c r="Z9" i="1"/>',
    'cell with a value': '<c r="A1" i="1"/>',
    'twice the same entry': '<c r="B1" i="1"/><c r="B1" i="1"/>',
    'no entries': '',
    'lower case': '<c r="b1" i="1"/>',
    'dollars': '<c r="$B$1" i="1"/>',
    'one-cell array, no flag': '<c r="B1" i="4"/>',
    'one-cell array, flag': '<c r="B1" i="4" a="1"/>',
    'array master, no flag': '<c r="C1" i="4"/>',
    'array child, no flag': '<c r="C2" i="4"/>',
    'array child, flag': '<c r="C2" i="4" a="1"/>',
    'cell after an array, no flag': '<c r="C4" i="4"/>',
    'plain formula, flag': '<c r="E1" i="4" a="1"/>',
    'spill master': '<c r="D1" i="4"/>',
    'spill child, flag': '<c r="D2" i="4" a="1"/>',
    'data table first cell': '<c r="I3" i="4"/>',
    'data table child': '<c r="I4" i="4"/>',
    'shared child': '<c r="F2" i="4"/>',
    'unknown attribute': '<c r="E1" i="4" z="1"/>',
}
FORMATS = ('.xlsm', '.xltx', '.xltm', '.xlam')
RENAMES = ('rename a sheet', 'move and rename a sheet with formulas', 'rename a sheet and clear a formula')


def edited_in_model(source: Path, out: Path, edit: str) -> str:
    """The model's save of this edit, or what it refused."""
    app = ExcelApplication.open(source, with_vba=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1)\n    {edit}\nEnd Sub\n', name='E')
    try:
        app.run('E.Edit')
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return f'{type(refused).__name__} {str(refused)[:120]}'
    return ''


def with_chain(source: Path, out: Path, chain: bytes) -> None:
    """``source`` holding ``chain`` as its calculation chain, with the Override and Relationship one needs."""
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as new:
        had = CHAIN in old.namelist()
        for name in old.namelist():
            data = old.read(name)
            if name == CHAIN:
                data = chain
            elif name == TYPES and not had:
                data = data.replace(b'</Types>', b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>')
            elif name == RELATIONSHIPS and not had:
                data = data.replace(b'</Relationships>', b'<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>')
            new.writestr(name, data)
        if not had:
            new.writestr(CHAIN, chain)


def without_chain(source: Path, out: Path) -> None:
    """``source`` with no calculation chain, as the model drops one."""
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as new:
        for name in old.namelist():
            data = old.read(name)
            if name == CHAIN:
                continue
            if name == TYPES:
                data = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b'', data)
            if name == RELATIONSHIPS:
                data = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b'', data)
            new.writestr(name, data)


def chain_entries(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        if CHAIN not in package.namelist():
            return []
        return re.findall(r'<c [^>]*/>', package.read(CHAIN).decode('utf-8'))


def reads(excel: ExcelSession, path: Path) -> str:
    done = excel.run_vba(READS, 'Reads', (str(path),), timeout=120.0, line_numbers=False)
    assert done.ok, (path.name, done)
    return str(done.value)


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        stem = base / 'book'
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (str(stem), True), timeout=120.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(stem.with_suffix('.xlsx'), BOOK)
            plain = base / 'plain'
            made = excel.run_vba(MAKE, 'Make', (str(plain), False), timeout=120.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(plain.with_suffix('.xlsx'), PLAIN)
            with zipfile.ZipFile(BOOK) as package:
                chain_bytes = package.read(CHAIN)
            with zipfile.ZipFile(PLAIN) as package:
                plain_chain_bytes = package.read(CHAIN)
            chain_head = chain_bytes[:chain_bytes.index(b'<c ')]
            edits: dict[str, dict[str, object]] = {}
            for number, (name, edit) in enumerate(EDITS.items()):
                own = base / f'own{number}.xlsx'
                source = READS + EDITOR.format(edit=edit)
                start = PLAIN if name in RENAMES else BOOK
                done = excel.run_vba(source, 'Edited', (str(start), str(own)), timeout=60.0, line_numbers=False)
                row: dict[str, object] = {'edit': edit}
                if not (done.ok and done.value == 'ok'):
                    row['excel_refuses'] = str(done.outcome)
                else:
                    row['excel_chain'] = chain_entries(own)
                    row['excel_reads'] = reads(excel, own)
                    # Excel's own save with its chain taken out: whether a missing chain changes anything by itself.
                    bare = base / f'bare{number}.xlsx'
                    without_chain(own, bare)
                    row['excel_reads_without_chain'] = reads(excel, bare)
                saved = base / f'model{number}.xlsx'
                refused = edited_in_model(start, saved, edit)
                if refused:
                    row['model_refuses'] = refused
                    edits[name] = row
                    print(name.ljust(40), 'MODEL REFUSES', refused[:90])
                    continue
                with zipfile.ZipFile(saved) as package:
                    row['chain_kept'] = CHAIN in package.namelist()
                    row['full_calc_on_load'] = b'fullCalcOnLoad' in package.read('xl/workbook.xml')
                stale = base / f'stale{number}.xlsx'
                with_chain(saved, stale, plain_chain_bytes if name in RENAMES else chain_bytes)
                row['opens_with_old_chain'] = excel.run_vba(READS, 'Tryit', (str(stale),), timeout=60.0, line_numbers=False).value == '0'
                row['model_reads'] = reads(excel, saved)
                edits[name] = row
                print(name.ljust(40), 'kept' if row['chain_kept'] else 'dropped', 'old chain opens' if row['opens_with_old_chain'] else 'old chain refused',
                      'same' if row.get('excel_reads') == row['model_reads'] else 'READS DIFFER',
                      'bare same' if row.get('excel_reads') == row.get('excel_reads_without_chain') else 'BARE DIFFERS', 'FULLCALC' if row['full_calc_on_load'] else '')
            hand: dict[str, bool] = {}
            for number, (name, body) in enumerate(HAND.items()):
                out = base / f'hand{number}.xlsx'
                with_chain(BOOK, out, chain_head + body.encode() + b'</calcChain>')
                hand[name] = excel.run_vba(READS, 'Tryit', (str(out),), timeout=60.0, line_numbers=False).value == '0'
                print(name.ljust(50), hand[name])
            formats: dict[str, dict[str, object]] = {}
            for suffix in FORMATS:
                saved = base / f'cleared{suffix}'
                refused = edited_in_model(stem.with_suffix(suffix), saved, 'ws.Range("B2").ClearContents')
                with zipfile.ZipFile(stem.with_suffix(suffix)) as package:
                    had = CHAIN in package.namelist()
                opened = '' if refused else excel.run_vba(READS, 'Tryit', (str(saved),), timeout=60.0, line_numbers=False).value
                # The same file after an edit the chain has nothing to do with: what a failure owes to the chain.
                other = base / f'other{suffix}'
                refused_other = edited_in_model(stem.with_suffix(suffix), other, 'ws.Range("Z100").Value = 1')
                opened_other = '' if refused_other else excel.run_vba(READS, 'Tryit', (str(other),), timeout=60.0, line_numbers=False).value
                formats[suffix] = {'chain_in_excels_file': had, 'model_refuses': refused, 'opened': opened, 'opened_after_unrelated_edit': opened_other}
                print(suffix, formats[suffix])
            with zipfile.ZipFile(stem.with_suffix('.xlsb')) as package:
                formats['.xlsb'] = {'chain_parts': [name for name in package.namelist() if 'calcchain' in name.lower()]}
            print('.xlsb', formats['.xlsb'])
        # Last and in a session of its own, as nothing answers the question Excel asks and the session goes with it.
        stale = base / 'alerts.xlsx'
        assert not edited_in_model(BOOK, base / 'cleared.xlsx', 'ws.Range("B2").ClearContents')
        with_chain(base / 'cleared.xlsx', stale, chain_bytes)
        alerts: dict[str, object] = {}
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            asked = excel.run_vba(READS, 'WithAlerts', (str(stale),), timeout=30.0, line_numbers=False)
            alerts = {'outcome': str(asked.outcome), 'value': str(asked.value), 'dialogs': [one.message.replace(str(stale), '<file>').replace(str(base), '<folder>') for one in asked.dialogs]}
            print('alerts', alerts)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'reader': READER, 'edits': edits, 'hand': HAND,
                               'hand_opens': hand, 'formats': formats, 'alerts': alerts}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
