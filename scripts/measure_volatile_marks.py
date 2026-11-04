"""Which formulas a file marks ca="1", and whether a formula without the mark falls behind the volatile one it reads.

    python scripts/measure_volatile_marks.py

Excel makes a workbook whose formulas read volatile ones every way they can (tests/fixtures/volatile_marks.xlsx). The record has the f attributes Excel
writes; what Excel works out on opening its file with the marks taken off and a wrong value left in each; and, for each edit, the f attributes of Excel's
save and the model's, and on opening each whether every formula still agrees with what it reads.

Writes tests/fixtures/volatile_marks.json. tests/test_excel_volatile_marks.py replays it.
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

OUT = ROOT / 'tests/fixtures/volatile_marks.json'
BOOK = ROOT / 'tests/fixtures/volatile_marks.xlsx'

MAKE = r'''Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, s As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    wb.Worksheets.Add After:=wb.Worksheets(1)
    wb.Worksheets(1).Name = "V": wb.Worksheets(2).Name = "O"
    wb.Names.Add "VolName", "=RAND()"
    wb.Names.Add "PlainName", "=V!$A$1"
    Set s = wb.Worksheets("V")
    s.Range("A1").Formula = "=RAND()"
    s.Range("B1").Formula = "=A1+1"
    s.Range("C1").Formula = "=B1+1"
    s.Range("D1").Formula = "=SUM(A1:A2)"
    s.Range("E1").Formula = "=IF(FALSE,A1,5)"
    s.Range("F1").Formula = "=VolName"
    s.Range("G1").Formula = "=PlainName+1"
    s.Range("H1").Formula = "=TODAY()"
    s.Range("I1").Formula = "=H1+1"
    s.Range("J1").Formula = "=1+1"
    s.Range("K1").Formula = "=J1+1"
    s.Range("L1").Formula = "=RANDBETWEEN(1,1)"
    s.Range("M1").Formula = "=L1+1"
    s.Range("N1").Formula = "=OFFSET(J1,0,0)"
    s.Range("P1").Formula = "=N1+1"
    s.Range("Q1").Formula = "=INDIRECT(""J1"")"
    s.Range("R1:R3").Formula = "=$A$1+ROW()"
    s.Range("S1:S2").FormulaArray = "=A1+1"
    s.Range("T1:T2").Formula = "=A1+1"
    s.Range("U1").Formula = "=IF(J1>0,A1,5)"
    s.Range("U2").Formula = "=IF(J1<0,A1,5)"
    s.Range("U3").Formula = "=CHOOSE(1,5,A1)"
    s.Range("W1").Formula = "=ISNUMBER(A1)"
    s.Range("X1").Formula = "=IFERROR(A1,5)"
    s.Range("Y1").Formula = "=ROW(A1)"
    Set s = wb.Worksheets("O")
    s.Range("A1").Formula = "=V!B1"
    s.Range("B1").Formula = "=A1+1"
    s.Range("C1").Formula = "=V!J1"
    s.Range("D1").Formula = "=C1+1"
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
#: Each is 0 while a formula agrees with what it reads.
AGREES = {'V!B1': 'v.Range("B1").Value2 - v.Range("A1").Value2 - 1', 'V!C1': 'v.Range("C1").Value2 - v.Range("A1").Value2 - 2',
          'V!D1': 'v.Range("D1").Value2 - v.Range("A1").Value2', 'V!G1': 'v.Range("G1").Value2 - v.Range("A1").Value2 - 1',
          'V!I1': 'v.Range("I1").Value2 - v.Range("H1").Value2 - 1', 'V!K1': 'v.Range("K1").Value2 - v.Range("J1").Value2 - 1',
          'V!R1': 'v.Range("R1").Value2 - v.Range("A1").Value2 - 1', 'V!R3': 'v.Range("R3").Value2 - v.Range("A1").Value2 - 3',
          'V!S1': 'v.Range("S1").Value2 - v.Range("A1").Value2 - 1', 'V!S2': 'v.Range("S2").Value2 - v.Range("A1").Value2 - 1',
          'V!T1': 'v.Range("T1").Value2 - v.Range("A1").Value2 - 1', 'V!U1': 'v.Range("U1").Value2 - v.Range("A1").Value2',
          'O!A1': 'o.Range("A1").Value2 - v.Range("A1").Value2 - 1', 'O!B1': 'o.Range("B1").Value2 - v.Range("A1").Value2 - 2',
          'O!C1': 'o.Range("C1").Value2 - v.Range("J1").Value2', 'O!D1': 'o.Range("D1").Value2 - v.Range("J1").Value2 - 1'}
READS = 'Public Function Agrees(ByVal path As String) As String\n    Dim wb As Workbook, v As Worksheet, o As Worksheet, out As String\n    Application.DisplayAlerts = False\n' \
    '    Set wb = Workbooks.Open(path): Set v = wb.Worksheets("V"): Set o = wb.Worksheets("O")\n    On Error Resume Next\n' + ''.join(
        f'    out = out & "|{name}=" & Round({expression}, 8)\n    If Err.Number Then out = out & "E" & Err.Number: Err.Clear\n' for name, expression in AGREES.items()) + \
    '    On Error GoTo 0\n    wb.Close False\n    Application.DisplayAlerts = True\n    Agrees = out\nEnd Function\n'
EDITOR = r'''Public Function Edited(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path): Set ws = wb.Worksheets("V")
    {edit}
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
'''
EDITS = {
    'write on the sheet, so it is rewritten': 'ws.Range("AA1").Value = 1',
    'write on the other sheet': 'wb.Worksheets("O").Range("AA1").Value = 1',
    'a plain formula made volatile': 'ws.Range("J1").Formula = "=RAND()"',
    'a volatile formula made plain': 'ws.Range("A1").Formula = "=0.5"',
    'a value in place of a volatile formula': 'ws.Range("A1").Value = 0.5',
    'a new formula reading a volatile one': 'wb.Worksheets("O").Range("E1").Formula = "=V!A1+1"',
    'a name made volatile': 'wb.Names("PlainName").RefersTo = "=RAND()"',
}
_FORMULA = re.compile(r'<c r="([A-Z]+\d+)"[^>]*><f([^>]*?)(?:/>|>)')


def marks(path: Path) -> dict[str, dict[str, str]]:
    """Each formula cell's f attributes, by sheet."""
    with zipfile.ZipFile(path) as package:
        return {name: {reference: head.strip() for reference, head in _FORMULA.findall(package.read(part).decode('utf-8'))}
                for name, part in (('V', 'xl/worksheets/sheet1.xml'), ('O', 'xl/worksheets/sheet2.xml'))}


def without_marks(source: Path, out: Path, *, wrong: bool) -> None:
    """``source`` with no formula that reads a volatile one marked, and with ``wrong`` a value in each that is not its own."""
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as new:
        for name in old.namelist():
            data = old.read(name)
            if name.startswith('xl/worksheets/sheet'):
                def edit(found: re.Match[str]) -> str:
                    cell = found.group(0)
                    if '<f' not in cell or re.search(r'<f[^>]*>[^<]*(RAND|TODAY|OFFSET|INDIRECT|VolName)', cell):
                        return cell
                    cell = cell.replace(' ca="1"', '')
                    return re.sub(r'<v>[^<]*</v>', '<v>999</v>', cell) if wrong and not re.match(r'<c\b[^>]*\st=', cell) else cell
                data = re.sub(r'<c\b[^>]*?(?:/>|>.*?</c>)', edit, data.decode('utf-8'), flags=re.S).encode('utf-8')
            new.writestr(name, data)


def in_model(out: Path, edit: str) -> str:
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("V")\n    {edit}\nEnd Sub\n', name='E')
    try:
        app.run('E.Edit')
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return f'{type(refused).__name__} {str(refused)[:140]}'
    return ''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (str(base / 'book.xlsx'),), timeout=60.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(base / 'book.xlsx', BOOK)

            def agrees(path: Path) -> str:
                done = excel.run_vba(READS, 'Agrees', (str(path),), timeout=60.0, line_numbers=False)
                assert done.ok, (path.name, done)
                return str(done.value)

            bare = base / 'bare.xlsx'
            without_marks(BOOK, bare, wrong=False)
            wrong = base / 'wrong.xlsx'
            without_marks(BOOK, wrong, wrong=True)
            effect = {'marked': agrees(BOOK), 'marks_gone': agrees(bare), 'marks_gone_wrong_values': agrees(wrong)}
            print(effect)
            rows: dict[str, dict[str, object]] = {}
            for number, (name, edit) in enumerate(EDITS.items()):
                own = base / f'own{number}.xlsx'
                done = excel.run_vba(READS + EDITOR.format(edit=edit), 'Edited', (str(BOOK), str(own)), timeout=60.0, line_numbers=False)
                assert done.ok and done.value == 'ok', (name, done)
                row: dict[str, object] = {'edit': edit, 'excel_marks': marks(own), 'excel_agrees': agrees(own)}
                saved = base / f'model{number}.xlsx'
                refused = in_model(saved, edit)
                if refused:
                    row['model_refuses'] = refused
                else:
                    row['model_marks'] = marks(saved)
                    row['model_agrees'] = agrees(saved)
                rows[name] = row
                print(name, refused or ('same' if row.get('model_agrees') == row['excel_agrees'] else 'AGREES DIFFERS'), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'agrees': AGREES, 'start': marks(BOOK), 'effect': effect, 'edits': rows}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
