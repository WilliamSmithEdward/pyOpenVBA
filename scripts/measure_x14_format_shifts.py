"""What inserts and deletes do to conditional formats Excel keeps in a worksheet's x14 extension alone: cell-value and formula
rules whose formulas read another sheet.

    python scripts/measure_x14_format_shifts.py build      # Excel makes tests/fixtures/x14_format_base.xlsx
    python scripts/measure_x14_format_shifts.py            # the first look: tests/fixtures/x14_format_shifts.json
    python scripts/measure_x14_format_shifts.py predict 2  # tests/fixtures/x14_format_shifts2_predictions.json, from the model
    python scripts/measure_x14_format_shifts.py measure 2  # asks Excel those and scores them

The base book has numbers in Sheet1!A1:H12 and in Sheet2!A1:A6. Sheet1 holds two rules that read Sheet2, which Excel keeps in the
x14 extension alone (B2:B6, cell value greater than Sheet2!$A$1; D2:E6, the formula =D2>Sheet2!$A$2), and one it keeps in the
plain part (F2:F8, cell value greater than 50). Each case opens a copy, makes one edit on either sheet, reads every rule of Sheet1
(its type, range, priority and Formula1), saves the book, and keeps Sheet1's conditionalFormatting elements and the extension's
x14:conditionalFormattings as the saved file spells them.

tests/fixtures/x14_format_shifts_bar.json is the first look from tests/fixtures/x14_format_base_bar.xlsx, the same book with a data
bar on G2:G8 too, built while this BUILD held that rule: it is why the model refuses to reshape a data bar, icon set or colour
scale beside the extension's rules. Each fixture names the base it was measured from.
"""
from __future__ import annotations

import json
import random
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'
BASE = FIXTURES / 'x14_format_base.xlsx'
BUILD = r'''Public Function Build(ByVal path As String) As String
    Dim wb As Object, ws As Object, other As Object, i As Long, j As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Sheet1"
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "Sheet2"
    For i = 1 To 6
        other.Cells(i, 1).Value = i * 20
    Next
    For i = 1 To 12
        For j = 1 To 8
            ws.Cells(i, j).Value = i * 10 + j
        Next
    Next
    ws.Range("B2:B6").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=Sheet2!$A$1"
    ws.Range("D2:E6").FormatConditions.Add Type:=xlExpression, Formula1:="=D2>Sheet2!$A$2"
    ws.Range("F2:F8").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=50"
    ws.Activate
    wb.SaveAs path, 51
    wb.Close False
    Build = "built"
End Function
'''
SOURCE = r'''Private Function Rules(ws As Object) As String
    Dim fc As Object, out As String, i As Long, formula As String
    For i = 1 To ws.Cells.FormatConditions.Count
        Set fc = ws.Cells.FormatConditions(i)
        formula = ""
        On Error Resume Next
        formula = fc.Formula1
        On Error GoTo 0
        out = out & fc.Type & ":" & fc.AppliesTo.Address(False, False) & ":" & fc.Priority & ":" & formula & ";"
    Next
    Rules = out
End Function

Public Function Go(ByVal path As String, ByVal saved As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, place As Object, out As String, parts() As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets("Sheet1")
    parts = Split(edit, " ")
    If parts(0) = "other" Then
        Set place = wb.Worksheets("Sheet2")
    Else
        Set place = ws
    End If
    Set wt = place
    On Error Resume Next
    Select Case parts(1) & " " & parts(3)
        Case "insert down": If typed Then wt.Range(parts(2)).Insert Shift:=xlShiftDown Else place.Range(parts(2)).Insert Shift:=xlShiftDown
        Case "insert right": If typed Then wt.Range(parts(2)).Insert Shift:=xlShiftToRight Else place.Range(parts(2)).Insert Shift:=xlShiftToRight
        Case "delete up": If typed Then wt.Range(parts(2)).Delete Shift:=xlShiftUp Else place.Range(parts(2)).Delete Shift:=xlShiftUp
        Case "delete left": If typed Then wt.Range(parts(2)).Delete Shift:=xlShiftToLeft Else place.Range(parts(2)).Delete Shift:=xlShiftToLeft
        Case "insert lines": If typed Then wt.Range(parts(2)).Insert Else place.Range(parts(2)).Insert
        Case "delete lines": If typed Then wt.Range(parts(2)).Delete Else place.Range(parts(2)).Delete
    End Select
    out = "err=" & Err.Number & " rules=" & Rules(ws)
    Err.Clear
    On Error GoTo 0
    wb.SaveAs saved, 51
    wb.Close False
    Go = out
End Function
'''
EDITS = ('sheet insert B3 down', 'sheet insert B1 down', 'sheet insert A1:H1 down', 'sheet insert A3:H3 down', 'sheet delete B3 up',
         'sheet delete A3:H3 up', 'sheet delete B2:B6 up', 'sheet insert A3:A4 right', 'sheet insert C2:C6 right', 'sheet delete A2:A6 left',
         'sheet delete C4 left', 'sheet insert E3 down', 'sheet delete D3 up', 'sheet insert D4:E4 down', 'sheet delete D2:E6 left',
         'sheet insert 3:3 lines', 'sheet delete 3:3 lines', 'sheet delete 2:6 lines', 'sheet insert C:C lines', 'sheet delete D:D lines',
         'sheet delete B:B lines', 'other insert A1 down', 'other delete A1 up', 'other insert 1:1 lines', 'other delete 2:2 lines',
         'other delete A1:A6 up', 'other insert A:A lines', 'sheet insert F3 down', 'sheet delete G4 up', 'sheet insert A5:H5 down')


def first_look() -> list[list[object]]:
    return [[edit, typed] for edit in EDITS for typed in (True, False)]


def held_out(round_: str) -> list[list[object]]:
    chooser = random.Random('x14-formats' + round_)
    columns = 'ABCDEFGH'
    cases: list[list[object]] = []
    while len(cases) < 100:
        where = chooser.choice(('sheet', 'sheet', 'sheet', 'other'))
        verb = chooser.choice(('insert', 'delete'))
        kind = chooser.choice(('cells', 'cells', 'lines'))
        if kind == 'lines':
            if chooser.random() < .5:
                top = chooser.randint(1, 10)
                target = f'{top}:{top + chooser.choice((0, 0, 1, 3))}'
            else:
                left = chooser.randrange(7)
                target = f'{columns[left]}:{columns[min(7, left + chooser.choice((0, 0, 1)))]}'
            edit = f'{where} {verb} {target} lines'
        else:
            left, top = chooser.randrange(7), chooser.randint(1, 10)
            target = f'{columns[left]}{top}:{columns[min(7, left + chooser.choice((0, 0, 1, 3)))]}{top + chooser.choice((0, 0, 1, 4))}'
            how = chooser.choice(('down', 'right')) if verb == 'insert' else chooser.choice(('up', 'left'))
            edit = f'{where} {verb} {target} {how}'
        cases.append([edit, chooser.random() < .5])
    return cases


def parts_of(path: Path) -> str:
    """Sheet1's conditional formats as the saved file spells them: the plain elements, then the extension's."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    plain = ''.join(re.findall(r'<conditionalFormatting\b.*?</conditionalFormatting>', xml, re.DOTALL))
    extended = re.search(r'<x14:conditionalFormattings\b.*?</x14:conditionalFormattings>', xml, re.DOTALL)
    return plain + ' | ' + (extended.group(0) if extended else '')


def build() -> None:
    sys.path.insert(0, str(ROOT / 'scripts'))
    from fixture_workbook import copy_saved
    from pyvbaharness import ExcelSession, HarnessConfig

    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        made = Path(temp) / 'base.xlsx'
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(BUILD, 'Build', (str(made),), timeout=60.0, line_numbers=False)
            assert result.ok, result
        copy_saved(made, BASE)
    print('built', BASE.name, BASE.stat().st_size)


def _ask(cases: list[list[object]], base: Path = BASE) -> list[str]:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: list[str] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            for number, case in enumerate(cases):
                opened, saved = Path(temp) / f'in{number}.xlsx', Path(temp) / f'out{number}.xlsx'
                shutil.copyfile(base, opened)
                result = excel.run_vba(SOURCE, 'Go', (str(opened), str(saved), *case), timeout=60.0, line_numbers=False)
                answers.append(f'{result.value} ~ {parts_of(saved)}' if result.ok else f'ERROR {str(result)[:200]}')
    return answers


def _write(path: Path, cases: list[list[object]], answers: list[str]) -> None:
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'base': BASE.name, 'source': SOURCE, 'cases': cases,
                                'answers': answers},
                               indent=0), encoding='utf-8', newline='\n')


def model_answer(case: list[object], temp: Path, base: Path = BASE) -> str:
    """What the model answers for one case, opening the base book from the folder given and saving beside it."""
    from pyopenvba.apps.excel import ExcelApplication

    opened, saved = temp / 'in.xlsx', temp / 'out.xlsx'
    shutil.copyfile(base, opened)
    saved.unlink(missing_ok=True)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='Probe')
    found = str(app.run('Probe.Go', str(opened), str(saved), *case))
    return f'{found} ~ {parts_of(saved)}'


def predict(round_: str) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    found = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as temp:
        for case in held_out(round_):
            try:
                model: str | None = model_answer(case, Path(temp))
            except VBAUnsupportedError:
                model = None
            found.append({'call': case, 'model': model})
    (FIXTURES / f'x14_format_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


def measure(round_: str) -> None:
    predicted = json.loads((FIXTURES / f'x14_format_shifts{round_}_predictions.json').read_text('utf-8'))['cases']
    cases = [one['call'] for one in predicted]
    answers = _ask(cases)
    _write(FIXTURES / f'x14_format_shifts{round_}.json', cases, answers)
    wrong = [(one['call'], one['model'], answer) for one, answer in zip(predicted, answers) if one['model'] is not None and one['model'] != answer]
    refused = sum(one['model'] is None for one in predicted)
    print(f'{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}')
    for one in wrong[:20]:
        print(str(one)[:700])


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'build':
        build()
    elif len(sys.argv) > 2 and sys.argv[1] in ('predict', 'measure'):
        {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
    else:
        found = first_look()
        _write(FIXTURES / 'x14_format_shifts.json', found, _ask(found))
        print(len(found), 'measured')
