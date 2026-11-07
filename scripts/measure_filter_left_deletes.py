"""What a delete left over an AutoFilter's whole height does to it: its range, its Filters.Count, its criteria and its hidden name.

Each case builds a new workbook with a header and eleven rows of numbers in a block of columns, an AutoFilter over that block
(with no criteria, or with criteria on its first column that hide no row), and maybe a cell of text elsewhere, then deletes a
range with Shift:=xlToLeft. It reads the filter, saves the book as .xlsx, opens it again and reads it once more.

    python scripts/measure_filter_left_deletes.py           # the first look: tests/fixtures/filter_left_deletes.json
    python scripts/measure_filter_left_deletes.py predict 2 # tests/fixtures/filter_left_deletes2_predictions.json, from the model
    python scripts/measure_filter_left_deletes.py measure 2 # asks Excel those and scores them: tests/fixtures/filter_left_deletes2.json

The first look runs typed only; a held-out round runs each case typed (Dim ws As Worksheet) and late (As Object).
"""
from __future__ import annotations

import json
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures'
SOURCE = r'''Private Function Dump(ws As Object) As String
    Dim one As String, i As Long
    On Error Resume Next
    Dump = "mode=" & ws.AutoFilterMode
    If ws.AutoFilterMode Then
        Dump = Dump & " range=" & ws.AutoFilter.Range.Address(False, False) & " count=" & ws.AutoFilter.Filters.Count
        For i = 1 To ws.AutoFilter.Filters.Count
            If ws.AutoFilter.Filters(i).On Then Dump = Dump & " f" & i & "=" & ws.AutoFilter.Filters(i).Criteria1
        Next
    End If
    one = ws.Names("_FilterDatabase").RefersTo
    Dump = Dump & " db=" & one
End Function

Public Function Go(ByVal first As Long, ByVal width As Long, ByVal criteria As Boolean, ByVal other As String, ByVal target As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, i As Long, j As Long, out As String, path As String
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wt = ws
    For j = first To first + width - 1
        ws.Cells(1, j).Value = "h" & j
        For i = 2 To 12
            ws.Cells(i, j).Value = i * 10 + j
        Next
    Next
    If other <> "" Then ws.Range(other).Value = "r"
    If criteria Then
        ws.Range(ws.Cells(1, first), ws.Cells(12, first + width - 1)).AutoFilter Field:=1, Criteria1:=">0"
    Else
        ws.Range(ws.Cells(1, first), ws.Cells(12, first + width - 1)).AutoFilter
    End If
    On Error Resume Next
    If typed Then wt.Range(target).Delete Shift:=xlToLeft Else ws.Range(target).Delete Shift:=xlToLeft
    out = "err=" & Err.Number & " " & Dump(ws)
    Err.Clear
    path = Environ("TEMP") & "\pyovba_filter_left.xlsx"
    Kill path
    Err.Clear
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    Application.DisplayAlerts = True
    wb.Close False
    Set wb = Workbooks.Open(path)
    out = out & " ~ " & Dump(wb.Worksheets(1))
    wb.Close False
    Kill path
    Go = out
End Function
'''


def first_look() -> list[list[object]]:
    """The cases the rule was read from."""
    cases: list[list[object]] = []
    for first, width in ((1, 3), (1, 5), (3, 3), (2, 4)):
        for other in ('', 'G1:G12', 'F5'):
            for target in ('A1:A12', 'B1:B12', 'C1:C12', 'A1:B12', 'B1:C12', 'B1:B13', 'B1:B20', 'B1:D12', 'C1:D12', 'D1:D12', 'E1:E12',
                           'F1:F12', 'A1:E12', 'C1:E12', 'A1:F12', 'A1:C1', 'A1:A1', 'A1:E11', 'A2:E12', 'B1:B11', 'C2:C12'):
                cases.append([first, width, False, other, target, True])
    return cases


def held_out(round_: str) -> list[list[object]]:
    """Fresh filters and deletes, chosen at random for the round, criteria or none, typed and late."""
    chooser = random.Random('filter-left' + round_)
    cases: list[list[object]] = []
    while len(cases) < 160:
        first, width = chooser.randint(1, 6), chooser.randint(1, 6)
        left = chooser.randint(1, first + width + 1)
        right = chooser.randint(left, left + chooser.randint(0, 4))
        top = chooser.choice((1, 1, 1, 2))
        bottom = chooser.choice((12, 12, 12, 13, 30, 11))
        target = f'{_column(left)}{top}:{_column(right)}{bottom}'
        other = chooser.choice(('', '', f'{_column(first + width + 2)}3', 'A20'))
        criteria = chooser.random() < .35
        for typed in (True, False):
            cases.append([first, width, criteria, other, target, typed])
    return cases


def _column(number: int) -> str:
    text = ''
    while number:
        number, rest = divmod(number - 1, 26)
        text = chr(65 + rest) + text
    return text


def _ask(cases: list[list[object]]) -> list[str]:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for case in cases:
            result = excel.run_vba(SOURCE, 'Go', tuple(case), timeout=120.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs else f'ERROR {str(result)[:200]}')
    return answers


def _write(path: Path, cases: list[list[object]], answers: list[str]) -> None:
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases, 'answers': answers}, indent=0),
                    encoding='utf-8', newline='\n')


def predict(round_: str) -> None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    found = []
    for case in held_out(round_):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(SOURCE, name='Probe')
        try:
            model: str | None = str(app.run('Probe.Go', *case))
        except VBAUnsupportedError:
            model = None
        found.append({'call': case, 'model': model})
    (FIXTURES / f'filter_left_deletes{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


def measure(round_: str) -> None:
    predicted = json.loads((FIXTURES / f'filter_left_deletes{round_}_predictions.json').read_text('utf-8'))['cases']
    cases = [one['call'] for one in predicted]
    answers = _ask(cases)
    _write(FIXTURES / f'filter_left_deletes{round_}.json', cases, answers)
    wrong = [(one['call'], one['model'], answer) for one, answer in zip(predicted, answers) if one['model'] is not None and one['model'] != answer]
    refused = sum(one['model'] is None for one in predicted)
    print(f'{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}')
    for one in wrong[:20]:
        print(one)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        found = first_look()
        _write(FIXTURES / 'filter_left_deletes.json', found, _ask(found))
        print(len(found), 'measured')
    else:
        {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
