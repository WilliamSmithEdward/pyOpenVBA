"""How far AutoFilter.Range runs down: blank columns in the filter, and data written beside or below it.

Each case builds a new workbook with a header and six rows in A1:C7, filters A1:C7, clears some of it and writes "x" in
other cells, then reads AutoFilter.Range and A1's CurrentRegion.

    python scripts/measure_filter_range_growth.py           # the first look: tests/fixtures/filter_range_growth.json
    python scripts/measure_filter_range_growth.py predict 2 # tests/fixtures/filter_range_growth2_predictions.json, from the model
    python scripts/measure_filter_range_growth.py measure 2 # asks Excel those and scores them: tests/fixtures/filter_range_growth2.json
"""
from __future__ import annotations

import json
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

FIXTURES = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures'
SOURCE = r'''Public Function Go(ByVal blank As String, ByVal extra As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, i As Long, j As Long, one As Variant, out As String
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wt = ws
    For j = 1 To 3
        ws.Cells(1, j).Value = "h" & j
        For i = 2 To 7
            ws.Cells(i, j).Value = i * 10 + j
        Next
    Next
    ws.Range("A1:C7").AutoFilter
    If blank <> "" Then ws.Range(blank).ClearContents
    For Each one In Split(extra, ";")
        If one <> "" Then ws.Range(one).Value = "x"
    Next
    If typed Then
        out = wt.AutoFilter.Range.Address(False, False) & " cr=" & wt.Range("A1").CurrentRegion.Address(False, False)
    Else
        out = ws.AutoFilter.Range.Address(False, False) & " cr=" & ws.Range("A1").CurrentRegion.Address(False, False)
    End If
    wb.Close False
    Go = out
End Function
'''


def first_look() -> list[list[object]]:
    cases: list[list[object]] = []
    for blank in ('', 'C1:C7', 'B1:B7', 'C2:C7', 'B2:C7'):
        for extra in ('', 'D8', 'E8', 'D1:D7;E8', 'D7;E8', 'C8', 'D3;E8', 'A8', 'B9', 'D1:D7;E8:E12', 'E1:E12', 'D1:D7', 'C8:C9;E10'):
            cases.append([blank, extra, True])
    return cases


def held_out(round_: str) -> list[list[object]]:
    chooser = random.Random('filter-growth' + round_)
    cells = [f'{column}{row}' for column in 'ABCDEFG' for row in range(1, 15)]
    cases: list[list[object]] = []
    while len(cases) < 120:
        blank = chooser.choice(('', '', 'C1:C7', 'B1:B7', 'A2:A7', 'C2:C7', 'B3:C7', 'A1:A7', 'C4'))
        extra = ';'.join(chooser.sample(cells, chooser.randint(0, 4)))
        for typed in (True, False):
            cases.append([blank, extra, typed])
    return cases


def _ask(cases: list[list[object]]) -> list[str]:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for case in cases:
            result = excel.run_vba(SOURCE, 'Go', tuple(case), timeout=60.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs else f'ERROR {str(result)[:200]}')
    return answers


def _write(path: Path, cases: list[list[object]], answers: list[str]) -> None:
    path.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': cases, 'answers': answers}, indent=0),
                    encoding='utf-8', newline='\n')


def predict(round_: str) -> None:
    from pyopenvba.apps.excel import ExcelApplication

    found = []
    for case in held_out(round_):
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(SOURCE, name='Probe')
        found.append({'call': case, 'model': str(app.run('Probe.Go', *case))})
    (FIXTURES / f'filter_range_growth{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted')


def measure(round_: str) -> None:
    predicted = json.loads((FIXTURES / f'filter_range_growth{round_}_predictions.json').read_text('utf-8'))['cases']
    cases = [one['call'] for one in predicted]
    answers = _ask(cases)
    _write(FIXTURES / f'filter_range_growth{round_}.json', cases, answers)
    wrong = [(one['call'], one['model'], answer) for one, answer in zip(predicted, answers) if one['model'] != answer]
    print(f'{len(answers) - len(wrong)} right, {len(wrong)} wrong of {len(answers)}')
    for one in wrong[:20]:
        print(one)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        found = first_look()
        _write(FIXTURES / 'filter_range_growth.json', found, _ask(found))
        print(len(found), 'measured')
    else:
        {'predict': predict, 'measure': measure}[sys.argv[1]](sys.argv[2])
