"""Which way Range.Insert and Range.Delete with no Shift move cells beside sparklines.

Each case builds a new workbook, one sheet holding numbers in A1:D8 (row * 10 + column) and dates in A12:D12, and
one layout of sparklines:

- column: a group in F1:F4 reading A1:D4;
- row: a group in A10:D10 reading A1:D8;
- alone: one sparkline in H8 reading A1:D1, nothing near it;
- none: no sparkline.

A held-out round's layouts also hold values in a few cells chosen at random, named after the layout.

Then one edit with no Shift; in a held-out round ("... named") names on the range's first cell, the
cell below it and the one right of it go first, so which way cells went shows where no value moves. It reads A1:L20 cell by cell -- the value, and "*" where a sparkline sits -- and the
error, if any. tests/test_excel_sparkline_directions.py replays it.

    python scripts/measure_sparkline_directions.py [round]
    python scripts/measure_sparkline_directions.py predict 1   # tests/fixtures/sparkline_directions1_predictions.json
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Grid(ws As Object) As String
    Dim r As Long, c As Long, out As String, s As String, n As Long
    On Error Resume Next
    For r = 1 To 20
        For c = 1 To 12
            s = CStr(ws.Cells(r, c).Value)
            n = 0
            n = ws.Cells(r, c).SparklineGroups.Count
            If n > 0 Then s = s & "*"
            out = out & s & ","
        Next
        out = out & "/"
    Next
    Grid = out
End Function

Public Function Go(ByVal kind As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Workbook, ws As Object, wt As Worksheet, r As Long, c As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set wt = ws
    ws.Name = "S"
    ws.Activate
    For r = 1 To 8
        For c = 1 To 4
            ws.Cells(r, c).Value = r * 10 + c
        Next
    Next
    For c = 1 To 4
        ws.Cells(12, c).Value = c
    Next
    ' "column G3 K9": the layout, then a value in each cell named after it.
    Dim spots() As String, i As Long
    spots = Split(kind, " ")
    kind = spots(0)
    For i = 1 To UBound(spots)
        ws.Range(spots(i)).Value = "x"
    Next
    Select Case kind
        Case "column"
            ws.Range("F1:F4").SparklineGroups.Add xlSparkLine, "A1:D4"
        Case "row"
            ws.Range("A10:D10").SparklineGroups.Add xlSparkColumn, "A1:D8"
        Case "alone"
            ws.Range("H8").SparklineGroups.Add xlSparkLine, "A1:D1"
    End Select
    Dim parts() As String
    parts = Split(edit, " ")
    ' With "named": names on the range's first cell, the cell below it and the cell right of it, which show which way
    ' cells went where no value moves.
    Dim area As Range
    If UBound(parts) >= 2 Then
        Set area = ws.Range(parts(1))
        wb.Names.Add "First", "=S!" & area.Cells(1, 1).Address
        wb.Names.Add "Below", "=S!" & area.Cells(area.Rows.Count + 1, 1).Address
        wb.Names.Add "Beside", "=S!" & area.Cells(1, area.Columns.Count + 1).Address
    End If
    On Error Resume Next
    Err.Clear
    If typed Then
        If parts(0) = "insert" Then wt.Range(parts(1)).Insert Else wt.Range(parts(1)).Delete
    Else
        If parts(0) = "insert" Then ws.Range(parts(1)).Insert Else ws.Range(parts(1)).Delete
    End If
    Go = "err=" & Err.Number & " " & Grid(ws)
    If UBound(parts) >= 2 Then
        Go = Go & " names=" & wb.Names("First").RefersTo & " " & wb.Names("Below").RefersTo & " " & _
            wb.Names("Beside").RefersTo
    End If
    wb.Close False
End Function
'''
TARGETS = {
    'column': ('F2:F3', 'F4:F5', 'F5:F6', 'F6:F7', 'G2:G3', 'H2:H3', 'E2:E3', 'G5:G6', 'F4', 'F5', 'G4', 'E5:F5',
               'G4:H4', 'G5:H5', 'F1:F4', 'E1:E4', 'G1:G4', 'F2:G2'),
    'row': ('C9:C10', 'C8:C9', 'C11:C12', 'C13:C14', 'E9:E10', 'E10:E11', 'F9:F10', 'E10', 'E11', 'D11', 'A11:B11',
            'E10:F10', 'E11:F11', 'A13:B13', 'B10:C10', 'A9:D9', 'A11:D11', 'D9:D10'),
    'alone': ('H8', 'H7:H8', 'H8:H9', 'H6:H7', 'H9:H10', 'G7:G8', 'I7:I8', 'G8:G9', 'I8:I9', 'J7:J8', 'G8:H8',
              'H8:I8', 'G9:H9', 'G7:H7', 'I8:J8', 'F8:G8', 'H7:I7', 'H9:I9'),
}
CASES = [[kind, f'{verb} {target}', typed] for kind, targets in TARGETS.items()
         for target in targets for verb in ('insert', 'delete') for typed in (True, False)]
CASES += [['none', f'{verb} {target}', True] for kind in ('alone',) for target in TARGETS[kind] for verb in ('insert', 'delete')]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    stem = 'sparkline_directions' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    cases: list[dict[str, Any]] = (json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists()
                                   else [{'call': one} for one in CASES])
    answers: list[str] = []
    crashes: list[list[object]] = []
    number = 0
    while number < len(cases):
        # A crash is recorded, a new Excel takes the case again once, and the run goes on.
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                while number < len(cases):
                    result = excel.run_vba(SOURCE, 'Go', tuple(cases[number]['call']), timeout=120.0, line_numbers=False)
                    answers.append(str(result.value) if result.ok and not result.dialogs
                                   else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
                    number += 1
        except HarnessError as error:
            crashes.append([number, cases[number]['call'], str(error)[:200]])
            print('crash at case', number, cases[number]['call'], file=sys.stderr, flush=True)
            if len(crashes) > 1 and crashes[-2][0] == number:
                answers.append('CRASH twice')
                number += 1
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(),
                                                       'source': SOURCE, 'cases': [case['call'] for case in cases],
                                                       'answers': answers, 'crashes': crashes}, indent=0),
                                           encoding='utf-8', newline='\n')
    if predicted.exists():
        wrong = [(case['call'], case['model'], answer) for case, answer in zip(cases, answers, strict=True)
                 if case['model'] is not None and case['model'] != answer]
        refused = sum(case['model'] is None for case in cases)
        print(f"{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}")
        for one in wrong[:20]:
            print(str(one)[:600])
    else:
        print(len(answers), 'asked,', sum(answer.startswith(('ERROR', 'CRASH')) for answer in answers), 'errors')


def model_answer(case: list[object]) -> str | None:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='Probe')
    try:
        return str(app.run('Probe.Go', *case))
    except VBAUnsupportedError:
        return None


def _old_pulls_up(self: Any, area: Any, *, inserting: bool = False) -> bool:
    """The no-Shift rule before this look (#606): tall Inserts always go right, and where nothing lies the way the
    range's shape points, it goes the other way."""
    if area.rows <= area.columns:
        first = area.top if inserting else area.bottom + 1
        return any(row >= first and area.left <= column <= area.right and not cell.is_blank()
                   for (row, column), cell in self.sheet.cells_.items())
    return not inserting and not any(column > area.right and area.top <= row <= area.bottom and not cell.is_blank()
                                     for (row, column), cell in self.sheet.cells_.items())


def held_out(round_: str) -> list[list[object]]:
    """Fresh edits for the round, each layout with values in up to four random cells of A1:L20: half of them where
    the rule and the #606 rule answer apart, which a random draw seldom reaches, half drawn as they come. The ones
    apart are held out on sheets with no sparkline as much as beside them."""
    import random

    from pyopenvba.apps.excel._model import Range

    pick = random.Random('sparkline_directions' + round_)
    letters = 'ABCDEFGHIJKL'
    apart: list[list[object]] = []
    plain: list[list[object]] = []
    while len(apart) < 50 or len(plain) < 50:
        layout = pick.choice(('column', 'row', 'alone', 'none', 'none'))
        spots = sorted({f'{pick.choice(letters[4:])}{pick.randint(1, 20)}' for _ in range(pick.randint(0, 4))})
        kind = ' '.join([layout, *spots])
        top, left = pick.randint(1, 18), pick.randint(0, 10)
        bottom, right = top + pick.choice((0, 0, 1, 2)), min(11, left + pick.choice((0, 0, 1, 2)))
        target = f'{letters[left]}{top}' + ('' if (bottom, right) == (top, left) else f':{letters[right]}{bottom}')
        case: list[object] = [kind, f'{pick.choice(("insert", "delete"))} {target} named', True]
        new = model_answer(case)
        kept = Range._pulls_up
        Range._pulls_up = _old_pulls_up  # type: ignore[method-assign]
        try:
            old = model_answer(case)
        finally:
            Range._pulls_up = kept  # type: ignore[method-assign]
        if new != old and len(apart) < 50:
            apart.append(case)
        elif new == old and len(plain) < 50:
            plain.append(case)
    return [[kind, edit, typed] for kind, edit, _ in apart + plain for typed in (True, False)]


def predict(round_: str) -> None:
    found = [{'call': case, 'model': model_answer(case)} for case in held_out(round_)]
    (FIXTURES / f'sparkline_directions{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                             encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
