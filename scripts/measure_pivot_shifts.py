"""Inserts and deletes of cells, rows and columns beside, through and across a pivot table: where Excel moves the
table, refuses the edit or changes it, what a PivotTable object held from before the edit and one fetched again read,
where the cache's source goes, and what a refresh, a save and a reopen leave.

Each case builds a new workbook with a source list (k: a to d, v: 1 to 4, g: x or y) and one layout:

- other: the list on Main at A1:C5, the table on Second at C3, rows by k, sum of v;
- same: the list and the table on one sheet, Second, the list at A1:C5 and the table at E3;
- page: as other, the table at C5 with g as a page field above it;
- source: as other, the edit made on Main, where the list is.

Then one edit, on Second (on Main for source), "verb target how". It reads the error, then from the PivotTable held before the edit and
from Second.PivotTables(1) fetched again: TableRange1, TableRange2, DataBodyRange, RowRange and SourceData, the cells
A1:H12 of Second, then after RefreshTable the same again, then saves the book as .xlsx, opens it again and reads
the table once more. Each case runs typed (Dim ws As Worksheet) and late (As Object).

    python scripts/measure_pivot_shifts.py [_a | _b]   # the development looks
    python scripts/measure_pivot_shifts.py predict 1   # tests/fixtures/pivot_shifts1_predictions.json
    python scripts/measure_pivot_shifts.py 1           # Excel answers the round predicted

writes tests/fixtures/pivot_shifts<round>.json, which tests/test_excel_pivot_shifts.py replays.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Table(pt As Object) As String
    Dim s As String, v As String
    On Error Resume Next
    v = "-": v = pt.TableRange1.Address(False, False): s = s & "t1=" & v
    v = "-": v = pt.TableRange2.Address(False, False): s = s & " t2=" & v
    v = "-": v = pt.DataBodyRange.Address(False, False): s = s & " body=" & v
    v = "-": v = pt.RowRange.Address(False, False): s = s & " rows=" & v
    v = "-": v = pt.SourceData: s = s & " src=" & v
    Table = s
End Function

Private Function Cells(ws As Object) As String
    Dim r As Long, c As Long, out As String
    On Error Resume Next
    For r = 1 To 12
        For c = 1 To 8
            out = out & ws.Cells(r, c).Formula & ","
        Next
        out = out & "/"
    Next
    Cells = out
End Function

Public Function Go(ByVal kind As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Workbook, src As Object, ws As Object, wt As Worksheet, held As Object, out As String, path As String
    Dim r As Long, at As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set src = wb.Worksheets(1): src.Name = "Main"
    Set ws = wb.Worksheets.Add(After:=src): ws.Name = "Second"
    If kind = "same" Then Set src = ws
    Set wt = ws
    If kind = "source" Then Set wt = src
    src.Range("A1:C1").Value = Array("k", "v", "g")
    For r = 2 To 5
        src.Cells(r, 1).Value = Chr(95 + r): src.Cells(r, 2).Value = r - 1: src.Cells(r, 3).Value = IIf(r Mod 2 = 0, "x", "y")
    Next
    at = "C3"
    If kind = "same" Then at = "E3"
    If kind = "page" Then at = "C5"
    Set held = wb.PivotCaches.Create(xlDatabase, src.Range("A1:C5")).CreatePivotTable(ws.Range(at), "PT")
    held.PivotFields("k").Orientation = xlRowField
    held.AddDataField held.PivotFields("v"), "Sum of v", xlSum
    If kind = "page" Then held.PivotFields("g").Orientation = xlPageField
    wt.Activate
    Dim es As Object
    Set es = wt
    Dim parts() As String, target As String, how As String
    parts = Split(edit, " ")
    target = parts(1)
    how = parts(2)
    On Error Resume Next
    Err.Clear
    If how = "auto" And typed Then
        If parts(0) = "insert" Then wt.Range(target).Insert Else wt.Range(target).Delete
    ElseIf how = "auto" Then
        If parts(0) = "insert" Then es.Range(target).Insert Else es.Range(target).Delete
    ElseIf typed Then
        If parts(0) = "insert" Then
            wt.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            wt.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    Else
        If parts(0) = "insert" Then
            es.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            es.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    End If
    out = "err=" & Err.Number & ":" & Err.Description
    Err.Clear
    out = out & " n=" & ws.PivotTables.Count & " held[" & Table(held) & "]"
    If ws.PivotTables.Count > 0 Then out = out & " again[" & Table(ws.PivotTables(1)) & "]"
    out = out & " " & Cells(ws)
    Err.Clear
    If ws.PivotTables.Count > 0 Then
        ws.PivotTables(1).RefreshTable
        out = out & " refresh=" & Err.Number & " [" & Table(ws.PivotTables(1)) & "] " & Cells(ws)
    End If
    Err.Clear
    path = Environ("TEMP") & "\pyovba_pivot_shift.xlsx"
    Kill path
    Err.Clear
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    Application.DisplayAlerts = True
    out = out & " save=" & Err.Number
    wb.Close False
    Set wb = Workbooks.Open(path)
    Set ws = wb.Worksheets("Second")
    out = out & " ~ n=" & ws.PivotTables.Count
    If ws.PivotTables.Count > 0 Then out = out & " [" & Table(ws.PivotTables(1)) & "]"
    wb.Close False
    Kill path
    Go = out
End Function
'''
EDITS = {
    'other': ('insert 1:1 auto', 'insert 1:2 auto', 'delete 1:1 auto', 'delete 2:2 auto', 'insert 3:3 auto',
              'insert 4:4 auto', 'insert 8:8 auto', 'insert 9:9 auto', 'delete 4:4 auto', 'delete 8:8 auto',
              'delete 3:8 auto', 'delete 2:9 auto', 'insert A:A auto', 'delete A:A auto', 'insert C:C auto',
              'insert D:D auto', 'insert E:E auto', 'delete C:C auto', 'delete D:D auto', 'delete C:D auto',
              'insert A1:H1 down', 'insert C1:D1 down', 'insert B1:E1 down', 'insert C1 down', 'insert D2 down',
              'insert A3 right', 'insert A3:A8 right', 'insert A1:A12 right', 'insert A4:B4 right',
              'delete A1:H1 up', 'delete C1:D1 up', 'delete C2 up', 'delete B3:B8 left', 'delete A3:B8 left',
              'insert C9:D9 down', 'insert E3 right', 'delete C4:D4 up', 'delete C4 up', 'insert C4:D4 down',
              'insert C3:D8 down', 'delete C3:D8 up', 'insert 1:1 down', 'insert C2:D2 auto', 'delete A1:B2 auto'),
    'same': ('insert 1:1 auto', 'insert 3:3 auto', 'insert 4:4 auto', 'delete 3:3 auto', 'delete 2:2 auto',
             'insert A:A auto', 'insert B:B auto', 'delete B:B auto', 'insert E:E auto', 'insert D:D auto',
             'insert A2:C2 down', 'delete A2:C2 up', 'insert A1:F1 down', 'insert E1:F1 down',
             'insert A6:C6 down', 'delete A5:C5 up'),
    'page': ('insert 1:1 auto', 'insert 3:3 auto', 'insert 4:4 auto', 'delete 3:3 auto', 'delete 4:4 auto',
             'insert C1:D1 down', 'insert C3 down', 'insert C4:D4 down', 'delete C4:D4 up', 'insert A3 right',
             'insert A3:A4 right', 'insert A1:A12 right'),
}
CASES = [[kind, edit, typed] for kind, edits in EDITS.items() for edit in edits for typed in (True, False)]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    stem = 'pivot_shifts' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    looks = {'_a': [case for case in CASES if case[0] == 'other'],
             '_b': [case for case in CASES if case[0] != 'other']}
    cases: list[dict[str, Any]] = (json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists()
                                   else [{'call': one} for one in looks.get(stem[len('pivot_shifts'):], CASES)])
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


def held_out(round_: str) -> list[list[object]]:
    """Fresh cases for the round: a layout, then whole rows or columns, or a block of cells, going in or out with a
    shift given or left to Excel, drawn over A1:H12, typed and late."""
    import random

    pick = random.Random('pivot_shifts' + round_)
    cases: list[list[object]] = []
    for _ in range(60):
        kind = pick.choice(['other', 'same', 'page', 'source'])
        verb = pick.choice(['insert', 'delete'])
        shape = pick.random()
        if shape < 0.3:
            first = pick.randint(1, 11)
            target, how = f'{first}:{first + pick.choice([0, 0, 1, 2])}', 'auto'
        elif shape < 0.5:
            first = pick.randint(1, 7)
            target = f'{"ABCDEFGH"[first - 1]}:{"ABCDEFGH"[first - 1 + pick.choice([0, 0, 1])]}'
            how = 'auto'
        else:
            top, left = pick.randint(1, 11), pick.randint(1, 7)
            bottom, right = min(12, top + pick.choice([0, 0, 1, 2, 5, 9])), min(8, left + pick.choice([0, 0, 1, 2, 4]))
            target = f'{"ABCDEFGH"[left - 1]}{top}' + ('' if (top, left) == (bottom, right) else
                                                       f':{"ABCDEFGH"[right - 1]}{bottom}')
            how = pick.choice(['auto', 'down' if verb == 'insert' else 'up', 'right' if verb == 'insert' else 'left'])
        cases.append([kind, f'{verb} {target} {how}', pick.random() < 0.5])
    return cases


def predict(round_: str) -> None:
    found = [{'call': case, 'model': model_answer(case)} for case in held_out(round_)]
    (FIXTURES / f'pivot_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                     encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


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


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
