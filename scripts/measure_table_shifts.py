"""Inserts and deletes of cells, rows and columns beside and through tables: a filtered table, a table shifted by
cells moving into or out of it, and a table that another one's ListRows.Add or ListColumns.Add pushes along.

Each case builds a new workbook, one sheet with numbers in A1:G20 (A to C: row, row * 2, row * 3; E to G the same)
and text where a header row stands (A1:C1, A9:C9, E1:G1, E3:G3),
a table T1 over A1:C6, and one layout:

- one: nothing more;
- filtered: T1 filtered to column 2 above 6, hiding its first two data rows;
- below: a table T2 over A9:C12;
- right: a table T2 over E1:G5;
- beside: a table T2 over E3:G7.

Then one edit. It reads the error, each table's name, range, filter and filter mode, which of rows 1 to 20 are
hidden, and A1:H20, saves the book as .xlsx, opens it again and reads the same. Each case runs typed (Dim ws As
Worksheet) and late (As Object).

    python scripts/measure_table_shifts.py [round]
    python scripts/measure_table_shifts.py predict 1   # tests/fixtures/table_shifts1_predictions.json

writes tests/fixtures/table_shifts<round>.json, which tests/test_excel_table_shifts.py replays.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Dump(ws As Object) As String
    Dim out As String, t As Object, s As String, r As Long, c As Long
    On Error Resume Next
    For Each t In ws.ListObjects
        out = out & "[" & t.Name & " " & t.Range.Address(False, False)
        s = "-": s = t.ShowAutoFilter
        out = out & " af=" & s
        s = "-": s = t.AutoFilter.FilterMode
        out = out & " fm=" & s & "]"
    Next
    out = out & " hid="
    For r = 1 To 20
        If ws.Rows(r).Hidden Then out = out & r & ","
    Next
    out = out & " "
    For r = 1 To 20
        For c = 1 To 8
            out = out & ws.Cells(r, c).Formula & ","
        Next
        out = out & "/"
    Next
    Dump = out
End Function

Public Function Go(ByVal kind As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Workbook, ws As Object, wt As Worksheet, t As ListObject, out As String, path As String, r As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set wt = ws
    ws.Name = "S"
    ws.Activate
    For r = 1 To 20
        ws.Cells(r, 1).Value = r: ws.Cells(r, 2).Value = r * 2: ws.Cells(r, 3).Value = r * 3
        ws.Cells(r, 5).Value = r: ws.Cells(r, 6).Value = r * 2: ws.Cells(r, 7).Value = r * 3
    Next
    ' Text where a table's header row stands: a header of numbers names new columns a way not measured here.
    ws.Range("A1:C1").Value = Array("a", "b", "c")
    ws.Range("A9:C9").Value = Array("p", "q", "r")
    ws.Range("E1:G1").Value = Array("e", "f", "g")
    ws.Range("E3:G3").Value = Array("x", "y", "z")
    Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C6"), , xlYes)
    t.Name = "T1"
    Select Case kind
        Case "filtered"
            t.Range.AutoFilter 2, ">6"
        Case "below"
            ws.ListObjects.Add(xlSrcRange, ws.Range("A9:C12"), , xlYes).Name = "T2"
        Case "right"
            ws.ListObjects.Add(xlSrcRange, ws.Range("E1:G5"), , xlYes).Name = "T2"
        Case "beside"
            ws.ListObjects.Add(xlSrcRange, ws.Range("E3:G7"), , xlYes).Name = "T2"
    End Select
    Dim parts() As String, target As String, how As String
    parts = Split(edit, " ")
    On Error Resume Next
    Err.Clear
    Select Case parts(0)
        Case "rows+"
            If typed Then wt.ListObjects("T1").ListRows.Add Else ws.ListObjects("T1").ListRows.Add
        Case "rowsat+"
            If typed Then wt.ListObjects("T1").ListRows.Add CLng(parts(1)) Else ws.ListObjects("T1").ListRows.Add CLng(parts(1))
        Case "rows-"
            If typed Then wt.ListObjects("T1").ListRows(CLng(parts(1))).Delete Else ws.ListObjects("T1").ListRows(CLng(parts(1))).Delete
        Case "columns+"
            If typed Then wt.ListObjects("T1").ListColumns.Add Else ws.ListObjects("T1").ListColumns.Add
        Case "columns-"
            If typed Then wt.ListObjects("T1").ListColumns(CLng(parts(1))).Delete Else ws.ListObjects("T1").ListColumns(CLng(parts(1))).Delete
        Case Else
            target = parts(1)
            how = parts(2)
            If how = "auto" And typed Then
                If parts(0) = "insert" Then wt.Range(target).Insert Else wt.Range(target).Delete
            ElseIf how = "auto" Then
                If parts(0) = "insert" Then ws.Range(target).Insert Else ws.Range(target).Delete
            ElseIf typed Then
                If parts(0) = "insert" Then
                    wt.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
                Else
                    wt.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
                End If
            Else
                If parts(0) = "insert" Then
                    ws.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
                Else
                    ws.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
                End If
            End If
    End Select
    out = "err=" & Err.Number & ":" & Err.Description & " " & Dump(ws)
    Err.Clear
    path = Environ("TEMP") & "\pyovba_table_shift.xlsx"
    Kill path
    Err.Clear
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    Application.DisplayAlerts = True
    out = out & " save=" & Err.Number
    wb.Close False
    Set wb = Workbooks.Open(path)
    out = out & " ~ " & Dump(wb.Worksheets(1))
    wb.Close False
    Kill path
    Go = out
End Function
'''
EDITS = {
    'one': ('insert A3 down', 'insert A3:C3 down', 'delete A3:C3 up', 'insert B3:C3 down', 'insert A3:D3 down',
            'delete A3:C4 up', 'insert A1:C1 down', 'insert A1:C6 down', 'insert A7:C7 down', 'delete A7:C7 up',
            'insert A1:A6 right', 'insert A3:A4 right', 'delete D1:D6 left', 'insert A3:G3 down', 'insert A3 auto',
            'delete A3:C3 auto', 'insert A2:C2 auto', 'delete A2 up', 'insert D3 right', 'delete A1:C6 up'),
    'filtered': ('rows+ -', 'rowsat+ 2', 'rowsat+ 4', 'rows- 1', 'rows- 3', 'rows- 5', 'columns+ -', 'columns- 2',
                 'insert 3:3 auto', 'delete 4:4 auto', 'insert A8:C8 down', 'delete A3:C3 up'),
    'below': ('rows+ -', 'rowsat+ 2', 'rows- 2', 'insert A8:C8 down', 'delete A7:C7 up', 'delete A8:C8 up',
              'delete A7:C8 up', 'insert A9:C9 down', 'delete A9:C12 up', 'insert A9:C12 down', 'delete A8:C20 up',
              'insert A3:C3 down', 'insert A7:C9 down', 'delete A10:C10 up'),
    'right': ('columns+ -', 'columns- 2', 'insert D2:D3 right', 'insert D1:D5 right', 'insert D1:D6 right',
              'delete D1:D5 left', 'delete D1:D20 left', 'insert E1:E5 right', 'rows+ -', 'insert A3:C3 down'),
    'beside': ('rows+ -', 'rowsat+ 2', 'rows- 2', 'columns+ -', 'insert D3:D7 right', 'insert A6:C6 down'),
}
CASES = [[kind, edit, typed] for kind, edits in EDITS.items() for edit in edits for typed in (True, False)]
EDITS_C = {
    'one': ('delete B3 up', 'delete A3:D3 up', 'delete A3:G3 up', 'delete A1:C1 up', 'delete A1:A6 left',
            'delete B1:B6 left', 'delete A2:A3 left', 'delete B2:C2 up', 'delete A2:B3 up', 'delete C6 up',
            'delete A6:C7 up', 'delete A5:C8 up', 'insert A6:C7 down', 'insert A7:C7 auto', 'insert B3 auto',
            'delete B3 auto', 'delete A1:C6 left', 'insert A1:C6 right', 'insert C1:C6 right', 'insert B1:B6 right',
            'insert A3:C3 right', 'delete D1:D3 left', 'insert A1:C1 auto', 'delete A2:C2 left'),
    'filtered': ('delete A4:C5 up', 'delete A2:C4 up', 'delete A4 up', 'insert A4:C4 down', 'delete A2:C6 up',
                 'insert A7:C7 down', 'delete A4:C4 auto', 'delete 2:3 auto', 'delete A1:A6 left'),
    'below': ('delete A6:C6 up', 'insert A6:C6 down', 'delete A12:C12 up', 'delete A1:C12 up', 'insert A1:C12 down',
              'delete A9:C9 up', 'insert A10:C10 auto'),
    'right': ('delete C1:C6 left', 'insert C1:C6 right', 'delete E1:G5 left', 'insert A1:A6 right',
              'delete A1:A6 left'),
}
CASES_C = [[kind, edit, typed] for kind, edits in EDITS_C.items() for edit in edits for typed in (True, False)]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    stem = 'table_shifts' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    looks = {'_a': [case for case in CASES if case[0] in ('one', 'filtered')],
             '_b': [case for case in CASES if case[0] in ('below', 'right', 'beside')], '_c': CASES_C}
    cases: list[dict[str, Any]] = (json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists()
                                   else [{'call': one} for one in looks.get(stem[len('table_shifts'):], CASES)])
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


def held_out(round_: str) -> list[list[object]]:
    """Fresh edits for the round over every layout: ListRows and ListColumns, whole rows and columns, and blocks of
    cells in and beside the tables with each Shift and none; typed and late. A Delete with no Shift on the filtered
    table is left out: Excel asks first there, with DisplayAlerts on."""
    import random

    pick = random.Random('table_shifts' + round_)
    letters = 'ABCDEFGH'
    found: list[list[object]] = []
    while len(found) < 100:
        kind = pick.choice(tuple(EDITS))
        shape = pick.choice(('list', 'lines', 'cells', 'cells', 'cells'))
        if shape == 'list':
            verb = pick.choice(('rows+', 'rowsat+', 'rows-', 'columns+', 'columns-'))
            edit = f'{verb} ' + ('-' if verb in ('rows+', 'columns+') else str(pick.randint(1, 3 if 'columns' in verb else 5)))
        elif shape == 'lines':
            verb = pick.choice(('insert', 'delete'))
            if pick.random() < 0.5:
                top = pick.randint(1, 13)
                edit = f'{verb} {top}:{top + pick.randint(0, 1)} auto'
            else:
                left = pick.randint(0, 6)
                edit = f'{verb} {letters[left]}:{letters[left + pick.randint(0, 1)]} auto'
        else:
            verb = pick.choice(('insert', 'delete'))
            top, left = pick.randint(1, 13), pick.randint(0, 6)
            bottom, right = top + pick.choice((0, 0, 1, 2, 5)), min(7, left + pick.choice((0, 0, 1, 2)))
            target = f'{letters[left]}{top}' + ('' if (bottom, right) == (top, left) else f':{letters[right]}{bottom}')
            how = pick.choice(('auto', 'down', 'right') if verb == 'insert' else ('auto', 'up', 'left'))
            if kind == 'filtered' and verb == 'delete' and how == 'auto':
                continue
            edit = f'{verb} {target} {how}'
        if [kind, edit] not in found:
            found.append([kind, edit])
    return [[kind, edit, typed] for kind, edit in found for typed in (True, False)]


def predict(round_: str) -> None:
    found = [{'call': case, 'model': model_answer(case)} for case in held_out(round_)]
    (FIXTURES / f'table_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                     encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
