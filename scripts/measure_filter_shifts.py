"""What a cell shift -- Range.Insert or Range.Delete with a Shift -- does to a sheet's AutoFilter that has criteria.

Each case builds a new workbook: a header and eleven rows in A1:C12 (Name, Num 1 to 11, Cat x, y and z in turn), an
AutoFilter over A1:C12 with one kind of criteria (or none, or rows hidden by hand), markers in columns E and Z, then
one shift. It reads the filter's range, each filter's criteria, which of rows 1 to 16 are hidden, FilterMode, the
hidden _FilterDatabase name, cells A1:D14, the markers and any error, then saves the book as .xlsx, opens it again and
reads the same. Each case runs typed (Dim ws As Worksheet) and late (As Object).

    python scripts/measure_filter_shifts.py [round]    # a first look: '' for CASES, _b, _c or _d for the later ones
    python scripts/measure_filter_shifts.py predict 2  # tests/fixtures/filter_shifts2_predictions.json, from the model
    python scripts/measure_filter_shifts.py 2          # asks Excel the cases written there and scores them

writes tests/fixtures/filter_shifts<round>.json. The first look (no markers) and _b were measured before the markers
reached row 25; each fixture keeps the source it was measured with.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Dump(ws As Object) As String
    Dim out As String, i As Long, c As Object, f As Object, one As String
    On Error Resume Next
    out = "mode=" & ws.AutoFilterMode & " fm=" & ws.FilterMode
    If ws.AutoFilterMode Then
        out = out & " range=" & ws.AutoFilter.Range.Address(False, False)
        For i = 1 To ws.AutoFilter.Filters.Count
            Set f = ws.AutoFilter.Filters(i)
            If f.On Then
                one = ""
                one = Show(f.Criteria1)
                out = out & " f" & i & "=" & f.Operator & ":" & one
                one = ""
                one = Show(f.Criteria2)
                If one <> "" Then out = out & "|" & one
            End If
        Next
    End If
    out = out & " hidden="
    For i = 1 To 16
        If ws.Rows(i).Hidden Then out = out & i & ","
    Next
    one = ""
    one = ws.Names("_FilterDatabase").RefersTo
    out = out & " db=" & one & " cells="
    For Each c In ws.Range("A1:D14").Cells
        out = out & c.Formula & ","
    Next
    out = out & " e="
    For i = 1 To 25
        out = out & ws.Cells(i, 5).Formula & ","
    Next
    out = out & " z="
    For i = 1 To 25
        out = out & ws.Cells(i, 26).Formula & ","
    Next
    Dump = out
End Function

Private Function Show(v As Variant) As String
    Dim i As Long, out As String
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            out = out & "[" & v(i) & "]"
        Next
        Show = out
    Else
        Show = CStr(v)
    End If
End Function

Public Function Go(ByVal kind As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Object, ws As Object, wt As Worksheet, i As Long, out As String, path As String
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set wt = ws
    ws.Range("A1:C1").Value = Array("Name", "Num", "Cat")
    For i = 2 To 12
        ws.Cells(i, 1).Value = "n" & (i - 1)
        ws.Cells(i, 2).Value = i - 1
        ws.Cells(i, 3).Value = Mid("xyz", ((i - 2) Mod 3) + 1, 1)
    Next
    For i = 1 To 25
        ws.Cells(i, 5).Value = "e" & i
        ws.Cells(i, 26).Value = "z" & i
    Next
    Select Case kind
        Case "values": ws.Range("A1:C12").AutoFilter Field:=3, Criteria1:=Array("x", "z"), Operator:=xlFilterValues
        Case "greater": ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=">5"
        Case "or": ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:="<3", Operator:=xlOr, Criteria2:=">9"
        Case "top": ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:="3", Operator:=xlTop10Items
        Case "two"
            ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=">5"
            ws.Range("A1:C12").AutoFilter Field:=3, Criteria1:="x"
        Case "above": ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic
        Case "none": ws.Range("A1:C12").AutoFilter Field:=2, Criteria1:=">0"
        Case "plain": ws.Range("A1:C12").AutoFilter
        Case "hidden"
            ws.Range("A1:C12").AutoFilter
            ws.Rows(5).Hidden = True
            ws.Rows(6).Hidden = True
    End Select
    On Error Resume Next
    Err.Clear
    Dim target As String, how As String, parts() As String, place As Object, placeTyped As Worksheet
    parts = Split(edit, " ")
    target = parts(1)
    how = parts(2)
    Set place = ws
    If Left(parts(0), 5) = "other" Then
        Set place = wb.Worksheets.Add(After:=ws)
        place.Range("A1:C3").Value = 7
        parts(0) = Mid(parts(0), 6)
    End If
    Set placeTyped = place
    If how = "auto" And typed Then
        If parts(0) = "insert" Then placeTyped.Range(target).Insert Else placeTyped.Range(target).Delete
    ElseIf how = "auto" Then
        If parts(0) = "insert" Then place.Range(target).Insert Else place.Range(target).Delete
    ElseIf typed Then
        If parts(0) = "insert" Then
            placeTyped.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            placeTyped.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    Else
        If parts(0) = "insert" Then
            place.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            place.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    End If
    out = "err=" & Err.Number & ":" & Err.Description & " " & Dump(ws)
    Err.Clear
    path = Environ("TEMP") & "\pyovba_filter_shift.xlsx"
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
KINDS = ('values', 'greater', 'or', 'top', 'two', 'above')
EDITS = ('insert A1:C1 down', 'insert A5:C5 down', 'delete A5:C5 up', 'insert A3:C4 down', 'delete A3:C4 up', 'insert A14:C14 down',
         'insert A1:A12 right', 'insert B1:B12 right', 'delete B1:B12 left', 'delete C1:C12 left', 'insert A5:B5 down', 'delete A5:B5 up',
         'insert A1:E1 down', 'delete A1:C1 up', 'delete A1:C12 up', 'insert A13:C13 down', 'insert D1:D12 right', 'delete A2:C2 up')
#: The second look: where an insert fails, deletes outside the filter, deletes over hidden and shown rows, and filters with no
#: criteria or with criteria that hide nothing.
KINDS_B = ('greater', 'values', 'none', 'plain', 'hidden')
EDITS_B = ('insert Z100 down', 'insert Z100 right', 'delete Z100 up', 'delete Z100 left', 'insert A14:C14 down', 'insert E3 down', 'delete E3 up',
           'delete A5:C8 up', 'delete A6:C9 up', 'delete A2:A4 left', 'delete B5:B8 left', 'delete A20:C20 up', 'delete D1:D12 left',
           'delete A5:D5 up', 'delete A13:C13 up', 'delete B5:B6 up', 'delete A7:B8 up', 'otherinsert A1:C1 down', 'otherdelete A1:C1 up',
           'insert A5:C5 down', 'delete A5:C5 up')
#: The third look: how far a delete up reaches across the sheet, a delete left outside the filter's columns, targets that
#: run past the filter's rows or take in its header, and an edit with no Shift.
KINDS_C = ('greater', 'values', 'none', 'plain')
EDITS_C = ('delete E3 up', 'delete Z3 up', 'delete A7:B8 up', 'delete A5:C8 up', 'delete E5:E8 left', 'delete Z5:Z8 left', 'delete B5:B8 left',
           'delete A11:C14 up', 'delete A12:C13 up', 'delete A1:C3 up', 'delete A1:A3 left', 'delete A13:C14 up', 'delete E13 up',
           'delete A5:C5 auto', 'delete A5:A8 auto', 'insert A5:C5 auto', 'delete A2:C12 up', 'delete A5:C8 left', 'delete A6:A7 up')
#: The fourth look, markers down to row 25: a delete up beside or below the filter, and a delete left over whole filter columns.
KINDS_D = ('greater', 'values', 'none', 'plain')
EDITS_D = ('delete A20:C20 up', 'delete Z20 up', 'delete E14 up', 'delete A15:A16 up', 'delete D3 up', 'delete B1:B12 left', 'delete A1:A12 left',
           'delete A1:C12 left', 'delete A13:A14 left', 'delete A1:C1 left', 'delete A20:A21 left', 'delete E2:E12 left', 'delete A3:C3 up',
           'delete A7:C7 up', 'delete B2:B12 left', 'delete C2:C12 left', 'delete Z20 left')
CASES = [[kind, edit, typed] for kind in KINDS for edit in EDITS for typed in (True, False)]
CASES_D = [[kind, edit, typed] for kind in KINDS_D for edit in EDITS_D for typed in (True, False)]
#: The fifth look: inserts that cut through a filter with no criteria, and deletes that do the same, against whole-band ones.
KINDS_E = ('plain', 'hidden')
EDITS_E = ('insert A5:B5 down', 'insert B5:C5 down', 'insert A5 down', 'insert D5 down', 'insert A12 right', 'insert A5:A6 right',
           'insert B1:B5 right', 'insert A1:A11 right', 'insert A2:A12 right', 'insert A1:A12 right', 'insert A13 right', 'insert A5:C5 right',
           'insert E5 right', 'insert A1:C1 down', 'insert A1:B1 down', 'insert A13:C13 down', 'insert A12:C12 down', 'insert B12 down',
           'insert A1:A20 right', 'insert A3:C3 down', 'delete A5:B5 up', 'delete A12 left', 'delete B1:B5 left', 'insert A1 down')
CASES_E = [[kind, edit, typed] for kind in KINDS_E for edit in EDITS_E for typed in (True, False)]
#: The sixth look: where an insert through a filter with no criteria starts, and whether it takes the header row in.
EDITS_F = ('insert B5 right', 'insert C5 right', 'insert B5:C5 right', 'insert A5:B5 right', 'insert A5:D5 right', 'insert A1:C1 right',
           'insert B1 right', 'insert C1:C3 right', 'insert A1:A2 right', 'insert A11:A12 right', 'insert B12 right', 'insert D3 right',
           'insert A5:C6 right', 'insert Z5 right', 'insert B1 down', 'insert C1 down', 'insert B1:C1 down', 'insert A1:D1 down',
           'insert B2 down', 'insert A2:B2 down', 'insert A1:C2 down', 'insert B1:B3 down', 'insert D1 down', 'insert A1:A2 down',
           'insert B2:B12 right', 'insert C1:C12 right', 'insert B1:B12 right', 'insert A1:B5 right')
CASES_F = [['plain', edit, typed] for edit in EDITS_F for typed in (True, False)]
CASES_C = [[kind, edit, typed] for kind in KINDS_C for edit in EDITS_C for typed in (True, False)]
CASES_B = [[kind, edit, typed] for kind in KINDS_B for edit in EDITS_B for typed in (True, False)]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    stem = 'filter_shifts' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    cases = json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists() else [{'call': one} for one in (CASES_B if stem.endswith('_b') else CASES_C if stem.endswith('_c') else CASES_D if stem.endswith('_d') else CASES_E if stem.endswith('_e') else CASES_F if stem.endswith('_f') else CASES)]
    from pyvbaharness.results import HarnessError

    answers: list[str] = []
    crashes: list[list[object]] = []
    number = 0
    while number < len(cases):
        # Excel has ended mid-run here (case 27 of the third look, a late-bound Delete with no Shift, which passes alone):
        # a crash is recorded, a new Excel takes the case again once, and the run goes on.
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                while number < len(cases):
                    case = cases[number]
                    result = excel.run_vba(SOURCE, 'Go', tuple(case['call']), timeout=120.0, line_numbers=False)
                    answers.append(str(result.value) if result.ok and not result.dialogs else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
                    number += 1
        except HarnessError as error:
            crashes.append([number, cases[number]['call'], str(error)[:200]])
            print('crash at case', number, cases[number]['call'], file=sys.stderr, flush=True)
            if len(crashes) > 1 and crashes[-2][0] == number:
                answers.append('CRASH twice')
                number += 1
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE,
                                                         'cases': [case['call'] for case in cases], 'answers': answers, 'crashes': crashes}, indent=0),
                                           encoding='utf-8', newline='\n')
    if predicted.exists():
        wrong = [(case['call'], case['model'], answer) for case, answer in zip(cases, answers) if case['model'] is not None and case['model'] != answer]
        print(f"{len(answers) - len(wrong) - sum(case['model'] is None for case in cases)} right, {len(wrong)} wrong, "
              f"{sum(case['model'] is None for case in cases)} refused of {len(answers)}")
        for one in wrong[:20]:
            print(str(one)[:600])
    else:
        print(len(answers), 'asked,', sum(answer.startswith('ERROR') for answer in answers), 'errors')


def held_out(round_: str) -> list[list[object]]:
    """Fresh shifts, chosen at random for the round: inserts and deletes up and left anywhere near the filter, each kind."""
    import random

    chooser = random.Random('filter-shifts' + round_)
    columns = 'ABCDEFGZ'
    cases: list[list[object]] = []
    while len(cases) < 240:
        kind = chooser.choice(('values', 'greater', 'or', 'top', 'two', 'above', 'none', 'plain', 'hidden'))
        left = chooser.randrange(len(columns))
        right = min(len(columns) - 1, left + chooser.choice((0, 0, 1, 2, 3)))
        top = chooser.choice((1, 2, 3, 5, 7, 11, 12, 13, 18))
        bottom = top + chooser.choice((0, 0, 1, 2, 4, 11))
        verb = chooser.choice(('delete', 'delete', 'delete', 'insert'))
        how = chooser.choice(('up', 'left')) if verb == 'delete' else chooser.choice(('down', 'right'))
        edit = f'{verb} {columns[left]}{top}:{columns[right]}{bottom} {how}'
        for typed in (True, False):
            cases.append([kind, edit, typed])
    return cases


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
    (FIXTURES / f'filter_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0), encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
