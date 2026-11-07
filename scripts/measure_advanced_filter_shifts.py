"""What inserts and deletes -- Range.Insert and Range.Delete, of cells, rows or columns -- do on a sheet with an
advanced filter.

Each case builds a new workbook: a header and eleven rows in A1:C12 (Name, Num 1 to 11, Cat x, y and z in turn),
markers in columns E and Z, a criteria range at G1, then one AdvancedFilter: in place, or copying to J1. Then one
edit. It reads FilterMode, which of rows 1 to 20 are hidden, the sheet's _FilterDatabase, Criteria and Extract names,
cells A1:D14, G1:H4 and J1:L8, the markers and any error, then saves the book as .xlsx, opens it again and reads the
same. Each case runs typed (Dim ws As Worksheet) and late (As Object).

    python scripts/measure_advanced_filter_shifts.py [round]    # a first look: '' for CASES, _b for the second
    python scripts/measure_advanced_filter_shifts.py predict 2  # tests/fixtures/advanced_filter_shifts2_predictions.json
    python scripts/measure_advanced_filter_shifts.py 2          # asks Excel the cases written there and scores them

writes tests/fixtures/advanced_filter_shifts<round>.json.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Named(ws As Object, ByVal name As String) As String
    On Error Resume Next
    Named = "-"
    Named = ws.Names(name).RefersTo
End Function

Private Function Dump(ws As Object) As String
    Dim out As String, i As Long, c As Object
    On Error Resume Next
    out = "fm=" & ws.FilterMode & " auto=" & ws.AutoFilterMode & " hidden="
    For i = 1 To 20
        If ws.Rows(i).Hidden Then out = out & i & ","
    Next
    out = out & " db=" & Named(ws, "_FilterDatabase") & " crit=" & Named(ws, "Criteria") & " ext=" & Named(ws, "Extract")
    out = out & " cells="
    For Each c In ws.Range("A1:D14").Cells
        out = out & c.Formula & ","
    Next
    out = out & " g="
    For Each c In ws.Range("G1:H4").Cells
        out = out & c.Formula & ","
    Next
    out = out & " j="
    For Each c In ws.Range("J1:L8").Cells
        out = out & c.Formula & ","
    Next
    out = out & " e="
    For i = 1 To 20
        out = out & ws.Cells(i, 5).Formula & ","
    Next
    out = out & " z="
    For i = 1 To 20
        out = out & ws.Cells(i, 26).Formula & ","
    Next
    Dump = out
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
    For i = 1 To 20
        ws.Cells(i, 5).Value = "e" & i
        ws.Cells(i, 26).Value = "z" & i
    Next
    Select Case kind
        Case "greater"
            ws.Range("G1:G2").Value = Application.Transpose(Array("Num", ">5"))
            ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:G2")
        Case "or"
            ws.Range("G1:G3").Value = Application.Transpose(Array("Cat", "x", "z"))
            ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:G3")
        Case "two"
            ws.Range("G1:H1").Value = Array("Num", "Cat")
            ws.Range("G2:H2").Value = Array(">3", "y")
            ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:H2")
        Case "all"
            ws.Range("G1:G2").Value = Application.Transpose(Array("Num", ">0"))
            ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:G2")
        Case "copy"
            ws.Range("G1:G2").Value = Application.Transpose(Array("Num", ">5"))
            ws.Range("A1:C12").AdvancedFilter xlFilterCopy, ws.Range("G1:G2"), ws.Range("J1")
        Case "nomatch"
            ws.Range("G1:G2").Value = Application.Transpose(Array("Num", ">100"))
            ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:G2")
        Case "far": ws.Range("AC1").Value = "far"
        Case "low": ws.Range("AC30").Value = "low"
        Case "near": ws.Range("AA14").Value = "near"
        Case "fmtnear": ws.Range("AA14").Interior.Color = RGB(255, 0, 0)
        Case "fmtlow": ws.Range("F22").Interior.Color = RGB(255, 0, 0)
        Case "textnear": ws.Range("AA14").Formula = "="""""
        Case "textlow": ws.Range("F22").Formula = "="""""
        Case "distinct": ws.Range("A1:C12").AdvancedFilter xlFilterInPlace, , , True
        Case "repeats": ws.Range("C1:C12").AdvancedFilter xlFilterInPlace, , , True
        Case "allrepeats"
            ws.Range("G1:G2").Value = Application.Transpose(Array("Num", ">0"))
            ws.Range("C1:C12").AdvancedFilter xlFilterInPlace, ws.Range("G1:G2"), , True
    End Select
    If edit = "none" Then
        Go = Dump(ws)
        wb.Close False
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    Dim target As String, how As String, parts() As String
    parts = Split(edit, " ")
    target = parts(1)
    how = parts(2)
    If parts(0) = "showall" Then
        If typed Then wt.ShowAllData Else ws.ShowAllData
    ElseIf how = "auto" And typed Then
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
    out = "err=" & Err.Number & ":" & Err.Description
    If UBound(parts) = 3 Then
        ' "... show": ShowAllData after the edit.
        Err.Clear
        If typed Then wt.ShowAllData Else ws.ShowAllData
        out = out & " show=" & Err.Number & ":" & Err.Description
    End If
    out = out & " " & Dump(ws)
    Err.Clear
    path = Environ("TEMP") & "\pyovba_advanced_filter_shift.xlsx"
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
KINDS = ('greater', 'or', 'two', 'all', 'copy')
EDITS = ('insert A5:C5 down', 'delete A5:C5 up', 'delete A7:C7 up', 'insert A1:C1 down', 'delete A1:C1 up', 'insert A14:C14 down',
         'delete A13:C13 up', 'insert A1:A12 right', 'delete B1:B12 left', 'insert D1:D12 right', 'insert Z5 down', 'delete Z5 up',
         'delete E5 up', 'insert E5 down', 'insert G1 down', 'delete G2 up', 'insert G1:G3 right', 'delete F1:F4 left',
         'insert A5:B5 down', 'delete A5 left', 'delete A5:C5 auto', 'insert A5:C5 auto',
         'insert 5:5 auto', 'delete 5:5 auto', 'delete 7:7 auto', 'insert 1:1 auto', 'insert 20:20 auto', 'delete 2:12 auto',
         'insert B:B auto', 'delete B:B auto', 'delete F:F auto', 'insert H:H auto')
CASES: list[list[object]] = [[kind, edit, typed] for kind in KINDS for edit in EDITS for typed in (True, False)]
#: The second look: filters in place that hide nothing, every row or some, with Unique and with criteria, read as
#: AdvancedFilter leaves them ("none") and through a few edits.
KINDS_B = ('all', 'nomatch', 'distinct', 'repeats', 'allrepeats')
EDITS_B = ('none', 'insert A14:C14 down', 'insert A5:C5 down', 'delete A7:C7 up', 'delete B1:B12 left', 'insert 5:5 auto',
           'delete 5:5 auto', 'delete 2:12 auto')
CASES_B: list[list[object]] = [[kind, edit, typed] for kind in KINDS_B for edit in EDITS_B for typed in (True, False)]
#: The third look: ShowAllData alone and after an edit, on filters that hide rows, hide none and copy.
KINDS_C = ('greater', 'all', 'nomatch', 'distinct', 'repeats', 'copy')
EDITS_C = ('showall A1 auto', 'delete A7:C7 up show', 'delete 5:5 auto show', 'insert 5:5 auto show', 'delete B1:B12 left show',
           'insert B:B auto show', 'delete 2:12 auto show', 'delete A5:C5 up show')
CASES_C: list[list[object]] = [[kind, edit, typed] for kind in KINDS_C for edit in EDITS_C for typed in (True, False)]
#: The fourth look: which way a Delete or Insert with no Shift goes where a tall range has nothing to its right, on a
#: sheet with no filter ('plain') and one an extracting filter left its names on.
KINDS_D = ('plain', 'copy')
EDITS_D = ('delete Z12:Z16 auto', 'delete E12:E16 auto', 'delete Y12:Y16 auto', 'delete AA12:AA16 auto', 'delete Z2:Z3 auto',
           'delete E2:E3 auto', 'delete Z12:AA16 auto', 'insert Z12:Z16 auto', 'insert AA12:AA16 auto', 'delete Z21:Z25 auto',
           'delete C2:C12 auto', 'delete Z1:Z20 auto')
CASES_D: list[list[object]] = [[kind, edit, typed] for kind in KINDS_D for edit in EDITS_D for typed in (True, False)]
#: The fifth look: what "nothing to its right" is, with a cell further right in another row ('far', AC1), below the
#: markers ('low', AC30) or beside the range ('near', AA14); and whether a wide range with nothing below goes left.
KINDS_E = ('plain', 'far', 'low', 'near')
EDITS_E = ('delete Z12:Z16 auto', 'delete Z2:Z3 auto', 'delete F20:Y20 auto', 'delete Z12:AA16 auto', 'delete Y12:Y16 auto',
           'delete F19:Y19 auto')
CASES_E: list[list[object]] = [[kind, edit, typed] for kind in KINDS_E for edit in EDITS_E for typed in (True, False)]
#: The sixth look: whether a cell with only a fill, or a formula showing "", counts as lying to the right or below.
KINDS_F = ('fmtnear', 'fmtlow', 'textnear', 'textlow')
EDITS_F = ('delete Z12:Z16 auto', 'delete F20:Y20 auto', 'delete Z2:Z3 auto')
CASES_F: list[list[object]] = [[kind, edit, typed] for kind in KINDS_F for edit in EDITS_F for typed in (True, False)]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    stem = 'advanced_filter_shifts' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    cases: list[dict[str, Any]] = json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists() else [{'call': one} for one in {'_b': CASES_B, '_c': CASES_C, '_d': CASES_D, '_e': CASES_E, '_f': CASES_F}.get(stem[len('advanced_filter_shifts'):], CASES)]
    answers: list[str] = []
    crashes: list[list[object]] = []
    number = 0
    while number < len(cases):
        # A crash is recorded, a new Excel takes the case again once, and the run goes on.
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                while number < len(cases):
                    case = cases[number]
                    result = excel.run_vba(SOURCE, 'Go', tuple(case['call']), timeout=120.0, line_numbers=False)
                    answers.append(str(result.value) if result.ok and not result.dialogs
                                   else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
                    number += 1
        except HarnessError as error:
            crashes.append([number, cases[number]['call'], str(error)[:200]])
            print('crash at case', number, cases[number]['call'], file=sys.stderr, flush=True)
            if len(crashes) > 1 and crashes[-2][0] == number:
                answers.append('CRASH twice')
                number += 1
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE,
                                                       'cases': [case['call'] for case in cases], 'answers': answers, 'crashes': crashes},
                                                      indent=0), encoding='utf-8', newline='\n')
    if predicted.exists():
        wrong = [(case['call'], case['model'], answer) for case, answer in zip(cases, answers, strict=True) if case['model'] is not None and case['model'] != answer]
        refused = sum(case['model'] is None for case in cases)
        print(f"{len(answers) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(answers)}")
        for one in wrong[:20]:
            print(str(one)[:600])
    else:
        print(len(answers), 'asked,', sum(answer.startswith(('ERROR', 'CRASH')) for answer in answers), 'errors')


def held_out(round_: str) -> list[list[object]]:
    """Fresh edits, chosen at random for the round: cells, rows and columns near the list and the criteria, each kind."""
    import random

    chooser = random.Random('advanced-filter-shifts' + round_)
    columns = 'ABCDEFGHJZ'
    cases: list[list[object]] = []
    while len(cases) < 200:
        kind = chooser.choice((*KINDS, 'nomatch', 'distinct', 'repeats'))
        verb = chooser.choice(('delete', 'delete', 'insert'))
        shape = chooser.choice(('cells', 'cells', 'cells', 'rows', 'columns'))
        if shape == 'rows':
            top = chooser.choice((1, 2, 3, 5, 7, 11, 12, 13, 18))
            edit = f'{verb} {top}:{top + chooser.choice((0, 0, 1, 3))} auto'
        elif shape == 'columns':
            left = chooser.randrange(len(columns) - 1)
            edit = f'{verb} {columns[left]}:{columns[min(len(columns) - 2, left + chooser.choice((0, 0, 1)))]} auto'
        else:
            left = chooser.randrange(len(columns))
            right = min(len(columns) - 1, left + chooser.choice((0, 0, 1, 2)))
            top = chooser.choice((1, 2, 3, 5, 7, 11, 12, 13, 18))
            bottom = top + chooser.choice((0, 0, 1, 2, 4, 11))
            how = chooser.choice(('up', 'left', 'auto')) if verb == 'delete' else chooser.choice(('down', 'right', 'auto'))
            edit = f'{verb} {columns[left]}{top}:{columns[right]}{bottom} {how}'
        if chooser.random() < 0.2:
            edit += ' show'
        for typed in (True, False):
            cases.append([kind, edit, typed])
    return cases


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


def predict(round_: str) -> None:
    found = [{'call': case, 'model': model_answer(case)} for case in {'_c': CASES_C, '_d': CASES_D, '_e': CASES_E, '_f': CASES_F}.get(round_) or held_out(round_)]
    (FIXTURES / f'advanced_filter_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                               encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
