"""What inserts and deletes -- Range.Insert and Range.Delete, of cells, rows or columns -- do to sparklines.

Each case builds a new workbook with two sheets, S and T, each holding numbers in A1:D8 (row * 10 + column), and
one layout of sparkline groups on S:

- rows: a line group in F1:F4 reading A1:D4, a row each;
- cols: a column group in A10:D10 reading A1:D8, a column each;
- dates: the rows group with its date range A12:D12;
- other: the rows group reading T!A1:D4;
- two: the rows group and the cols group;
- abs: one sparkline in F6 reading $A$6:$D$6.

Then one edit, "verb target how", on S, or on T where the target is "T!...". It reads every group on S and T,
newest first: its type, location, source, date range and each sparkline's location and source, plus A1 and F1,
then saves the book as .xlsx, opens it again and reads the same. Each case runs typed (Dim ws As Worksheet) and
late (As Object).

    python scripts/measure_sparkline_shifts.py [round]    # a first look: '' for CASES
    python scripts/measure_sparkline_shifts.py predict 2  # tests/fixtures/sparkline_shifts2_predictions.json
    python scripts/measure_sparkline_shifts.py 2          # asks Excel the cases written there and scores them

writes tests/fixtures/sparkline_shifts<round>.json.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'

SOURCE = r'''Private Function Groups(ws As Object) As String
    Dim out As String, g As Object, i As Long, s As String
    On Error Resume Next
    out = ws.Name & "[" & ws.Cells.SparklineGroups.Count & "]"
    For Each g In ws.Cells.SparklineGroups
        s = "-": s = g.Type
        out = out & " {" & s
        s = "-": s = g.Location.Address(False, False)
        out = out & " at=" & s
        s = "-": s = g.SourceData
        out = out & " src=" & s
        s = "-": s = g.DateRange
        out = out & " dates=" & s
        For i = 1 To g.Count
            s = "-": s = g.Item(i).Location.Address(False, False)
            out = out & " " & s
            s = "-": s = g.Item(i).SourceData
            out = out & "=" & s
        Next
        out = out & "}"
        Err.Clear
    Next
    Groups = out
End Function

Private Function Dump(wb As Object) As String
    Dump = Groups(wb.Worksheets(1)) & " | " & Groups(wb.Worksheets(2)) & " | A1=" & wb.Worksheets(1).Range("A1").Formula & _
        " F1=" & wb.Worksheets(1).Range("F1").Formula
End Function

Public Function Go(ByVal kind As String, ByVal edit As String, ByVal typed As Boolean) As String
    Dim wb As Workbook, ws As Object, wt As Worksheet, other As Object, out As String, path As String
    Dim r As Long, c As Long, g As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "S"
    Set other = wb.Worksheets.Add(After:=ws)
    other.Name = "T"
    For r = 1 To 8
        For c = 1 To 4
            ws.Cells(r, c).Value = r * 10 + c
            other.Cells(r, c).Value = r * 10 + c
        Next
    Next
    For c = 1 To 4
        ws.Cells(12, c).Value = DateSerial(2026, c, 1)
    Next
    ' "rows+reload": the layout saved and opened again before the edit.
    Dim reload As Boolean, first As String
    reload = Right(kind, 7) = "+reload"
    If reload Then kind = Left(kind, Len(kind) - 7)
    Select Case kind
        Case "rows"
            ws.Range("F1:F4").SparklineGroups.Add xlSparkLine, "A1:D4"
        Case "cols"
            ws.Range("A10:D10").SparklineGroups.Add xlSparkColumn, "A1:D8"
        Case "dates"
            Set g = ws.Range("F1:F4").SparklineGroups.Add(xlSparkLine, "A1:D4")
            g.DateRange = "A12:D12"
        Case "other"
            ws.Range("F1:F4").SparklineGroups.Add xlSparkLine, "T!A1:D4"
        Case "two"
            ws.Range("F1:F4").SparklineGroups.Add xlSparkLine, "A1:D4"
            ws.Range("A10:D10").SparklineGroups.Add xlSparkColumn, "A1:D8"
        Case "abs"
            ws.Range("F6").SparklineGroups.Add xlSparkLine, "$A$6:$D$6"
        Case "under"
            ws.Range("A10:D10").SparklineGroups.Add xlSparkColumn, "A1:D8"
            ws.Range("F1:F4").SparklineGroups.Add xlSparkLine, "A1:D4"
        Case "gapped"
            Set g = ws.Range("F1:F4").SparklineGroups.Add(xlSparkLine, "A1:D4")
            g.Item(2).SourceData = "A6:D6"
    End Select
    If reload Then
        first = Environ("TEMP") & "\pyovba_sparkline_shift_first.xlsx"
        On Error Resume Next
        Kill first
        On Error GoTo 0
        Application.DisplayAlerts = False
        wb.SaveAs first, 51
        Application.DisplayAlerts = True
        wb.Close False
        Set wb = Workbooks.Open(first)
        Set ws = wb.Worksheets(1)
        Set other = wb.Worksheets(2)
    End If
    ws.Activate
    Dim target As String, how As String, parts() As String, sheet As Object
    parts = Split(edit, " ")
    target = parts(1)
    how = parts(2)
    Set sheet = ws
    If Left(target, 2) = "T!" Then
        Set sheet = other
        target = Mid(target, 3)
    End If
    Set wt = sheet
    On Error Resume Next
    Err.Clear
    If parts(0) = "none" Then
    ElseIf how = "auto" And typed Then
        If parts(0) = "insert" Then wt.Range(target).Insert Else wt.Range(target).Delete
    ElseIf how = "auto" Then
        If parts(0) = "insert" Then sheet.Range(target).Insert Else sheet.Range(target).Delete
    ElseIf typed Then
        If parts(0) = "insert" Then
            wt.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            wt.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    Else
        If parts(0) = "insert" Then
            sheet.Range(target).Insert Shift:=IIf(how = "down", xlShiftDown, xlShiftToRight)
        Else
            sheet.Range(target).Delete Shift:=IIf(how = "up", xlShiftUp, xlShiftToLeft)
        End If
    End If
    out = "err=" & Err.Number & ":" & Err.Description & " " & Dump(wb)
    Err.Clear
    path = Environ("TEMP") & "\pyovba_sparkline_shift.xlsx"
    Kill path
    Err.Clear
    Application.DisplayAlerts = False
    wb.SaveAs path, 51
    Application.DisplayAlerts = True
    out = out & " save=" & Err.Number
    wb.Close False
    Set wb = Workbooks.Open(path)
    out = out & " ~ " & Dump(wb)
    wb.Close False
    If reload Then Kill first
    Go = out
End Function
'''
KINDS = ('rows', 'cols', 'dates', 'other', 'two', 'abs')
EDITS = ('insert 2:2 auto', 'delete 2:2 auto', 'insert B:B auto', 'delete B:B auto', 'delete 1:4 auto',
         'delete F:F auto', 'insert F2 down', 'delete F2 up', 'insert B2 right', 'delete B2:B4 left',
         'delete A1:D1 up', 'insert A1:D8 right', 'delete A:D auto', 'insert 1:1 auto', 'delete F1:F4 left',
         'insert E1:E4 right', 'delete 9:20 auto', 'insert A2:D2 down', 'delete 12:12 auto', 'delete A10 up',
         'delete A10:D10 up', 'insert A10 right', 'delete 1:8 auto')
EDITS_OTHER = ('insert T!2:2 auto', 'delete T!2:2 auto', 'delete T!B:B auto', 'insert T!A2:D2 down',
               'delete T!A:D auto', 'delete T!1:4 auto')
EDITS_2 = {
    'rows': ('none - -', 'insert 5:5 auto', 'insert F5 down', 'insert F1 down', 'insert F3 right', 'insert 3:4 auto',
             'delete 2:3 auto', 'delete A2:D3 up', 'delete F2:F3 up', 'insert F2:F3 down', 'insert G:G auto',
             'insert A:A auto', 'delete A:A auto', 'delete A1:A4 left', 'insert E2 right'),
    'cols': ('none - -', 'insert E:E auto', 'insert A:A auto', 'insert B10 right', 'insert B10:C10 right',
             'insert A10 down', 'delete B1:B8 left', 'delete A1:A4 up', 'insert 11:11 auto', 'insert E10 right'),
    'dates': ('none - -', 'delete C:C auto', 'delete A12 left', 'insert C12 right', 'delete A12:D12 up',
              'delete A:A auto', 'delete B12:C12 left', 'insert 12:12 auto'),
    'gapped': ('none - -', 'insert F2 down', 'insert F3 down', 'delete F1 up', 'insert 5:5 auto', 'delete 6:6 auto',
               'insert 2:2 auto', 'insert G:G auto'),
    'other': ('none - -', 'insert T!B:B auto', 'delete T!A1:D1 up', 'insert T!1:1 auto', 'insert 5:5 auto',
              'insert F5 down'),
    'abs': ('none - -', 'insert F7 down', 'insert 7:7 auto', 'insert G:G auto', 'insert G6 right', 'delete A6:D6 up'),
    'two': ('none - -', 'delete A:A auto', 'insert 5:5 auto'),
}
CASES_2 = [[kind, edit, typed] for kind, edits in EDITS_2.items() for edit in edits for typed in (True, False)]
CASES = [[kind, edit, typed] for kind in KINDS for edit in EDITS + (EDITS_OTHER if kind == 'other' else ())
         for typed in (True, False)]


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig
    from pyvbaharness.results import HarnessError

    stem = 'sparkline_shifts' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    looks = {'_a': [case for case in CASES if case[0] in KINDS[:3]], '_b': [case for case in CASES if case[0] in KINDS[3:]],
             '_c': CASES_2,
             '_d': [[kind, edit, typed] for kind in ('under', 'two') for edit in ('none - -', 'insert 2:2 auto')
                    for typed in (True, False)],
             '_e': [[kind + '+reload', edit, typed] for kind, edit in (
                 ('two', 'none - -'), ('under', 'none - -'), ('two', 'insert 2:2 auto'), ('rows', 'delete A1:D1 up'),
                 ('dates', 'delete 12:12 auto'), ('gapped', 'insert F3 down'), ('other', 'delete T!2:2 auto'),
                 ('abs', 'insert G:G auto')) for typed in (True, False)],
             '_f': [[kind, f'{verb} {target} auto', typed] for kind in ('rows', 'cols', 'plain') for verb, targets in (
                 ('insert', ('C9:C10', 'C10:C11', 'C8:C9', 'C11:C12', 'A10:A11', 'E9:E10', 'F2:F3', 'F3:F5', 'F4:F5',
                             'F5:F6', 'G2:G3', 'E4:F4', 'E5:F5', 'F4:G4', 'G4:H4', 'E1:F1', 'A10:B10', 'A11:B11',
                             'A9:B9', 'C10', 'F4', 'B10:C10', 'F1:F4', 'A10:D10')),
                 ('delete', ('C9:C10', 'D9:D10', 'E9:E10', 'F2:F3', 'F4:F5', 'G2:G3', 'E4:F4', 'A11:B11', 'A9:B9',
                             'G1:H1', 'C10', 'F4', 'A10:B10', 'F3:F4'))) for target in targets
                    for typed in (True, False)]}
    cases: list[dict[str, Any]] = (json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists()
                                   else [{'call': one} for one in looks.get(stem[len('sparkline_shifts'):], CASES)])
    answers: list[str] = []
    files: list[list[str]] = []
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
                    files.append(saved())
                    number += 1
        except HarnessError as error:
            crashes.append([number, cases[number]['call'], str(error)[:200]])
            print('crash at case', number, cases[number]['call'], file=sys.stderr, flush=True)
            if len(crashes) > 1 and crashes[-2][0] == number:
                answers.append('CRASH twice')
                files.append([])
                number += 1
    (FIXTURES / f'{stem}.json').write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(),
                                                       'source': SOURCE, 'cases': [case['call'] for case in cases],
                                                       'answers': answers, 'files': files, 'crashes': crashes}, indent=0),
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


def saved() -> list[str]:
    """The sparkline groups each sheet of the saved book holds, as its part has them, and the book gone."""
    import os
    import re
    import zipfile

    path = Path(os.environ['TEMP']) / 'pyovba_sparkline_shift.xlsx'
    if not path.exists():
        return []
    with zipfile.ZipFile(path) as book:
        parts = [book.read(f'xl/worksheets/sheet{number}.xml').decode('utf-8') for number in (1, 2)]
    path.unlink()
    return [''.join(re.findall(r'<x14:sparklineGroups.*?</x14:sparklineGroups>', part, re.S)) for part in parts]


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


HELD_KINDS = ('rows', 'cols', 'dates', 'other', 'two', 'abs', 'under', 'gapped')


def held_out(round_: str) -> list[list[object]]:
    """Fresh edits, chosen at random for the round: rows and columns, and blocks of cells with each Shift and none,
    over and beside every layout, half of them on a reopened book, typed and late."""
    import random

    pick = random.Random('sparkline_shifts' + round_)
    letters = 'ABCDEFGHIJKL'
    cases: list[list[object]] = []
    while len(cases) < 200:
        kind = pick.choice(HELD_KINDS) + pick.choice(('', '+reload'))
        verb = pick.choice(('insert', 'delete'))
        shape = pick.choice(('rows', 'columns', 'cells', 'cells'))
        if shape == 'rows':
            top = pick.randint(1, 14)
            target, how = f'{top}:{top + pick.randint(0, 2)}', 'auto'
        elif shape == 'columns':
            left = pick.randint(0, 8)
            target, how = f'{letters[left]}:{letters[left + pick.randint(0, 1)]}', 'auto'
        else:
            top, left = pick.randint(1, 13), pick.randint(0, 7)
            bottom, right = top + pick.choice((0, 0, 1, 3)), left + pick.choice((0, 0, 1, 3))
            target = f'{letters[left]}{top}' + ('' if (bottom, right) == (top, left) else f':{letters[right]}{bottom}')
            how = pick.choice(('auto', 'down', 'right') if verb == 'insert' else ('auto', 'up', 'left'))
        if kind.startswith('other') and pick.random() < 0.5:
            target = 'T!' + target
        edit = f'{verb} {target} {how}'
        cases += [[kind, edit, True], [kind, edit, False]]
    return cases


def predict(round_: str) -> None:
    found = [{'call': case, 'model': model_answer(case), 'files': saved()} for case in held_out(round_)]
    (FIXTURES / f'sparkline_shifts{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                         encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
