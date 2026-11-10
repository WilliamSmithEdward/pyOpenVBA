"""Application._Run2, the hidden twin of Application.Run: whether it runs a macro, what it answers, and its errors.

The rivals, written before measuring: (R1) it runs a macro as Run does, its answers and errors the same, or (R2) it raises.
Each case runs in the probe's own project, with Twice and Hello beside it, and compares _Run2 with Run on the same call.

    python scripts/measure_run2.py

writes tests/fixtures/run2.json, which tests/test_excel_run2.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/run2.json'
#: Each case: the arguments after the method name, the same for Run and _Run2.
CALLS: list[str] = ['("Twice", 4)', '("Twice", "4")', '("Hello")', '("Twice")', '("Nope")', '()', '("Twice", 1, 2)', '("Sum3", 1, 2, 3)', '("P.Twice", 5)',
                    '(1)', '("Twice", Empty)', '("Sum3", 1, , 3)']
# Round 2. Run gave 1004 for Sum3, a name that is also a cell's: (A1) a macro named as a cell is not run, or (A2) it was its
# Optional arguments. And _Run2 answered 8 as a Double where Run gave an Integer: (D1) every number becomes a Double, the rest
# as they are, or (D2) only Integer does. An error the macro raises is not here: Run shows Excel's own run-time error dialog,
# which no handler of the caller's reaches, and the harness ends it.
CALLS += ['("Plus3", 1, 2, 3)', '("Plus3", 1, , 3)', '("AB12")', '("Kind", 1)', '("Kind", 2)', '("Kind", 3)', '("Kind", 4)', '("Kind", 5)',
          '("Kind", 6)', '("Kind", 7)', '("Kind", 8)', '("Kind", 9)', '("Twice", 2.5)', '("Twice", True)']
# Held out: the rules from the calls above (Run and _Run2 alike refuse a name that is a cell's; _Run2 answers numbers and
# dates as Doubles, an array from 1, an object as #VALUE!, and #VALUE! for no macro or the wrong arguments), on calls written
# after them.
CALLS += ['("XFE1")', '("P.Plus3", 1)', '(2.5)', '("Hello", 1)', '("Twice", 1.5)', '("Kind", 10)', '("Kind", 11)',
          '("Kind", 12)', '("Kind", 13)', '("Kind", 14)', '("A1")']
# Held out again: values crossing as a worksheet function's do (numbers and Empty in as Doubles; an Integer or Byte out as a
# Long and every other number or date as a Double; a flat array from 1, an empty or nested one and an object as #VALUE!).
CALLS += ['("Kind", 15)', '("Kind", 16)', '("Kind", 17)']
# How each argument arrives, read by Shown, which only names it: a date given to Twice stopped on a type mismatch inside the
# macro, which Run shows in Excel's own dialog. The rivals for a date: (G1) as it was, (G2) a Double, (G3) text.
CALLS += ['("Shown", 4)', '("Shown", CByte(2))', '("Shown", CLng(3))', '("Shown", CCur(1.5))', '("Shown", #1/2/2024#)', '("Shown", CSng(0.5))',
          '("Shown", Empty)', '("Shown", True)', '("Shown", "4")', '("Shown", 2.5)']
# Held out: a date arriving as its text, on arguments written after it.
CALLS += ['("Shown", #3:15:00 PM#)', '("Shown", CDate("2024-02-29 13:05"))', '("Shown", CDec(2.5))', '("Shown", DateSerial(2024, 12, 31))']
SOURCE = '''Public Function Twice(x As Variant) As Variant
    Twice = x * 2
End Function

Public Function Sum3(Optional a As Variant = 10, Optional b As Variant = 20, Optional c As Variant = 30) As Variant
    Sum3 = a + b + c
End Function

Public Sub Hello()
End Sub

Public Function Plus3(Optional a As Variant = 10, Optional b As Variant = 20, Optional c As Variant = 30) As Variant
    Plus3 = a + b + c
End Function

Public Function Shown(x As Variant) As String
    Shown = TypeName(x) & " " & CStr(x)
End Function

Public Function XFE1() As Variant
    XFE1 = 5
End Function

Public Function AB12() As Variant
    AB12 = 1
End Function

Public Function Kind(which As Long) As Variant
    Select Case which
    Case 1: Kind = CLng(70000)
    Case 2: Kind = "text"
    Case 3: Kind = True
    Case 4: Kind = DateSerial(2024, 1, 2)
    Case 5: Kind = CCur(1.5)
    Case 6: Kind = Array(1, "a")
    Case 7: Set Kind = ThisWorkbook
    Case 8: Kind = CSng(1.5)
    Case 9: Kind = CVErr(xlErrNA)
    Case 10: Kind = CByte(7)
    Case 11: Kind = CDec(2.5)
    Case 12: Kind = Null
    Case 13: Kind = Array(Array(1, 2), 3)
    Case 14: Kind = Array()
    Case 15: Kind = CInt(3)
    Case 16: Kind = Array("a")
    Case 17: Kind = Array(1, ThisWorkbook)
    End Select
End Function

Public Function Raises() As Variant
    Err.Raise 11
End Function

Private Function Tried(ByVal how As Long, ByVal which As Long) As String
    Dim r As Variant, o As Object
    r = Empty
    On Error GoTo Bad
    Set o = Application
    Select Case which
@CASES@
    End Select
    If IsObject(r) Then
        Tried = TypeName(r)
    ElseIf IsArray(r) Then
        Tried = TypeName(r) & " " & LBound(r) & " " & UBound(r)
    Else
        Tried = TypeName(r) & " " & CStr(r)
    End If
    Exit Function
Bad:
    Tried = "E" & Err.Number
End Function

Public Function Probe() As String
    Dim s As String, which As Long, how As Long
    For which = 0 To @LAST@
        For how = 0 To 2
            s = s & "|" & Tried(how, which)
        Next
        s = s & "~"
    Next
    Probe = s
End Function
'''


def source() -> str:
    cases = []
    for index, call in enumerate(CALLS):
        cases += [f'    Case {index}',
                  f'        If how = 0 Then r = Application.Run{call}',
                  f'        If how = 1 Then r = Application.[_Run2]{call}',
                  f'        If how = 2 Then r = o.[_Run2]{call}']
    return SOURCE.replace('@CASES@', '\n'.join(cases)).replace('@LAST@', str(len(CALLS) - 1))


def main() -> None:
    code = source()
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        done = excel.run_vba(code, 'Probe', (), timeout=120.0, line_numbers=False, module_name='P')
    assert done.ok and not done.dialogs, (done.outcome, done.dialogs, done.message)
    answers = str(done.value).split('~')[:-1]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': code, 'cases': [[call, answer] for call, answer in zip(CALLS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for call, answer in record['cases']:
        print(call, '=>', answer)


if __name__ == '__main__':
    main()
