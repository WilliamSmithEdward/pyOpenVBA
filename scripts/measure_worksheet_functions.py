"""Worksheet functions called from VBA, as Excel answers them.

Every case calls one function twice, through WorksheetFunction, where an
error is raised, and through Application, where it comes back as a value,
and records the type and value of each answer -- an array with its bounds
and items -- or the error. The arguments are numbers, text, Booleans,
Dates, errors made with CVErr, VBA arrays of one and two dimensions, and
ranges of a small sheet: 1 to 5 in A1:A5, a to e in B1:B5, a mix of a
number, text, a blank, a Boolean and an error in C1:C5, a block of six
numbers in D1:E3, 3, 1, 2 in F1:F3, and a table laid across in G1:J2.

    python scripts/measure_worksheet_functions.py

writes tests/fixtures/worksheet_functions.json, which
tests/test_excel_worksheet_functions.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "worksheet_functions.json"

HELPER = '''Private Function Item(v As Variant) As String
    If IsError(v) Then
        Item = "Error:" & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        Item = "Empty"
    ElseIf IsObject(v) Then
        Item = "Object"
    Else
        Item = TypeName(v) & ":" & CStr(v)
    End If
End Function

Private Function Show(v As Variant) As String
    Dim out As String, r As Long, c As Long, rows As Long, cols As Long, two As Boolean
    If IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = TypeName(v) & "@" & v.Address(False, False)
        Exit Function
    End If
    If Not IsArray(v) Then
        Show = Item(v)
        Exit Function
    End If
    On Error Resume Next
    cols = UBound(v, 2)
    two = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If two Then
        out = "Array(" & LBound(v, 1) & " To " & UBound(v, 1) & ", " & LBound(v, 2) & " To " & UBound(v, 2) & ")"
        For r = LBound(v, 1) To UBound(v, 1)
            For c = LBound(v, 2) To UBound(v, 2)
                out = out & " " & Item(v(r, c))
            Next
        Next
    Else
        out = "Array(" & LBound(v) & " To " & UBound(v) & ")"
        For r = LBound(v) To UBound(v)
            out = out & " " & Item(v(r))
        Next
    End If
    Show = out
End Function
'''

SETUP = '''Private Sub Fill(ws As Object)
    Dim r As Long
    For r = 1 To 5
        ws.Cells(r, 1).Value = r
        ws.Cells(r, 2).Value = Chr(96 + r)
    Next
    ws.Range("C1").Value = 10
    ws.Range("C2").Value = "x"
    ws.Range("C4").Value = True
    ws.Range("C5").Value = CVErr(xlErrNA)
    ws.Range("D1:E1").Value = Array(1, 2)
    ws.Range("D2:E2").Value = Array(3, 4)
    ws.Range("D3:E3").Value = Array(5, 6)
    ws.Range("F1").Value = 3
    ws.Range("F2").Value = 1
    ws.Range("F3").Value = 2
    ws.Range("G1:J1").Value = Array("p", "q", "r", "s")
    ws.Range("G2:J2").Value = Array(1, 2, 3, 4)
End Sub
'''

A, B, C, D = 'ws.Range("A1:A5")', 'ws.Range("B1:B5")', 'ws.Range("C1:C5")', 'ws.Range("D1:E3")'

#: (function, arguments as VBA); each is called through WorksheetFunction and through Application.
CASES: list[tuple[str, str]] = [
    ("And", "True, 1"), ("And", A), ("And", 'False, "x"'),
    ("Average", '1, 2, "3"'), ("Average", C), ("Average", f"{A}, 10"), ("Average", "Array(1, 2, 3)"),
    ("AverageIf", f'{A}, ">2"'), ("AverageIf", f'{B}, "b", {A}'),
    ("Ceiling", "2.5, 1"), ("Ceiling", "-2.5, -2"), ("Ceiling", "2.5, -1"),
    ("Ceiling_Math", "2.5"), ("Ceiling_Math", "-2.5, 2, 1"),
    ("Choose", '2, "a", "b", "c"'), ("Choose", '4, "a", "b"'),
    ("Concat", '"a", 1, True'), ("Concat", 'ws.Range("B1:B3")'),
    ("Count", '1, "2", "x", True'), ("Count", C), ("Count", 'Array(1, "a", True)'),
    ("CountA", C), ("CountA", '1, "", "x"'), ("CountBlank", C),
    ("CountIf", f'{A}, ">2"'), ("CountIf", f'{B}, "b"'), ("CountIf", f'{B}, "*"'), ("CountIf", f"{A}, 3"),
    ("CountIfs", f'{A}, ">1", ws.Range("F1:F5"), "<3"'),
    ("Days", "#1/10/2020#, #1/1/2020#"), ("Days", '"2020-01-10", "2020-01-01"'),
    ("DevSq", "Array(1, 2, 3, 4)"), ("DevSq", C), ("DevSq", f"{A}, 10"), ("DevSq", '1, "2", True'), ("DevSq", B),
    ("EDate", "#1/31/2020#, 1"), ("EoMonth", "#1/15/2020#, 0"),
    ("Find", '"b", "abc"'), ("Find", '"z", "abc"'), ("Find", '"B", "abc"'),
    ("Floor", "2.5, 1"), ("Floor_Math", "-2.5"),
    ("HLookup", '"q", ws.Range("G1:J2"), 2, False'), ("HLookup", '"z", ws.Range("G1:J2"), 2, False'),
    ("IfError", '5, 0'), ("IfError", 'CVErr(xlErrDiv0), "x"'), ("IfNa", 'CVErr(xlErrNA), "na"'),
    ("Index", f"{D}, 2, 2"), ("Index", f"{A}, 3"), ("Index", "Array(1, 2, 3), 2"), ("Index", f"{D}, 0, 1"),
    ("Index", f"{D}, 2, 0"), ("Index", "Array(Array(1, 2), Array(3, 4)), 2, 1"),
    ("IsErr", "CVErr(xlErrNA)"), ("IsError", "CVErr(xlErrNA)"), ("IsNA", 'ws.Range("C5")'),
    ("IsLogical", "True"), ("IsNonText", '"x"'), ("IsNumber", '"1"'), ("IsText", '"x"'),
    ("Large", f"{A}, 2"), ("Small", "Array(5, 1, 3), 1"),
    ("Ln", "10"), ("Log", "100"), ("Log", "8, 2"), ("Log10", "1000"),
    ("Match", f'"c", {B}, 0'), ("Match", f"3.5, {A}, 1"), ("Match", f'"z", {B}, 0'), ("Match", 'Array(3, 1), 3, 0'),
    ("Match", f'"C", {B}, 0'),
    ("Max", '1, "5", ws.Range("C1:C4")'), ("Max", "Array(1, 7, 3)"), ("Min", C), ("Min", "-1, 2"),
    ("Median", "Array(3, 1, 2)"), ("MRound", "10, 3"),
    ("Or", "False, 0"), ("Pi", ""), ("Power", "2, 10"), ("Product", A),
    ("Proper", '"hello WORLD o\'neil 2nd"'),
    ("Replace", '"abcdef", 2, 3, "X"'), ("Rept", '"ab", 3'),
    ("Round", "2.345, 2"), ("Round", "2.5, 0"), ("Round", "-2.5, 0"), ("RoundDown", "2.345, 1"), ("RoundUp", "2.341, 1"),
    ("Search", '"C", "abc"'), ("Search", '"?c", "abc"'),
    ("StDev", "Array(1, 2, 3, 4)"), ("StDev_P", "Array(1, 2, 3, 4)"), ("StDev_S", "Array(1, 2, 3, 4)"),
    ("StDevP", "Array(1, 2, 3, 4)"), ("Var", "Array(1, 2, 3, 4)"), ("Var_P", "Array(1, 2, 3, 4)"),
    ("Var_S", "Array(1, 2, 3, 4)"), ("VarP", "Array(1, 2, 3, 4)"),
    ("Substitute", '"a-b-c", "-", "+"'), ("Substitute", '"a-b-c", "-", "+", 2'),
    ("Sum", '1, "2", True'), ("Sum", 'ws.Range("C1:C4")'), ("Sum", C), ("Sum", 'Array(1, "2", True)'),
    ("Sum", "#1/2/2020#, 1"),
    ("SumIf", f'{A}, ">2"'), ("SumIf", f'{B}, "b", {A}'),
    ("SumIfs", f'{A}, {A}, ">1", {A}, "<5"'),
    ("SumProduct", 'ws.Range("A1:A3"), ws.Range("F1:F3")'), ("SumProduct", "Array(1, 2), Array(3, 4)"),
    ("SumSq", "Array(1, 2, 3, 4)"), ("SumSq", C), ("SumSq", f"{A}, 10"), ("SumSq", '1, "2", True'), ("SumSq", B),
    ("Text", '1234.5, "#,##0.00"'), ("TextJoin", '",", True, ws.Range("B1:B3")'),
    ("TextJoin", '",", False, "a", "", "b"'),
    ("Transpose", "Array(1, 2, 3)"), ("Transpose", 'ws.Range("A1:A3")'), ("Transpose", D),
    ("Transpose", 'ws.Range("G1:J1")'), ("Transpose", "Array(Array(1, 2), Array(3, 4))"),
    ("Trim", '"  a   b  "'),
    ("VLookup", 'ws.Range("A3"), ws.Range("A1:B5"), 2, False'), ("VLookup", '3, ws.Range("A1:B5"), 2, False'),
    ("VLookup", '3.5, ws.Range("A1:B5"), 2, True'), ("VLookup", '9, ws.Range("A1:B5"), 2, False'),
    ("VLookup", '"C", ws.Range("B1:B5"), 1, False'), ("VLookup", '3, ws.Range("A1:B5"), 2'),
    ("Weekday", "#1/5/2020#"), ("Weekday", "#1/5/2020#, 2"),
    ("XLookup", f'"c", {B}, {A}'), ("XLookup", f'"z", {B}, {A}, "none"'),
    ("Xor", "True, True, False"),
    # Blanks found, Empty passed, and ranges where a number is wanted.
    ("VLookup", '3, ws.Range("A1:C5"), 3, False'), ("Index", 'ws.Range("C1:C5"), 3'), ("Choose", "1, Empty"),
    ("Sum", "Empty, 1"), ("Count", "Empty"), ("CountA", "Empty, 1"), ("Concat", '"a", Empty, "b"'),
    ("Max", "Empty, -1"), ("Round", 'ws.Range("A2"), 0'), ("Round", '"2.5", 0'), ("Round", 'ws.Range("A1:A2"), 0'),
    ("Power", "Empty, 0"), ("Rept", '"x", 2.9'), ("IsText", 'ws.Range("B1:B2")'), ("Sum", "Null"),
    ("Sum", "CCur(1.5), CDec(2.5)"), ("And", "Array(True, False)"), ("Match", '2, Array(1, 2, 3), 0'),
    ("Match", '"b", Array("a", "b"), 0'), ("Index", "Array(Array(1, 2), Array(3, 4)), 0, 2"),
    ("Transpose", "Array(Array(1), Array(2))"), ("Transpose", "5"), ("Index", 'ws.Range("A1:A5"), 0, 1'),
]


#: Cases per procedure: a VBA procedure has a size limit, and one past it stops the module compiling.
BATCH = 40


def module() -> str:
    batches: list[str] = []
    for start in range(0, len(CASES), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(ws As Object) As String", "Dim out As String, v As Variant",
                 "On Error Resume Next"]
        for name, args in CASES[start:start + BATCH]:
            call = f"({args})" if args else ""
            for target in (f"Application.WorksheetFunction.{name}{call}", f"Application.{name}{call}"):
                lines += ["Err.Clear", "v = Empty", f"v = {target}",
                          'If Err.Number <> 0 Then out = out & "E" & Err.Number & "^" Else out = out & Show(v) & "^"']
            lines.append('out = out & "|"')
        lines += ["On Error GoTo 0", f"Batch{start // BATCH} = out", "End Function"]
        batches.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "Fill ws",
             *[f"out = out & Batch{index}(ws)" for index in range(len(batches))],
             "wb.Close False", "Probe = out", "End Function"]
    return "\n".join([HELPER, SETUP, *batches, *probe]) + "\n"


def main() -> None:
    code = module()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(code, "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = [part.split("^")[:2] for part in str(result.value).split("|")[: len(CASES)]]
    record = {"helper": HELPER, "setup": SETUP, "cases": [
        {"function": name, "arguments": args, "worksheet_function": wf, "application": app}
        for (name, args), (wf, app) in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for (name, args), (wf, app) in zip(CASES, answers, strict=True):
        print(f"{name}({args}) => {wf} | {app}")


if __name__ == "__main__":
    main()
