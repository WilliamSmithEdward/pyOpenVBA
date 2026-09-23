"""What Excel's Evaluate answers, from VBA.

Each case hands an expression to Application.Evaluate, to the bracket
form [...], or to Worksheet.Evaluate on a sheet that is not active, and
records what comes back: a Range with its sheet and address, an array
with its bounds and items, an error value, or a value and its type. The
active sheet, Data, holds 1 to 5 in A1:A5, a to e in B1:B5, 0.1, 0.2 and
0.3 in C1:C3 and =A1*10 in G5; the sheet Other holds 100 and 200 in
A1:A2. The book names Block (Data!$A$1:$A$3), Rate (0.5) and Twice
(=Data!$A$1*2).

    python scripts/measure_evaluate.py

writes tests/fixtures/evaluate.json, which tests/test_excel_evaluate.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "evaluate.json"

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
    Dim out As String, r As Long, c As Long, cols As Long, two As Boolean
    If IsObject(v) Then
        If v Is Nothing Then Show = "Nothing" Else Show = TypeName(v) & "@" & v.Parent.Name & "!" & v.Address(False, False)
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

SETUP = '''Private Sub Fill(wb As Object)
    Dim ws As Object, r As Long
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    For r = 1 To 5
        ws.Cells(r, 1).Value = r
        ws.Cells(r, 2).Value = Chr(96 + r)
    Next
    ws.Range("C1").Value = 0.1
    ws.Range("C2").Value = 0.2
    ws.Range("C3").Value = 0.3
    ws.Range("G5").Formula = "=A1*10"
    wb.Worksheets.Add(After:=ws).Name = "Other"
    wb.Worksheets("Other").Range("A1").Value = 100
    wb.Worksheets("Other").Range("A2").Value = 200
    wb.Names.Add Name:="Block", RefersTo:="=Data!$A$1:$A$3"
    wb.Names.Add Name:="Rate", RefersTo:="=0.5"
    wb.Names.Add Name:="Twice", RefersTo:="=Data!$A$1*2"
    ws.Activate
End Sub
'''

#: Expressions for Application.Evaluate.
APPLICATION = [
    # Arithmetic and constants.
    "1+2", "=1+2", " 1 + 2 ", "2^10", "-A1", "A1%", "7/2", "1/0", '"abc"', "TRUE", "1E+300*1E+300",
    "0.3-0.2-0.1", "(0.3-0.2-0.1)", "A1&B1", "C1+C2=C3", "C1+C2-C3",
    # Blocks, which Evaluate works out whole.
    "A1:A3*2", "A1:A3=2", "A1:A3+B1:B3", "{1,2;3,4}*2", "{1,2,3}+{10;20}", "-A1:A2", "A1:B2&\"!\"",
    # References.
    "A1", "A1:C1", "Data!A2", "Other!A1", "$A$1", "A:A", "Block", "Rate", "Twice", "G5", "G5+1",
    "A1:A5 A3:B3", "(A1:A2,A4)", "A1:A2 B1:B2", "A1:B2 B2:C3", "Block A2:B2", "SUM(A1:A5 A2:B3)", "(A1,B1)",
    "CHOOSE(2,A1,A2)", "IF(TRUE,A1:A2)", "XLOOKUP(3,A1:A5,B1:B5)",
    # Functions.
    "SUM(A1:A5)", "SUM(A1:A3*A1:A3)", "SUMPRODUCT((A1:A5>2)*A1:A5)", "ROW(A1:A3)", "ROW()", "COLUMN()",
    "INDEX(A1:A5,3)", "OFFSET(A1,1,0)", "OFFSET(A1,0,0,3,1)", 'INDIRECT("A2")', 'IF(A1>0,"pos","neg")',
    'IF(A1:A3>1,"big","small")', "NOW()>0", "DATE(2020,1,2)", 'COUNTIF(A1:A5,">2")', "TRANSPOSE(A1:A3)",
    "UPPER(B1:B3)", "LEN(B1)", "MAX(A1:A5)-MIN(A1:A5)", "VLOOKUP(3,A1:B5,2,FALSE)", 'MATCH("c",B1:B5,0)',
    "ISNUMBER(A1)", "NA()", "Twice*A2", "SUM(Block)",
    # What Excel cannot read.
    "NOTAFUNCTION(1)", "notaname", "1+", "", "SUM(", "+".join(["1"] * 140),
]
#: Expressions for the bracket form, written into the code as they are.
BRACKETS = ["1+2", "A1+A2", "A1:A3*2", "SUM(A1:A5)", "A1", "Other!A1", "Twice", "0.3-0.2-0.1", "Block"]
#: Expressions for Other.Evaluate while Data is active.
SHEET = ["A1", "A1*2", "SUM(Data!A1:A5)", "Data!A1+A1", "ROW()", "Twice", "Block"]


def cases() -> list[tuple[str, str]]:
    return ([("Application", one) for one in APPLICATION] + [("Brackets", one) for one in BRACKETS]
            + [("Other", one) for one in SHEET])


def call(how: str, expression: str) -> str:
    quoted = expression.replace('"', '""')
    if how == "Application":
        return f'Application.Evaluate("{quoted}")'
    if how == "Other":
        return f'wb.Worksheets("Other").Evaluate("{quoted}")'
    return f"[{expression}]"


#: Cases per procedure: a VBA procedure has a size limit.
BATCH = 30


def module() -> str:
    batches: list[str] = []
    every = cases()
    for start in range(0, len(every), BATCH):
        lines = [f"Private Function Batch{start // BATCH}(wb As Object) As String", "Dim out As String",
                 "On Error Resume Next"]
        for how, expression in every[start:start + BATCH]:
            lines += ["Err.Clear", f'out = out & Show({call(how, expression)}) & "^"',
                      'If Err.Number <> 0 Then out = out & "E" & Err.Number & "^"', 'out = out & "|"']
        lines += ["On Error GoTo 0", f"Batch{start // BATCH} = out", "End Function"]
        batches.append("\n".join(lines))
    probe = ["Public Function Probe() As String", "Dim wb As Object, out As String",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)", "Fill wb",
             *[f"out = out & Batch{index}(wb)" for index in range(len(batches))],
             "wb.Close False", "Probe = out", "End Function"]
    return "\n".join([HELPER, SETUP, *batches, *probe]) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    every = cases()
    answers = [part.split("^")[0] for part in str(result.value).split("|")[: len(every)]]
    record = {"helper": HELPER, "setup": SETUP, "cases": [
        {"how": how, "expression": expression, "answer": answer}
        for (how, expression), answer in zip(every, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for (how, expression), answer in zip(every, answers, strict=True):
        print(f"{how:11} {expression[:40]:42} {answer}")


if __name__ == "__main__":
    main()
