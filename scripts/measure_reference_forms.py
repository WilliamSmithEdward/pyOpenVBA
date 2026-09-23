"""Formulas that work with references rather than values, as Excel works them.

The range operator between references (A1:INDEX(A:A,5)), the
intersection (A1:A5 A2:B3) and the union ((A1:A2,A4)), and the functions
that come to cells -- INDEX, OFFSET, INDIRECT, CHOOSE, IF -- used where a
reference is wanted. Each formula is written to a cell of column E with
Range.Formula; what the cell holds and how Excel spells the formula back
are recorded. The sheet holds 1 to 10 in A1:A10, 11 to 20 in B1:B10 and
x, y, z in C1:C3, and the book names Block ($A$1:$A$3), Dyn
(OFFSET($A$1,0,0,3,1)) and Pick (INDEX($A$1:$A$10,4)).

    python scripts/measure_reference_forms.py

writes tests/fixtures/reference_forms.json, which tests/test_formula_reference_forms.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "reference_forms.json"

SETUP = '''Private Sub Fill(ws As Object)
    Dim r As Long
    For r = 1 To 10
        ws.Cells(r, 1).Value = r
        ws.Cells(r, 2).Value = r + 10
    Next
    ws.Range("C1").Value = "x"
    ws.Range("C2").Value = "y"
    ws.Range("C3").Value = "z"
    ws.Parent.Names.Add Name:="Block", RefersTo:="=" & ws.Name & "!$A$1:$A$3"
    ws.Parent.Names.Add Name:="Dyn", RefersTo:="=OFFSET(" & ws.Name & "!$A$1,0,0,3,1)"
    ws.Parent.Names.Add Name:="Pick", RefersTo:="=INDEX(" & ws.Name & "!$A$1:$A$10,4)"
End Sub
'''

FORMULAS = [
    # The range operator with a function on either side.
    "=SUM(A1:INDEX(A1:A10,3))", "=SUM(INDEX(A1:A10,2):A5)", "=SUM(INDEX(A1:A10,2):INDEX(A1:A10,4))",
    "=SUM(A1:OFFSET(A1,2,0))", '=SUM(INDIRECT("A1"):A3)', "=SUM(Block:A5)", "=SUM(Pick:A6)", "=SUM(A1:A3:B5)",
    "=ROWS(A1:INDEX(A:A,4))", "=COLUMNS(A1:INDEX(A1:C10,1,3))", '=SUMIF(A1:INDEX(A1:A10,5),">2")',
    "=SUMPRODUCT(A1:INDEX(A1:A10,3))",
    # Intersections.
    "=SUM(A1:A5 A2:B3)", "=A1:A5 A3:B3", "=A1:A2 B3:B4", "=A1:A3 A2", "=SUM(A1:A10 A5:A6)", "=SUM(A:A A3)",
    # Unions.
    "=SUM((A1:A2,A4))", "=COUNT((A1:A2,C1))", "=COUNTA((A1:A2,C1))", "=AVERAGE((A1:A2,A4))",
    "=MAX((A1:A2,B1:B2))", "=MIN((A3,B1:B2))", "=(A1:A2,A4)", "=INDEX((A1:A2,B1:B2),1,1,2)",
    "=AREAS((A1:A2,A4))", "=SUBTOTAL(9,(A1:A2,A4))", '=COUNTIF((A1:A2,A4),">1")', "=LARGE((A1:A2,A4),1)",
    # Functions that come to cells.
    "=SUM(OFFSET(A1,1,0,3,1))", '=SUM(INDIRECT("A1:A3"))', '=SUM(INDIRECT("R1C1:R3C1",FALSE))',
    "=SUM(IF(TRUE,A1:A3))", "=SUM(CHOOSE(2,A1:A2,B1:B3))", "=SUM(Dyn)", "=ROW(INDEX(A1:A10,4))",
    "=SUM(OFFSET(Block,1,0))", "=SUM(INDEX(A1:B10,0,2))", "=SUM(INDEX(A1:B10,2,0))", "=Pick", "=Pick*2",
    '=SUM(INDIRECT("Block"))', '=SUM(INDIRECT("Sheet1!A1:A2"))', '=SUM(INDIRECT("Nowhere!A1"))',
]

PROBE_TOP = '''Private Function Item(v As Variant) As String
    If IsError(v) Then
        Item = "Error:" & CStr(CLng(v))
    ElseIf IsEmpty(v) Then
        Item = "Empty"
    Else
        Item = TypeName(v) & ":" & CStr(v)
    End If
End Function
'''


def module() -> str:
    lines = ["Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String, r As Long",
             "Application.DisplayAlerts = False", "Set wb = Workbooks.Add(xlWBATWorksheet)",
             "Set ws = wb.Worksheets(1)", "Fill ws"]
    for row, formula in enumerate(FORMULAS, 1):
        lines.append(f'ws.Range("E{row}").Formula = "{formula.replace(chr(34), chr(34) * 2)}"')
    lines += [f"For r = 1 To {len(FORMULAS)}",
              'out = out & ws.Cells(r, 5).Formula & "^" & Item(ws.Cells(r, 5).Value) & "|"', "Next r",
              "wb.Close False", "Probe = out", "End Function"]
    return "\n".join([PROBE_TOP, SETUP, *lines]) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), "Probe", timeout=600.0)
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    parts = [part.split("^") for part in str(result.value).split("|")[: len(FORMULAS)]]
    record = {"setup": SETUP, "cases": [{"formula": formula, "spelled": spelled, "value": value}
                                        for formula, (spelled, value) in zip(FORMULAS, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for formula, (spelled, value) in zip(FORMULAS, parts, strict=True):
        print(f"{formula:42} {spelled if spelled != formula else '':30} {value}")


if __name__ == "__main__":
    main()
