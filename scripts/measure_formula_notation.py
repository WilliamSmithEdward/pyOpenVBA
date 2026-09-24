"""What Excel makes of a formula written in the other notation.

Range.Formula reads a formula in A1, and where A1 cannot read it, in
R1C1: =R[-1]C+1 written to C8 is =C7+1. A name that looks like an R1C1
reference, R1C1, RC, R or C, is one A1 cannot read. FormulaR1C1 reads
R1C1 alone, and takes what looks like an A1 cell for a name. Each
formula is written to C8 of a sheet called Data with 1 in A1:C10, through
Formula, Formula2, Value, FormulaArray or FormulaR1C1, and its Formula,
FormulaR1C1 and value read back; a few are written to a block.

    python scripts/measure_formula_notation.py

writes tests/fixtures/formula_notation.json, which
tests/test_excel_formula_notation.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "formula_notation.json"

SETUP = 'ws.Name = "Data"\nws.Range("A1:C10").Value = 1'
#: Each case: the member written through, the formula, and the cells it is written to.
CASES: list[tuple[str, str, str]] = [
    ("Formula", "=R[-1]C+1", "C8"), ("Formula", "=R1C1", "C8"), ("Formula", "=RC[-1]", "C8"),
    ("Formula", "=SUM(R1C1:R2C2)", "C8"), ("Formula", "=R1C1+A1", "C8"), ("Formula", "=RC", "C8"),
    ("Formula", "=R[1]C[1]*2", "C8"), ("Formula", "=rc[-1]", "C8"), ("Formula", "=R", "C8"), ("Formula", "=C", "C8"),
    ("Formula", "=R2", "C8"), ("Formula", "=C3", "C8"), ("Formula", "=R2C", "C8"), ("Formula", "=RC2", "C8"),
    ("Formula", "=SUM(R:R)", "C8"), ("Formula", "=SUM(C:C)", "C8"), ("Formula", "=R1C1:R2C2", "C8"),
    ("Formula", "=INDEX(R1C1:R3C1,2)", "C8"), ("Formula", "=R[-1]C+A1", "C8"), ("Formula", "=LET(x,R1C1,x)", "C8"),
    ("Formula", "=Data!R1C1", "C8"), ("Formula", "=SUM(R1C1,R2C2)", "C8"), ("Formula", '=R1C1&"R1C1"', "C8"),
    ("Formula", "=R[-1]C:R[1]C", "C8"), ("Formula", "=SUM(R1:R2)", "C8"), ("Formula", "=SUM(C1:C2)", "C8"),
    ("Formula", "=R[-7]C[-2]", "C8"), ("Formula", "=r1c1", "C8"), ("Formula", "=R1c1", "C8"),
    ("Formula", "=R0C1", "C8"), ("Formula", "=R1C0", "C8"), ("Formula", "=R1048577C1", "C8"),
    ("Formula", "=R[-8]C", "C8"), ("Formula", "=RC[-3]", "C8"), ("Formula", "=R1C1 R1C2", "C8"),
    ("Formula", "=SUM(R[-7]C[-2]:R[-6]C[-1])", "C8"), ("Formula", "=R[-1]C", "D2:D4"), ("Formula", "=R1C1", "D2:D4"),
    ("Formula", "=R[-1]C", "F2,F5"), ("Formula2", "=R[-1]C+1", "C8"), ("Value", "=R[-1]C+1", "C8"),
    ("Value", "=R1C1", "C8"), ("FormulaArray", "=R[-1]C+1", "C8"), ("FormulaArray", "=R1C1", "C8"),
    ("FormulaR1C1", "=A1", "C8"), ("FormulaR1C1", "=A1+1", "C8"), ("FormulaR1C1", "=SUM(A1:B2)", "C8"),
    ("FormulaR1C1", "=B7+R1C1", "C8"), ("FormulaR1C1", "=$A$1", "C8"), ("FormulaR1C1", "=Data!A1", "C8"),
    ("FormulaR1C1", "=R1C1", "C8"), ("FormulaR1C1", "=XFD1", "C8"), ("FormulaR1C1", "=XFE1", "C8"),
    ("FormulaR1C1", "=A1048577", "C8"), ("FormulaR1C1", "=a1", "C8"), ("FormulaR1C1", "=Tax1", "C8"),
    # A name that starts with a cell: A1 reads the cell and cannot read what follows.
    ("Formula", "=A1B", "C8"), ("Formula", "=Q1Sales", "C8"), ("Formula", "=X1Y", "C8"), ("Formula", "=R1Foo", "C8"),
    ("Formula", "=AB12CD", "C8"), ("Formula", "=Tax1a", "C8"), ("Formula", "=A1_", "C8"), ("Formula", "=A1.B", "C8"),
    ("Formula", "=R1C", "C8"), ("Formula", "=C1R1", "C8"), ("Formula", "=RC1", "C8"), ("Formula", "=CR1", "C8"),
    ("Formula", "=RR", "C8"), ("Formula", "=CC", "C8"), ("Formula", "=r", "C8"), ("Formula", "=R_1", "C8"),
    ("Formula", "=R1C1C1", "C8"), ("Formula", "=XFD1048576x", "C8"), ("Formula", "=XFE1x", "C8"),
    ("Formula", "=A0x", "C8"), ("FormulaR1C1", "=Q1Sales", "C8"), ("FormulaR1C1", "=A1B", "C8"),
    ("FormulaR1C1", "=RC[-1]+A1", "C8"),
]


def _module() -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String, cell As Object",
             "Set ws = ActiveWorkbook.Worksheets.Add", SETUP, "On Error Resume Next"]
    for member, formula, target in CASES:
        quoted = formula.replace('"', '""')
        first = target.split(":")[0].split(",")[0]
        lines += ["Err.Clear", 'ws.Range("C8:F10").ClearContents', f'ws.Range("{target}").{member} = "{quoted}"',
                  'If Err.Number <> 0 Then',
                  '    out = out & "!" & Err.Number',
                  "Else",
                  '    out = out & "ok"',
                  f'    For Each cell In ws.Range("{target}").Cells',
                  '        out = out & "~" & cell.Formula & "~" & cell.FormulaR1C1 & "~" & cell.Text',
                  "    Next",
                  "End If",
                  f"If Err.Number <> 0 Then out = out & \"~read!\" & Err.Number  ' {first}",
                  'out = out & "|"']
    lines += ["On Error GoTo 0", "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True",
              "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(_module(), "Probe", timeout=300.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    answers = str(result.value).split("|")[:-1]
    assert len(answers) == len(CASES), (len(answers), len(CASES))
    cases = [{"member": member, "formula": formula, "target": target, "answer": answer}
             for (member, formula, target), answer in zip(CASES, answers, strict=True)]
    OUT.write_text(json.dumps({"setup": SETUP, "cases": cases}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
