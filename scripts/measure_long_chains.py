"""What long chains of cells, and the deepest formulas Excel takes, come to in Excel.

Chains: a running total thousands of rows long and its kin, each VBA run
on a fresh worksheet, then an expression read back. Formulas: each of
the deepest formulas Excel takes -- 4096 terms, 256 brackets, 65
functions, a thousand signs, 8190 percent signs -- written to C8 with 1
in A1, and its value and the length of its Formula, FormulaR1C1 and
Formula2 read back.

    python scripts/measure_long_chains.py

writes tests/fixtures/long_chains.json, which
tests/test_excel_recalculation.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "long_chains.json"
#: The longest piece of a formula one line of VBA spells.
PIECE = 400

#: Each case: the VBA run on a fresh sheet ``ws``, and the expression then read.
CHAINS: dict[str, tuple[str, str]] = {
    "running total": ('ws.Range("A1").Value = 1\nws.Range("A2:A3000").Formula = "=A1+1"', 'ws.Range("A3000").Value'),
    "running total, the top changed": (
        'ws.Range("A1").Value = 1\nws.Range("A2:A3000").Formula = "=A1+1"\nws.Range("A1").Value = 10',
        'ws.Range("A3000").Value'),
    "written in manual calculation": (
        'Application.Calculation = -4135\nws.Range("B2:B3000").Formula = "=B1+1"\nws.Range("B1").Value = 1\n'
        "Application.Calculation = -4105", 'ws.Range("B3000").Value'),
    "the bottom read from the top": (
        'Application.Calculation = -4135\nws.Range("C2:C3000").Formula = "=C1+1"\nws.Range("C1").Value = 1\n'
        'ws.Range("D1").Formula = "=C3000*2"\nApplication.Calculation = -4105', 'ws.Range("D1").Value'),
    "a sum over the chain": (
        'ws.Range("A1").Value = 1\nws.Range("A2:A3000").Formula = "=A1+1"\nws.Range("E1").Formula = "=SUM(A1:A3000)"',
        'ws.Range("E1").Value'),
    "each cell reading two": (
        'ws.Range("H1:H2").Value = 1\nws.Range("H3:H3000").Formula = "=MOD(H1+H2,1000)"', 'ws.Range("H3000").Value'),
    "running upward": ('ws.Range("I3000").Value = 1\nws.Range("I1:I2999").Formula = "=I2+1"', 'ws.Range("I1").Value'),
    "across two columns": (
        'ws.Range("J1").Value = 1\nws.Range("K1:K1500").FormulaR1C1 = "=RC[-1]+1"\n'
        'ws.Range("J2:J1500").FormulaR1C1 = "=R[-1]C[1]+1"', 'ws.Range("K1500").Value'),
}

#: The deepest formulas Excel takes, each within its limits (scripts/measure_formula_refusals.py).
FORMULAS: dict[str, str] = {
    "4096 terms": "=" + "1+" * 4095 + "1",
    "4096 powers": "=" + "1^" * 4095 + "1",
    "2730 cells added": "=" + "A1+" * 2729 + "A1",
    # In R1C1 each A1 is R[-7]C[-2], so 744 of them spell 8183 characters and 745 spell 8194.
    "744 cells added": "=" + "A1+" * 743 + "A1",
    "745 cells added": "=" + "A1+" * 744 + "A1",
    "2048 texts joined": "=" + '"a"&' * 2047 + '"a"',
    "256 brackets": "=" + "(" * 256 + "1" + ")" * 256,
    "256 brackets round a chain": "=" + "(" * 256 + "1+" * 2000 + "1" + ")" * 256,
    "a chain of 256 brackets": "=" + "(" * 256 + "1" + ")+1" * 256,
    "256 brackets opening rightward": "=" + "1+(" * 256 + "1" + ")" * 256,
    "65 functions": "=" + "ABS(" * 65 + "1" + ")" * 65,
    "65 functions each in brackets": "=" + "(ABS(" * 65 + "1" + "))" * 65,
    "200 brackets round 65 functions": "=" + "(" * 200 + "ABS(" * 65 + "1" + ")" * 265,
    "1000 minus signs": "=" + "-" * 1000 + "1",
    "999 minus signs": "=" + "-" * 999 + "1",
    "1000 plus signs": "=" + "+" * 1000 + "1",
    "8190 percent signs": "=1" + "%" * 8190,
    "4 percent signs": "=1" + "%" * 4,
}
#: Formulas written through FormulaR1C1 instead: 2000 whole rows, R1, spell $1:$1 in A1, past 8192 characters.
R1C1_FORMULAS: dict[str, str] = {
    "100 rows added": "=" + "R1+" * 99 + "R1",
    "1365 rows added": "=" + "R1+" * 1364 + "R1",
    "1366 rows added": "=" + "R1+" * 1365 + "R1",
    "2000 rows added": "=" + "R1+" * 1999 + "R1",
}


def _chains() -> list[str]:
    lines: list[str] = []
    for setup, expression in CHAINS.values():
        lines += ["Set ws = ActiveWorkbook.Worksheets.Add", setup, f"out = out & CStr({expression}) & \"|\"",
                  "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True"]
    return lines


def _written() -> list[tuple[str, str]]:
    """Each formula with the member it is written through."""
    return [*(("Formula", formula) for formula in FORMULAS.values()),
            *(("FormulaR1C1", formula) for formula in R1C1_FORMULAS.values())]


def _formulas() -> list[str]:
    lines = ["Set ws = ActiveWorkbook.Worksheets.Add", 'ws.Range("A1").Value = 1']
    for index, (member, _) in enumerate(_written()):
        lines += ["On Error Resume Next", 'ws.Range("C8").ClearContents', "Err.Clear",
                  f'ws.Range("C8").{member} = Formula{index}()',
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number & "|" Else out = out & Read(ws) & "|"',
                  "On Error GoTo 0"]
    return [*lines, "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True"]


#: What is read of each formula: its value, how long its Formula, FormulaR1C1 and Formula2 are, and how its
#: FormulaR1C1 ends; or ! and the error number Excel raised reading it.
_READ = [
    "Private Function Read(ws As Object) As String",
    "Dim parts(4) As String",
    "On Error Resume Next",
    'Err.Clear: parts(0) = CStr(ws.Range("C8").Value): If Err.Number <> 0 Then parts(0) = "!" & Err.Number',
    'Err.Clear: parts(1) = Len(ws.Range("C8").Formula): If Err.Number <> 0 Then parts(1) = "!" & Err.Number',
    'Err.Clear: parts(2) = Len(ws.Range("C8").FormulaR1C1): If Err.Number <> 0 Then parts(2) = "!" & Err.Number',
    'Err.Clear: parts(3) = Len(ws.Range("C8").Formula2): If Err.Number <> 0 Then parts(3) = "!" & Err.Number',
    'Err.Clear: parts(4) = Right(ws.Range("C8").FormulaR1C1, 12): If Err.Number <> 0 Then parts(4) = "!" & Err.Number',
    'Read = Join(parts, "~")',
    "End Function",
]


def _builders() -> list[str]:
    """A function giving each formula, built in pieces a line of VBA can hold; one each, so no procedure is too
    large."""
    lines: list[str] = []
    for index, (_, formula) in enumerate(_written()):
        pieces = [formula[start:start + PIECE].replace('"', '""') for start in range(0, len(formula), PIECE)]
        lines += [f"Private Function Formula{index}() As String", "Dim f As String",
                  *[f'f = f & "{piece}"' for piece in pieces], f"Formula{index} = f", "End Function"]
    return lines


def main() -> None:
    module = "\n".join(["Public Function Probe() As String", "Dim ws As Object, out As String, failed As Long",
                        *_chains(), *_formulas(), "Probe = out", "End Function", *_READ, *_builders()]) + "\n"
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module, "Probe", timeout=300.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    answers = str(result.value).split("|")[:-1]
    assert len(answers) == len(CHAINS) + len(_written()), len(answers)
    chains = {name: {"setup": setup, "read": expression, "answer": answer}
              for (name, (setup, expression)), answer in zip(CHAINS.items(), answers, strict=False)}
    formulas: dict[str, dict[str, str | list[str]]] = {}
    names = [*FORMULAS, *R1C1_FORMULAS]
    for name, (member, formula), answer in zip(names, _written(), answers[len(CHAINS):], strict=True):
        if answer.startswith("!"):
            formulas[name] = {"member": member, "formula": formula, "refused": answer}
            continue
        value, *lengths, end = answer.split("~")
        formulas[name] = {"member": member, "formula": formula, "value": value, "lengths": lengths, "r1c1_end": end}
    OUT.write_text(json.dumps({"chains": chains, "formulas": formulas}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
