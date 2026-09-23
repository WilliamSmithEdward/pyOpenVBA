"""What long chains of cells come to in Excel: a running total thousands of rows long, and its kin.

Each case is VBA run on a fresh worksheet, then an expression read back:

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

#: Each case: the VBA run on a fresh sheet ``ws``, and the expression then read.
CASES: dict[str, tuple[str, str]] = {
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


def _module() -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String"]
    for setup, expression in CASES.values():
        lines += ["Set ws = ActiveWorkbook.Worksheets.Add", setup, f"out = out & CStr({expression}) & \"|\"",
                  "Application.DisplayAlerts = False", "ws.Delete", "Application.DisplayAlerts = True"]
    lines += ["Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(_module(), "Probe", timeout=300.0)
    assert result.ok, f"{result.outcome}: {result.message}"
    answers = str(result.value).split("|")[:-1]
    record = {name: {"setup": setup, "read": expression, "answer": answer}
              for (name, (setup, expression)), answer in zip(CASES.items(), answers, strict=True)}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
