"""Which functions Excel takes where a formula has to name cells, for every function the formula engine has.

Excel will not take a formula (error 1004) that puts a function answering
with a value where cells are wanted, COUNTIF(SUM(A1:A3),1); a function
that can answer with cells, COUNTIF(INDEX(A1:A3,0),1), is taken. Each
function is written on its own and then into COUNTIF's range through
Range.Formula, with as many arguments as it needs, each A1, and what
Excel did is kept:

    python scripts/measure_cells_functions.py

writes tests/fixtures/formula/cells_functions.json, which
tests/test_formula.py replays: per function the call, whether Excel took
it on its own (and #NAME? there for one this Excel has not got), and
whether it took it in COUNTIF's range.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.registry import FUNCTIONS

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "formula" / "cells_functions.json"
#: Calls whose arguments are not a list of values, or which Excel takes with more than the engine needs.
#: ANCHORARRAY is the file's spelling of A1#, which Excel does not take written as a call.
SPECIAL = {"LET": "LET(x,A1,x)", "LAMBDA": "LAMBDA(x,x)(A1)", "INDEX": "INDEX(A1,1)", "MAP": "MAP(A1,LAMBDA(x,x))"}
BATCH = 60


def call(name: str) -> str:
    """The function called with as many arguments as it needs, each A1."""
    return SPECIAL.get(name) or f"{name}({','.join(['A1'] * FUNCTIONS[name].minimum)})"


def module(names: list[str]) -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String, v As Variant, shown As String",
             "Set ws = ActiveSheet", 'ws.Range("A1").Value = 1', "On Error Resume Next"]
    for name in names:
        text = call(name).replace('"', '""')
        lines += ["Err.Clear", f'ws.Range("H1").Formula = "={text}"',
                  "If Err.Number <> 0 Then",
                  f'    out = out & "{name}|!" & Err.Number & "|" & vbLf',
                  "Else",
                  '    v = ws.Range("H1").Value',
                  '    If IsError(v) Then shown = CStr(v) Else shown = "ok"',
                  "    Err.Clear",
                  f'    ws.Range("H2").Formula = "=COUNTIF({text},1)"',
                  f'    out = out & "{name}|" & shown & "|" & Err.Number & vbLf',
                  "End If",
                  'ws.Range("H1:H2").ClearContents']
    lines += ["On Error GoTo 0", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    names = sorted(FUNCTIONS)
    found: dict[str, dict[str, object]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for start in range(0, len(names), BATCH):
            batch = names[start:start + BATCH]
            result = excel.run_vba(module(batch), "Probe", timeout=300.0)
            assert result.ok, result.message
            for line in str(result.value).split("\n"):
                if not line:
                    continue
                name, alone, in_cells = line.split("|")
                found[name] = {"call": call(name), "alone": alone, "in_cells": int(in_cells) if in_cells else None}
            print(f"measured {start + len(batch)} of {len(names)}", flush=True)
    FIXTURE.write_text(json.dumps(found, indent=1) + "\n", encoding="utf-8")
    taken = sorted(name for name, one in found.items() if one["in_cells"] == 0 and one["alone"] != "Error 2029")
    print("taken where cells are wanted:", taken)


if __name__ == "__main__":
    main()
