"""The number format each function gives a General cell its formula is written to.

Excel gives a cell a format once, as a formula is written to it, from the
formula's shape (see src/pyopenvba/apps/excel/_formula_format.py). Each of
the formula engine's functions is written on its own to a General cell
through Range.Formula, as scripts/measure_cells_functions.py calls it, with
as many arguments as it needs, each A1 holding 1, and the cell's
NumberFormat is kept, then the cell is cleared. A second list writes
financial functions with arguments a workbook would give them, and calls
inside other formulas, so that what a function brings can be told from
what its arguments do.

    python scripts/measure_function_formats.py

writes tests/fixtures/function_formats.json, which
tests/test_excel_function_formats.py replays.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from measure_cells_functions import call
from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.registry import FUNCTIONS

ROOT = Path(__file__).resolve().parent.parent
FIXTURE = ROOT / "tests" / "fixtures" / "function_formats.json"
BATCH = 60
#: Formulas written as a workbook would write them: financial functions with their usual arguments, and calls
#: inside other formulas.
FORMULAS = [
    "=NPV(0.1,100,200)", "=PV(0.05,10,-100)", "=FV(0.05,10,-100)", "=PMT(0.05,10,1000)", "=IPMT(0.05,1,10,1000)",
    "=PPMT(0.05,1,10,1000)", "=RATE(10,-100,800)", "=IRR({-100,30,40,50})", "=MIRR({-100,30,40,50},0.1,0.1)",
    "=NPER(0.05,-100,800)", "=XNPV(0.1,{-100,50,60},{43831,44197,44562})", "=XIRR({-100,50,60},{43831,44197,44562})",
    "=EFFECT(0.1,12)", "=NOMINAL(0.1,12)", "=SLN(1000,100,5)", "=SYD(1000,100,5,1)", "=DB(1000,100,5,1)",
    "=DDB(1000,100,5,1)", "=VDB(1000,100,5,0,1)", "=CUMIPMT(0.05,10,1000,1,1,0)", "=CUMPRINC(0.05,10,1000,1,1,0)",
    "=ISPMT(0.05,1,10,1000)", "=PDURATION(0.05,100,200)", "=RRI(10,100,200)", "=FVSCHEDULE(100,{0.1,0.2})",
    "=DOLLARDE(1.02,16)", "=DOLLARFR(1.125,16)", "=PRICE(43831,47484,0.05,0.06,100,2)",
    "=YIELD(43831,47484,0.05,95,100,2)", "=ACCRINT(43831,44013,44197,0.05,1000,2)",
    "=NPV(0.1,100,200)+1", "=1+NPV(0.1,100,200)", "=NPV(0.1,100,200)*2", "=-PMT(0.05,10,1000)",
    "=ROUND(PMT(0.05,10,1000),2)", "=SUM(PV(0.05,10,-100),1)", "=IF(TRUE,PMT(0.05,10,1000))", "=(FV(0.05,10,-100))",
    "=NPV(0.1,100,200)-PV(0.05,10,-100)", "=DATE(2020,1,1)+PMT(0.05,10,1000)", "=PMT(0.05,10,1000)+DATE(2020,1,1)",
    "=ABS(PMT(0.05,10,1000))", "=MAX(PMT(0.05,10,1000),0)", "=RATE(10,-100,800)+0", "=IRR({-100,30,40,50})*100",
    "=HYPERLINK(\"http://example.com\",\"x\")", "=DOLLAR(1234.5)", "=TEXT(1,\"0%\")",
]
#: Formulas whose cell is read back: the type its Value arrives as, Value, Text and Value2.
READBACKS = ["=NPV(0.1,1,2)", "=PMT(0.05,10,1000)", "=RATE(10,-100,800)", "=SLN(1000,100,3)", "=IRR({-100,60,60})"]


def readback_module(formulas: list[str]) -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String", "Set ws = ActiveSheet"]
    for formula in formulas:
        text = formula.replace('"', '""')
        lines += ['ws.Range("H1").Clear', f'ws.Range("H1").Formula = "{text}"',
                  'out = out & TypeName(ws.Range("H1").Value) & "|" & CStr(ws.Range("H1").Value) & "|" & '
                  'ws.Range("H1").Text & "|" & CStr(ws.Range("H1").Value2) & vbLf']
    lines += ['ws.Range("H1").Clear', "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def module(formulas: list[str]) -> str:
    lines = ["Public Function Probe() As String", "Dim ws As Object, out As String",
             "Set ws = ActiveSheet", 'ws.Range("A1").Value = 1', "On Error Resume Next"]
    for formula in formulas:
        text = formula.replace('"', '""')
        lines += ['ws.Range("H1").Clear', "Err.Clear", f'ws.Range("H1").Formula = "{text}"',
                  'If Err.Number <> 0 Then out = out & "!" & Err.Number & vbLf Else '
                  'out = out & ws.Range("H1").NumberFormat & vbLf']
    lines += ['ws.Range("H1").Clear', "On Error GoTo 0", "Probe = out", "End Function"]
    return "\n".join(lines) + "\n"


def measure(excel: ExcelSession, formulas: list[str]) -> list[str]:
    found: list[str] = []
    for start in range(0, len(formulas), BATCH):
        batch = formulas[start:start + BATCH]
        result = excel.run_vba(module(batch), "Probe", timeout=300.0)
        assert result.ok and not result.dialogs, result
        lines = str(result.value).split("\n")[:-1]
        assert len(lines) == len(batch), (batch, lines)
        found += lines
        print(f"measured {start + len(batch)} of {len(formulas)}", flush=True)
    return found


def main() -> None:
    names = sorted(FUNCTIONS)
    calls = ["=" + call(name) for name in names]
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\nEnd Function',
                                "V")
        functions_found = measure(excel, calls)
        formulas_found = measure(excel, FORMULAS)
        readback = excel.run_vba(readback_module(READBACKS), "Probe", timeout=300.0)
        assert readback.ok and not readback.dialogs, readback
    FIXTURE.write_text(json.dumps({
        "excel_version_build": version.value,
        "measured_at": datetime.now(timezone.utc).isoformat(),
        "functions": {name: {"formula": formula, "format": shown}
                      for name, formula, shown in zip(names, calls, functions_found)},
        "formulas": dict(zip(FORMULAS, formulas_found)),
        "readback": dict(zip(READBACKS, str(readback.value).split("\n")[:-1])),
    }, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    brought = {name: shown for name, shown in zip(names, functions_found) if shown not in ("General",)}
    print("functions that bring a format:", brought)


if __name__ == "__main__":
    main()
