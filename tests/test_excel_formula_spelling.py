"""How Excel spells a formula a macro writes, replayed against the in-memory model.

tests/fixtures/formula_spelling.json is what scripts/measure_formula_spelling.py
saw in live Excel: 120 formulas written in turn through Range.Formula into
one cell of a sheet called Data, beside a sheet called My Sheet, with a
workbook name MyName, each read back through Formula and FormulaR1C1. The
model runs the same probe, in the same order, since the spelling of a name
nothing defines is the one the workbook saw first.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "formula_spelling.json").read_text(encoding="utf-8"))
CASES: list[dict[str, Any]] = RECORD["cases"]


def _module() -> str:
    lines = [RECORD["helper"], "Public Function Probe() As String", "Dim wb As Object, ws As Object, out As String",
             "Set wb = ThisWorkbook", "Set ws = wb.Worksheets(1)", 'ws.Name = "Data"',
             'wb.Worksheets.Add(After:=ws).Name = "My Sheet"', 'wb.Names.Add Name:="MyName", RefersTo:="=Data!$B$1"']
    lines += [f'out = out & Spelled(ws, {case["formula"]}) & "|"' for case in CASES]
    return "\n".join([*lines, "Probe = out", "End Function"]) + "\n"


@pytest.fixture(scope="module")
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_module(), name="Probe")
    return str(app.run("Probe")).split("|")[: len(CASES)]


@pytest.mark.parametrize("index", range(len(CASES)), ids=[case["formula"] for case in CASES])
def test_a_formula_is_spelled_as_excel_spells_it(answers: list[str], index: int) -> None:
    assert answers[index] == CASES[index]["answer"]
