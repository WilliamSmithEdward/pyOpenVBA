"""A formula written in the other notation, replayed against live Excel.

tests/fixtures/formula_notation.json is what
scripts/measure_formula_notation.py saw: Range.Formula reads a formula
in A1, and where A1 cannot read it, as a name that looks like an R1C1
reference does, in R1C1, =R[-1]C+1 written to C8 being =C7+1; and
FormulaR1C1 reads what looks like a cell in A1 as a name, spelled 'A1'
in A1. Each formula is written the same way, through Formula, Formula2,
Value, FormulaArray or FormulaR1C1, and each cell's Formula, FormulaR1C1
and Text read back.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = Path(__file__).parent / "fixtures" / "formula_notation.json"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
CASES: dict[str, dict[str, str]] = {f"{case['member']} {case['formula']} {case['target']}": case
                                    for case in RECORD["cases"]}

#: Cases the model does not answer as Excel does, and why.
GAPS = {
    "Formula =A1.B C8": "A1.B is a field of a data type in a cell in Excel, #FIELD!; the model has no data types and "
                        "reads a name",
}


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Function Probe(ByVal member As String, ByVal f As String, ByVal target As String) "
                   "As String\n"
                   "Dim ws As Object, out As String, cell As Object\nSet ws = ActiveSheet\n"
                   'ws.Range("C8:F10").ClearContents\n'
                   "Select Case member\n"
                   'Case "Formula": ws.Range(target).Formula = f\nCase "Formula2": ws.Range(target).Formula2 = f\n'
                   'Case "Value": ws.Range(target).Value = f\nCase "FormulaArray": ws.Range(target).FormulaArray = f\n'
                   'Case "FormulaR1C1": ws.Range(target).FormulaR1C1 = f\n'
                   "End Select\n"
                   'out = "ok"\n'
                   "For Each cell In ws.Range(target).Cells\n"
                   '    out = out & "~" & cell.Formula & "~" & cell.FormulaR1C1 & "~" & cell.Text\n'
                   "Next\nProbe = out\nEnd Function\n"
                   "Public Sub Setup()\nDim ws As Object\nSet ws = ActiveSheet\n" + RECORD["setup"] + "\nEnd Sub\n",
                   name="Notation")
    app.run("Setup")
    return app


@pytest.mark.parametrize("key", [
    pytest.param(key, marks=[pytest.mark.xfail(reason=GAPS[key], strict=True)] if key in GAPS else [])
    for key in CASES])
def test_a_formula_in_the_other_notation_reads_as_in_excel(app: ExcelApplication, key: str) -> None:
    case = CASES[key]
    try:
        found = str(app.run("Probe", case["member"], case["formula"], case["target"]))
    except VBARuntimeError as failure:
        found = f"!{failure.number}~read!{failure.number}"
    assert found == case["answer"]
