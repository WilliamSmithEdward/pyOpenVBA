"""What Excel refuses when a formula is written, replayed against live Excel.

tests/fixtures/formula_refusals.json is what
scripts/measure_formula_refusals.py saw writing formulas through
Range.Formula and its kin on a sheet with names over it: each of Excel's
functions with 0 to 16 and 249 to 256 arguments; an operation in each
argument, one at a time, of each function that reads cells; what such an
argument takes; LET and LAMBDA and their names; and the limits on a
formula's text, length, brackets, calls and signs. Excel refused some
with error 1004, a formula too long with error 7. The model writes each
the same way, and takes or refuses it as Excel did.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

FIXTURE = Path(__file__).parent / "fixtures" / "formula_refusals.json"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
#: The measurement's setup without the formula it spills into E1, which only E1# reads.
SETUP = "\n".join(line for line in RECORD["setup"].splitlines() if "Formula2" not in line)
FORMS = [(name, formula) for name, formulas in RECORD["forms"].items() for formula in formulas]

_SPILLS = "a formula reaching a spilled range with # waits for dynamic arrays that spill"
#: Values the model does not show as Excel does, and why.
VALUE_GAPS: dict[str, str] = {
    "=GETPIVOTDATA(A1,A1,A1)": "GETPIVOTDATA is not implemented",
}


def _app(*, manual: bool) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Setup(ByVal manual As Boolean)\nDim ws As Object\n"
                   "Set ws = ActiveWorkbook.Worksheets(1)\n" + SETUP + "\n"
                   "If manual Then Application.Calculation = -4135\nEnd Sub\n"
                   "Public Sub Assign(ByVal member As String, ByVal f As String)\n"
                   'Range("C8").ClearContents\nSelect Case member\nCase "Formula": Range("C8").Formula = f\n'
                   'Case "Formula2": Range("C8").Formula2 = f\nCase "FormulaR1C1": Range("C8").FormulaR1C1 = f\n'
                   'Case "FormulaArray": Range("C8").FormulaArray = f\nCase "Value": Range("C8").Value = f\n'
                   "End Select\nEnd Sub\n"
                   "Public Function Shown() As String\n"
                   'Shown = Range("C8").Formula & "~" & Range("C8").Text\nEnd Function\n', name="Setup")
    app.run("Setup", manual)
    return app


@pytest.fixture(scope="module")
def manual() -> ExcelApplication:
    """A workbook in manual calculation, as the measurement wrote its formulas, so none is worked out."""
    return _app(manual=True)


@pytest.fixture(scope="module")
def automatic() -> ExcelApplication:
    return _app(manual=False)


def _taken(app: ExcelApplication, formula: str, member: str = "Formula") -> str:
    """ok, or ! and the error number the model raised writing ``formula`` through ``member``."""
    try:
        app.run("Assign", member, formula)
    except VBARuntimeError as failure:
        return f"!{failure.number}"
    except VBAUnsupportedError as gap:
        return f"unsupported: {gap}"
    return "ok"


@pytest.mark.parametrize("name", list(RECORD["counts"]))
def test_a_function_takes_the_arguments_excel_takes(manual: ExcelApplication, name: str) -> None:
    measured = RECORD["counts"][name]
    found = {count: _taken(manual, f"={name}({','.join(['A1'] * int(count))})") for count in measured}
    assert found == measured


@pytest.mark.parametrize("name", list(RECORD["places"]))
def test_a_function_wants_cells_where_excel_does(manual: ExcelApplication, name: str) -> None:
    measured = RECORD["places"][name]
    assert {formula: _taken(manual, formula) for formula in measured} == measured


@pytest.mark.parametrize(("name", "formula"), [
    pytest.param(name, formula, id=formula,
                 marks=[pytest.mark.xfail(reason=_SPILLS, strict=True)] if "E1#" in formula else [])
    for name, formula in FORMS])
def test_where_cells_are_wanted_takes_what_excel_takes(manual: ExcelApplication, name: str, formula: str) -> None:
    assert _taken(manual, formula) == RECORD["forms"][name][formula]


@pytest.mark.parametrize("formula", [
    pytest.param(formula, id=formula[:60],
                 marks=[pytest.mark.xfail(reason=VALUE_GAPS[formula], strict=True)] if formula in VALUE_GAPS else [])
    for formula in RECORD["values"]])
def test_let_lambda_and_edge_calls_come_to_what_excel_shows(automatic: ExcelApplication, formula: str) -> None:
    taken = _taken(automatic, formula)
    if taken != "ok":
        assert taken == RECORD["values"][formula]
        return
    assert automatic.run("Shown") == RECORD["values"][formula]


@pytest.mark.parametrize("label", list(RECORD["limits"]))
def test_a_formula_past_excels_limits_is_refused(manual: ExcelApplication, label: str) -> None:
    member, formula, answer = RECORD["limits"][label]
    assert _taken(manual, formula, member) == answer
