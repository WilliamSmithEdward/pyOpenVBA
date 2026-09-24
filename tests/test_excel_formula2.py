"""Formula2's @ on a formula written through Range.Formula, replayed against live Excel.

tests/fixtures/implicit_intersection.json is what
scripts/measure_implicit_intersection.py saw: each of Excel's functions
written through Range.Formula with each argument a cell, a range and an
operation on a range, with its optional and repeating arguments too, and
150 formulas besides, each read back through Formula2 -- or refused, as
error 1004. The model writes each the same way and reads it back the same.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

FIXTURE = Path(__file__).parent / "fixtures" / "implicit_intersection.json"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
#: Every formula written and what Formula2 read back, by an id naming where it came from.
CASES: dict[str, tuple[str, str]] = {
    **{f"{name}/{variant}": (written, read) for name, one in RECORD["functions"].items()
       for variant, (written, read) in one.items()},
    **{f"formula {written}": (written, read) for written, read in RECORD["formulas"].items()},
}

_SINGLE = "Excel writes SINGLE(x) as @x when the formula is written; the model keeps SINGLE"
#: Cases the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {
    "SINGLE/cell": _SINGLE, "SINGLE/range": _SINGLE, "SINGLE/worked": _SINGLE,
    "formula =@A1:A3": "Range.Formula given an @ is not implemented: what Formula reads back is not measured",
}


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Sub Setup()\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets(1)\n"
                   + RECORD["setup"] + "\nEnd Sub\n", name="Setup")
    app.run("Setup")
    return app


def _read(app: ExcelApplication, formula: str) -> str:
    """What Formula2 reads back of ``formula`` written through Range.Formula, or ! and the error number."""
    sheet = app.sheet(1)
    try:
        sheet.set_value("C8", formula)
    except VBARuntimeError as failure:
        return f"!{failure.number}"
    except VBAUnsupportedError as gap:
        return f"unsupported: {gap}"
    return str(app.evaluate('Range("C8").Formula2'))


@pytest.mark.parametrize("case", [
    pytest.param(case, id=case, marks=[pytest.mark.xfail(reason=GAPS[case], strict=True)] if case in GAPS else [])
    for case in CASES])
def test_formula2_reads_a_formula_as_excel_does(app: ExcelApplication, case: str) -> None:
    written, read = CASES[case]
    assert _read(app, written) == read


def test_every_gap_is_a_case() -> None:
    assert set(GAPS) <= set(CASES)
