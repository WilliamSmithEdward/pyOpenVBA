"""Long chains of cells and the deepest formulas, worked out as Excel works them out.

tests/fixtures/long_chains.json is what scripts/measure_long_chains.py
saw in live Excel. Chains: a running total 3000 rows long, the same with
its top changed, written in manual calculation, read from a cell above
it, summed, running upward, and across two columns. Formulas: the
deepest Excel takes -- 4096 terms, 256 brackets, 65 functions, a
thousand signs, 8190 percent signs -- each with its value and the length
of its Formula, FormulaR1C1 and Formula2. The model works a cell out on
demand, one inside another, puts off a cell too many cells down, and
goes down a formula's tree on a deeper stack when Python's runs out
(pyopenvba.formula._deep), so neither kind runs Python out of stack.
"""

from __future__ import annotations

import json
import tempfile
import time
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = Path(__file__).parent / "fixtures" / "long_chains.json"
RECORD: dict[str, dict[str, dict[str, Any]]] = json.loads(FIXTURE.read_text(encoding="utf-8"))
CHAINS = RECORD["chains"]
FORMULAS = RECORD["formulas"]

_COMPILED = "Excel keeps a formula only while its compiled tokens fit in 16384 bytes; the model has no compiled form"
#: Formulas the model does not answer as Excel does, and why.
GAPS = {
    "2730 cells added": _COMPILED + ", and Excel reads one too big as R1C1, where A1 is a name",
    **dict.fromkeys(("1365 rows added", "1366 rows added", "2000 rows added"), _COMPILED),
}


def _run(setup: str, read: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Function Probe() As String\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets.Add\n"
                   f"{setup}\nProbe = CStr({read})\nEnd Function\n", name="Chains")
    return str(app.run("Probe"))


@pytest.mark.parametrize("name", list(CHAINS))
def test_a_long_chain_comes_to_what_excel_shows(name: str) -> None:
    case = CHAINS[name]
    assert _run(case["setup"], case["read"]) == case["answer"]


def test_a_change_at_the_top_of_a_long_chain_is_not_quadratic() -> None:
    # Marking what reads a changed cell finds the readers through an index, not by going through every formula.
    started = time.perf_counter()
    assert _run('ws.Range("A1").Value = 1\nws.Range("A2:A20000").Formula = "=A1+1"\nws.Range("A1").Value = 5',
                'ws.Range("A20000").Value') == "20004"
    assert time.perf_counter() - started < 30


def test_a_long_circle_of_cells_ends() -> None:
    # Cells that feed themselves, three hundred of them round: each is worked out once, and the circle noticed.
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.set_value("C1", "=C300+1")
    app.add_module('Public Sub Fill()\nRange("C2:C300").Formula = "=C1+1"\nEnd Sub\n', name="Circle")
    app.run("Fill")
    app.evaluate('Range("C300").Value')
    assert app.workbook.calculator.circular


@pytest.fixture(scope="module")
def app() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).set_value("A1", 1)
    app.add_module('Public Sub Assign(ByVal member As String, ByVal f As String)\nRange("C8").ClearContents\n'
                   'If member = "Formula" Then Range("C8").Formula = f Else Range("C8").FormulaR1C1 = f\nEnd Sub\n',
                   name="Writing")
    return app


def _read(app: ExcelApplication, expression: str) -> str:
    try:
        return str(app.evaluate(expression))
    except VBARuntimeError as failure:
        return f"!{failure.number}"


@pytest.mark.parametrize("name", [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in FORMULAS])
def test_the_deepest_formulas_come_to_what_excel_shows(app: ExcelApplication, name: str) -> None:
    case = FORMULAS[name]
    try:
        app.run("Assign", case["member"], case["formula"])
    except VBARuntimeError as failure:
        assert f"!{failure.number}" == case.get("refused")
        return
    assert "refused" not in case
    assert _read(app, 'CStr(Range("C8").Value)') == case["value"]
    assert [_read(app, f'Len(Range("C8").{member})') for member in ("Formula", "FormulaR1C1", "Formula2")] \
        == case["lengths"]
    assert _read(app, 'Right(Range("C8").FormulaR1C1, 12)') == case["r1c1_end"]


def test_the_deepest_formulas_fill_copy_and_save() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.sheet(1)
    sheet.set_value("A1", 1)
    deepest = {name: FORMULAS[name]["formula"] for name in ("4096 terms", "256 brackets", "8190 percent signs")}
    for row, formula in enumerate(deepest.values(), start=2):
        sheet.set_value(f"C{row}", formula)
    app.add_module('Public Sub Spread()\nRange("C2:E4").FillRight\nRange("C2:C4").Copy Range("G2")\nEnd Sub\n',
                   name="Spread")
    app.run("Spread")
    assert [app.evaluate(f'Range("{column}2").Value') for column in "CDEG"] == [4096.0] * 4
    with tempfile.TemporaryDirectory() as folder:
        path = app.save(Path(folder) / "deep.xlsx")
        again = ExcelApplication.open(path, with_vba=False)
        for row, formula in enumerate(deepest.values(), start=2):
            assert again.evaluate(f'Range("C{row}").Formula') == formula
        assert again.evaluate('Range("C3").Value') == 1.0
