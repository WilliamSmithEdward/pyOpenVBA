"""Long chains of cells, worked out as Excel works them out.

tests/fixtures/long_chains.json is what scripts/measure_long_chains.py
saw in live Excel: a running total 3000 rows long, the same with its top
changed, written in manual calculation, read from a cell above it, summed,
running upward, and across two columns. The model works a cell out on
demand, one inside another, and puts off a cell too many cells down, so
no chain runs Python out of stack.
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE = Path(__file__).parent / "fixtures" / "long_chains.json"
CASES: dict[str, dict[str, Any]] = json.loads(FIXTURE.read_text(encoding="utf-8"))


def _run(setup: str, read: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Function Probe() As String\nDim ws As Object\nSet ws = ActiveWorkbook.Worksheets.Add\n"
                   f"{setup}\nProbe = CStr({read})\nEnd Function\n", name="Chains")
    return str(app.run("Probe"))


@pytest.mark.parametrize("name", list(CASES))
def test_a_long_chain_comes_to_what_excel_shows(name: str) -> None:
    case = CASES[name]
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
