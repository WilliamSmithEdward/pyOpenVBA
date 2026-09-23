"""Range.SpecialCells, replayed against the in-memory model.

tests/fixtures/special_cells.json is what scripts/measure_special_cells.py
saw in live Excel: layouts asked, through several ranges, for constants
and formulas of every kind, blanks, the last cell and visible cells, each
answer the address of what came back and its number of areas. Each
layout is set up on a fresh sheet in the model and asked the same.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._special_cells import areas_of

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "special_cells.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _layout_module(layout: dict[str, Any]) -> str:
    lines = ["Public Function Report() As String", "Dim ws As Object, failed As String, out As String",
             "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
             'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
             *(f'out = out & {expression} & ";"' for expression in layout["reads"]),
             "Report = failed & out", "End Function"]
    return RECORD["helper"] + "\n".join(lines) + "\n"


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_layout_answers_as_excel_answers(layout: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_layout_module(layout), name="Probe")
    got = str(app.run("Report")).split(";")
    want = layout["answers"].split(";")
    wrong = [(read, expected, found) for read, expected, found in zip(layout["reads"], want, got, strict=False)
             if expected != found]
    assert not wrong


def test_a_block_is_split_from_its_last_cell_back() -> None:
    """The splitting on its own: stacked rows of one span merge, a lone column stays a column."""
    cells = [(row, column) for row in range(1, 9) for column in range(1, 7) if (row, column) != (2, 2)]
    assert areas_of(cells) == [[1, 2, 1, 2], [1, 3, 2, 6], [1, 1, 2, 1], [3, 1, 8, 6]]
