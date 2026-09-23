"""Cell shifts -- Range.Delete and Range.Insert with a Shift, or none -- replayed against the in-memory model.

tests/fixtures/cell_shifts.json is what scripts/measure_cell_shifts.py saw
in live Excel: a grid of numbers on a sheet named Grid, formulas reading
it from column H, from cells that move with it and from a second sheet,
two names, and sometimes a filter; then one delete or insert. The dump
holds every formula on Grid by address, the grid's cells, the other
sheet's formulas, the names, the filter's range and hidden name, and five
cells' formats.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "cell_shifts.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, other As Object, failed As String",
            "Dim r As Long, c As Long", "Set ws = ActiveWorkbook.Worksheets(1)", 'ws.Name = "Grid"',
            "Set other = ActiveWorkbook.Worksheets.Add(After:=ws)", 'other.Name = "Other"', "ws.Activate",
            *RECORD["grid"].splitlines(), "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            "Report = failed & Dump(ws)", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_references_follow_a_cell_shift_as_excel_moves_them(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]
