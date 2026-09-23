"""Range.Cut and the references that overlap what it moves, replayed against the in-memory model.

tests/fixtures/cut_references.json is what scripts/measure_cut_references.py
saw in live Excel: numbers on a sheet named Src, formulas reading blocks,
cells and columns of them, a formula on a second sheet and a name; then
one cut, on Src or to the other sheet, Dst. The dump holds Src's formulas,
Dst's, the name, and both sheets' cells.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "cut_references.json")
                                    .read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim src As Object, dst As Object, failed As String, r As Long",
            "Set src = ActiveWorkbook.Worksheets(1)", 'src.Name = "Src"',
            "Set dst = ActiveWorkbook.Worksheets.Add(After:=src)", 'dst.Name = "Dst"', "src.Activate",
            *RECORD["setup"].splitlines(), "On Error Resume Next", "Err.Clear", *layout["cut"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            "Report = failed & Dump(src, dst)", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_references_follow_a_cut_as_excel_moves_them(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]
