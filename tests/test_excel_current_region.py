"""CurrentRegion, and End over cells that hold only a format, replayed against the in-memory model.

tests/fixtures/current_region.json is what scripts/measure_current_region.py
saw in live Excel: layouts that each ask one question -- whether a
diagonal neighbour joins a region, whether a formula returning "" or a
cell with only a format does, what a merged cell or a hidden row does,
and what a whole row, a whole column or several areas answer -- and eight
seeded random grids read from every other cell. Each layout is set up on
a fresh sheet in the model and read the same way.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / "fixtures" / "current_region.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _layout_module(layout: dict[str, Any]) -> str:
    """The probe's own functions for one layout: its setup, then each read into one line."""
    reads = ["Private Function Reads(ws As Object) As String", "Dim out As String, v As Variant",
             "On Error Resume Next"]
    for expression in layout["reads"]:
        reads += ["Err.Clear", "v = Empty", f"v = {expression}",
                  'If Err.Number <> 0 Then out = out & "E" & Err.Number & ";" Else out = out & Show(v) & ";"']
    reads += ["On Error GoTo 0", "Reads = out", "End Function"]
    case = ["Public Function Report() As String", "Dim ws As Object, failed As String",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            "Report = failed & Reads(ws)", "End Function"]
    return RECORD["helper"] + "\n".join(reads + case) + "\n"


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_layout_reads_as_excel_reads_it(layout: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_layout_module(layout), name="Probe")
    assert app.run("Report") == layout["answers"]
