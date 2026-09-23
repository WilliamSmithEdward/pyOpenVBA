"""Range.FillDown, FillUp, FillRight and FillLeft, replayed against the in-memory model.

tests/fixtures/fill.json is what scripts/measure_fill.py saw in live Excel:
layouts that fill a range one way and dump A1:E6 -- each cell's formula,
number format, bold and fill colour -- with the used range and what the
method returned. Each layout runs on a fresh sheet in the model. A fill
over merged cells, which copies the merge along in ways the dump does not
show, reports itself unsupported.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "fill.json").read_text(encoding="utf-8"))
MERGED = [layout for layout in RECORD["layouts"] if ".Merge" in layout["setup"]]
PLAIN = [layout for layout in RECORD["layouts"] if layout not in MERGED]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", PLAIN, ids=[layout["name"] for layout in PLAIN])
def test_a_fill_leaves_what_excel_leaves(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]


@pytest.mark.parametrize("layout", MERGED, ids=[layout["name"] for layout in MERGED])
def test_a_fill_over_merged_cells_says_it_is_not_modelled(layout: dict[str, Any]) -> None:
    with pytest.raises(VBAUnsupportedError, match="merged"):
        _run(layout)
