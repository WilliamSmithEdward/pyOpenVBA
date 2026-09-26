"""Edges cell styles draw: scripts/measure_shared_edges.py, replayed against the model.

shared_edges.xlsx holds cells whose styles both draw the edge between
them; the model reads each edge as Excel did.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / "fixtures" / "cell_styles"
EDGES: dict[str, Any] = json.loads((FIXTURES / "shared_edges.json").read_text(encoding="utf-8"))


def test_an_edge_two_cells_draw_reads_as_excel_reads_it() -> None:
    app = ExcelApplication.open(FIXTURES / "shared_edges.xlsx", with_vba=False)
    app.add_module(EDGES["reader"] + "Public Function Run() As String\nRun = Edges(ActiveWorkbook.Worksheets(1))\n"
                   "End Function\n", name="Reader")
    ours = str(app.run("Run")).split("|")[: len(EDGES["pairs"])]
    assert [pair for pair, mine, want in zip(EDGES["pairs"], ours, EDGES["answers"], strict=True) if mine != want] == []
