"""Copy, Cut, Paste and PasteSpecial through the clipboard, replayed against the in-memory model.

tests/fixtures/paste.json is what scripts/measure_paste.py saw in live
Excel: layouts that fill a few cells and copy or cut them, pasting in one
of the ways a macro can, then dump A1:F6 -- each cell's formula, value
type, number format and bold -- with the used range, CutCopyMode, column
D's width, and what the last call returned. Each layout runs on a fresh
workbook in the model; the one that cuts to another sheet gets that sheet
first, named as the probe's first sheet was.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "paste.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
FIRST = LAYOUTS[0]["name"]


def _run(layout: dict[str, Any]) -> object:
    # The probe's first layout ran on its first sheet, which a later one cuts to.
    sheet = ["Set ws = ActiveWorkbook.Worksheets(1)"] if layout["name"] == FIRST else [
        f'ActiveWorkbook.Worksheets(1).Name = "{FIRST}"',
        "Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(1))"]
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant", *sheet,
            "ws.Activate", f'ws.Name = "{layout["name"][:31]}"', "On Error Resume Next", "Err.Clear",
            *layout["setup"].splitlines(), 'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"',
            "On Error GoTo 0", 'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_paste_leaves_what_excel_leaves(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]
