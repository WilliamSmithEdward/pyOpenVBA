"""Range.AutoFilter and the AutoFilter object, replayed against the in-memory model.

tests/fixtures/autofilter.json is what scripts/measure_autofilter.py saw in
live Excel: layouts that filter one small table, and a dump of which of
rows 1 to 14 are hidden, the sheet's AutoFilterMode and FilterMode, the
filter's range and each column's criteria, the hidden _FilterDatabase
name, and what the last call returned or the error it raised.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "autofilter.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant",
            "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{layout["name"][:31]}"', "On Error Resume Next",
            "Err.Clear", *RECORD["table"].splitlines(), *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_filter_hides_what_excel_hides(layout: dict[str, Any]) -> None:
    assert _run(layout) == layout["answers"]
