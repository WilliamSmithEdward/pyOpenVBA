"""SUBTOTAL, replayed against the in-memory model.

tests/fixtures/subtotal.json is what scripts/measure_subtotal.py saw in live
Excel: a table of names and numbers on two sheets, rows hidden or filtered
one way or another, then SUBTOTAL formulas -- every function number, bad
ones, other kinds of argument -- and each one's value, or the error
writing it raised.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "subtotal.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]

#: Layouts the model refuses, and why: it keeps no reference through a function's answer or an operator.
UNSUPPORTED = {"reference_arguments": "a reference that OFFSET, INDIRECT, IF, INDEX or a reference operator gives"}


def _run(layout: dict[str, Any]) -> object:
    items = ", ".join('"' + formula.replace('"', '""') + '"' for formula in layout["formulas"])
    code = ["Public Function Report() As String", "Dim ws As Object, other As Object, failed As String",
            "Dim r As Long, sheet As Variant", "Set ws = ActiveWorkbook.Worksheets(1)", 'ws.Name = "Data"',
            "Set other = ActiveWorkbook.Worksheets.Add(After:=ws)", 'other.Name = "Other"', "ws.Activate",
            *RECORD["table"].splitlines(), "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            f"Report = failed & Place(ws, Array({items}))", "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_subtotal_works_out_what_excel_does(layout: dict[str, Any]) -> None:
    if layout["name"] in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            _run(layout)
        return
    assert _run(layout) == layout["answers"]
