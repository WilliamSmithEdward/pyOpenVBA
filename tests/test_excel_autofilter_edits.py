"""Edits on a filtered sheet, replayed against the in-memory model.

tests/fixtures/autofilter_edits.json is what scripts/measure_autofilter_edits.py
saw in live Excel: a table filtered to its apples, each layout in a new
workbook, then an edit -- a write, a clear, a fill, a copy, a cut, rows or
columns deleted or inserted -- and a dump of A1:G8 (each cell's formula),
which of rows 1 to 14 are hidden, the filter's range and criteria, the
hidden _FilterDatabase name, what a copy landed, and what the last call
returned or the error it raised.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "autofilter_edits.json")
                                    .read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]

#: Layouts the model refuses, and why.
UNSUPPORTED = {
    **{name: "Excel reads a two-dimensional array for the second of several areas at another stride"
       for name in ("array_distinct", "array_distinct_3col", "array_small", "array_areas_no_filter",
                    "value_self_assign")},
    **{name: "the formula engine has no SUBTOTAL" for name in ("subtotal_filtered", "subtotal_by_hand")},
}

#: Layouts that meet a gap the model documents rather than refuses, and which one.
_CUT = "a cut to another sheet trims a reference it takes the top or bottom of, which the model leaves"
GAPS = {
    "cut_part": _CUT,
    "cut_part_no_filter": _CUT,
    "cut_bottom_no_filter": _CUT,
    "cut_column_part_no_filter": _CUT,
}


def _run(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant, dest As Object",
            "Set ws = ActiveWorkbook.Worksheets(1)", f'ws.Name = "{layout["name"][:28]}"', "On Error Resume Next",
            "Err.Clear", *RECORD["table"].splitlines(), *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


def _cases() -> list[Any]:
    cases: list[Any] = []
    for layout in LAYOUTS:
        marks = [pytest.mark.xfail(reason=GAPS[layout["name"]], strict=True)] if layout["name"] in GAPS else []
        cases.append(pytest.param(layout, id=layout["name"], marks=marks))
    return cases


@pytest.mark.parametrize("layout", _cases())
def test_a_filtered_sheet_takes_edits_as_excel_does(layout: dict[str, Any]) -> None:
    if layout["name"] in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            _run(layout)
        return
    assert _run(layout) == layout["answers"]


def test_every_listed_layout_is_in_the_fixture() -> None:
    names = {layout["name"] for layout in LAYOUTS}
    assert set(UNSUPPORTED) <= names
    assert set(GAPS) <= names
