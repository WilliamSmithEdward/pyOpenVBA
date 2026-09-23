"""Range.AutoFill, replayed against the in-memory model.

tests/fixtures/autofill.json is what scripts/measure_autofill.py saw in
live Excel. A layout fills a few cells on a sheet of its own and dumps
A1:E8 afterwards: each cell's formula, the type of its value, how far its
number sits from its own 15-digit spelling, its number format and bold,
then the used range and what AutoFill returned or the error it raised. A
column case fills one column from a text, a pair of values or a short
list, and reads the column back.

Some fills Excel makes the model reports it cannot make; each is listed
with the reason, and the replay checks that it says so.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "autofill.json").read_text(encoding="utf-8"))
LAYOUTS: list[dict[str, Any]] = RECORD["layouts"]
COLUMNS: list[dict[str, Any]] = RECORD["columns"]

_TREND = "a trend through three or more uneven numbers follows Excel's LINEST arithmetic"
_GROWTH = "the growth trend's arithmetic is not modelled"
#: Layouts the model reports it cannot fill, and why.
UNSUPPORTED_LAYOUTS: dict[str, str] = {
    "three_numbers": _TREND, "three_numbers_series": _TREND, "three_numbers_linear": _TREND,
    "least_squares": _TREND, "four_points": _TREND, "three_numbers_growth": _GROWTH, "growth_two": _GROWTH,
    "growth_fraction": _GROWTH, "growth_three": _GROWTH, "flash_fill": "Flash Fill is not modelled",
    "diagonal": "a destination running on two ways", "diagonal_series": "a destination running on two ways",
    "diagonal_text": "a destination running on two ways", "diagonal_two": "a destination running on two ways",
    "diagonal_row": "a destination running on two ways", "diagonal_offset": "a destination running on two ways",
    "diagonal_tall": "a destination running on two ways", "diagonal_wide": "a destination running on two ways",
    "date_time_pair": "dates whose times differ",
}


def _label(case: dict[str, Any]) -> str:
    return f"{' / '.join(case['values'])} {case['kind']} {case['format']}".strip()


_TIMES = "dates whose times differ"
_WEEKS = "weekday fills of dates more than a week apart"
#: Column cases the model reports it cannot fill, by label, and why.
UNSUPPORTED_COLUMNS: dict[str, str] = {
    "1 / 3 / 1 / 2": _TREND, "1 / 2 / 3 / 5 / 8": _TREND, "10 / 20 / 31": _TREND, "1 / 2 / 4 xlFillValues": _TREND,
    "1 / 3 xlGrowthTrend": _GROWTH, "2 / 3 xlGrowthTrend": _GROWTH, "1 / 1.1 xlGrowthTrend": _GROWTH,
    "3 / 1 xlGrowthTrend": _GROWTH, "1 / 2 / 4 / 8 xlGrowthTrend": _GROWTH, "1 / 10 xlGrowthTrend": _GROWTH,
    "0.5 / 0.25 xlGrowthTrend": _GROWTH, "1 / 2 / 5 xlGrowthTrend": _GROWTH,
    "#1/1/2020 6:00:00 AM# / #1/1/2020 7:00:00 AM#": _TIMES, "#1/1/2020# / #1/2/2020 1:00:00 AM#": _TIMES,
    "#1/1/1900 6:00:00 AM# / #1/2/1900 7:00:00 AM#": _TIMES,
    "1 / 2  ['h:mm', '[h]:mm']": "times a day or more apart",
    "#1/6/2020# / #1/20/2020# xlFillWeekdays": _WEEKS, "#1/6/2020# / #2/3/2020# xlFillWeekdays": _WEEKS,
}


def _layout(layout: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object, failed As String, v As Variant",
            "Set ws = ActiveWorkbook.Worksheets(1)", "On Error Resume Next", "Err.Clear", *layout["setup"].splitlines(),
            'If Err.Number <> 0 Then failed = "S" & Err.Number & ";"', "On Error GoTo 0",
            'Report = failed & Show(v) & ";" & Dump(ws)', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


def _column(case: dict[str, Any]) -> object:
    code = ["Public Function Report() As String", "Dim ws As Object", "Set ws = ActiveWorkbook.Worksheets(1)",
            "On Error Resume Next", "Err.Clear", *case["setup"].splitlines(),
            f'Report = Err.Number & ":" & Column(ws, {case["column"]}, {case["depth"]})', "End Function"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "\n".join(code) + "\n", name="Probe")
    return app.run("Report")


@pytest.mark.parametrize("layout", LAYOUTS, ids=[layout["name"] for layout in LAYOUTS])
def test_a_fill_leaves_what_excel_leaves(layout: dict[str, Any]) -> None:
    if layout["name"] in UNSUPPORTED_LAYOUTS:
        with pytest.raises(VBAUnsupportedError):
            _layout(layout)
        return
    assert _layout(layout) == layout["answers"]


@pytest.mark.parametrize("case", COLUMNS, ids=[_label(case) for case in COLUMNS])
def test_a_column_fills_as_excel_fills_it(case: dict[str, Any]) -> None:
    if _label(case) in UNSUPPORTED_COLUMNS:
        with pytest.raises(VBAUnsupportedError):
            _column(case)
        return
    assert _column(case) == case["answer"]
