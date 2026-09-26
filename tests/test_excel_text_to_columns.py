"""Range.TextToColumns, replayed against live Excel.

tests/fixtures/text_to_columns.json is what
scripts/measure_text_to_columns.py saw: lines of text down column A of a
fresh sheet split with TextToColumns -- by each delimiter and several,
runs of them, text qualifiers, fields typed as numbers, dates, times and
formulas, FieldInfo's kinds, other separators, a trailing minus, fixed
widths, a destination, and the shapes each date order reads -- and each
cell of A1:T10 read back: its value's type and text, its formula and its
number format, or the error the call raised. The model splits each case
the same way, with the year a date read without one falls in held at the
year the fixture was measured.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication, _typing
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE = Path(__file__).parent / "fixtures" / "text_to_columns.json"
RECORD: dict[str, Any] = json.loads(FIXTURE.read_text(encoding="utf-8"))
CASES: dict[str, dict[str, str]] = RECORD["cases"]

_UNWORKED = "FieldInfo whose pairs do not number the columns 1, 2, 3 in order does something in Excel not worked out"
#: Cases the model reports unsupported rather than answering, and why.
UNSUPPORTED: dict[str, str] = {"field_skip": _UNWORKED, "field_positional": _UNWORKED, "field_order": _UNWORKED}


def _run(name: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["helper"] + "Public Function Report() As String\nDim ws As Object\n"
                   "Application.DisplayAlerts = False\nSet ws = ActiveWorkbook.Worksheets.Add\nOn Error Resume Next\n"
                   + CASES[name]["body"] + "\n"
                   'If Err.Number <> 0 Then Report = "!" & Err.Number Else Report = Dump(ws)\nEnd Function\n',
                   name="Probe")
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr(_typing, "_this_year", lambda: RECORD["year"])
        return str(app.run("Report"))


@pytest.mark.parametrize("name", [name for name in CASES if name not in UNSUPPORTED])
def test_text_to_columns_leaves_what_excel_leaves(name: str) -> None:
    assert _run(name) == CASES[name]["cells"]


@pytest.mark.parametrize("name", list(UNSUPPORTED))
def test_field_info_excel_reads_otherwise_reports_itself(name: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(name)


def test_every_unsupported_case_was_measured() -> None:
    assert set(UNSUPPORTED) <= set(CASES)
