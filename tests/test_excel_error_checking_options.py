"""Application.ErrorCheckingOptions, replayed against live Excel (tests/fixtures/error_checking_options.json,
scripts/measure_error_checking_options.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/error_checking_options.json").read_text(
    encoding="utf-8"))


def test_a_new_session_reads_the_options_live_excel_read() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    assert app.run("Measure.Probe") == FIXTURE["answer"]


def test_an_option_set_is_kept_for_the_session() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module("Public Function P() As String\n"
                   "    Application.ErrorCheckingOptions.NumberAsText = False\n"
                   "    Application.ErrorCheckingOptions.IndicatorColorIndex = 3\n"
                   "    P = Application.ErrorCheckingOptions.NumberAsText & "
                   "Application.ErrorCheckingOptions.IndicatorColorIndex\n"
                   "End Function\n", name="Kept")
    assert app.run("Kept.P") == "False3"
