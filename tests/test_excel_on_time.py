"""Application.OnTime as Excel runs it (scripts/measure_on_time.py)."""

from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

MEASURED = json.loads((Path(__file__).parent / "fixtures/on_time.json").read_text(encoding="utf-8"))
READ = "Function ReadLog() As String\nReadLog = Logged()\nEnd Function\n"


def _ready(source: str = "") -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["timers"], name="Timers")
    app.add_module(READ, name="Reader")
    if source:
        app.add_module(source, name="ProbeModule")
    return app


@pytest.mark.parametrize("row", MEASURED["rows"], ids=lambda row: row["name"])
def test_on_time(row: dict[str, str]) -> None:
    app = _ready(MEASURED["probe"].format(action=row["action"]))
    assert app.run("Probe") == row["during"]
    assert app.run("ReadLog") == row["after"]


def test_a_later_time_waits_for_the_clock() -> None:
    app = _ready('Sub Later()\nApplication.OnTime #2026-11-03 12:00:00#, "A"\nEnd Sub\n')
    app.freeze_clock(dt.datetime(2026, 11, 3, 11, 0, 0))
    app.run("Later")
    assert app.run("ReadLog") == ""
    app.freeze_clock(dt.datetime(2026, 11, 3, 12, 0, 0))
    app.run_on_time()
    assert app.run("ReadLog") == "A|"


def test_a_procedure_that_keeps_coming_back_reports_itself() -> None:
    app = _ready('Sub Forever()\nApplication.OnTime Now - 1, "Forever"\nEnd Sub\n')
    with pytest.raises(VBAUnsupportedError, match="for ever"):
        app.run("Forever")


def test_an_argument_the_model_cannot_read_reports_itself() -> None:
    app = _ready("Sub Go()\nApplication.OnTime Now, \"'WithArgument 1 + 1'\"\nEnd Sub\n")
    with pytest.raises(VBAUnsupportedError, match="argument"):
        app.run("Go")
