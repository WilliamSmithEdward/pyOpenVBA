"""When Application.AfterCalculate is raised against SheetChange and SheetCalculate, replayed against the in-memory model.

tests/fixtures/after_calculate.json holds what scripts/measure_after_calculate.py saw in live Excel: the events an application's
handlers logged as a macro calculated, edited cells and wrote formulas on a new workbook, read at the macro's end and again by a
macro run after it.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "after_calculate.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("case", RECORD["cases"], ids=lambda case: case["name"])
def test_after_calculate_is_raised_where_excel_raised_it(case: dict[str, Any]) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD["log"], name="EventLog")
    app.add_module(RECORD["sink"], name="AfterSink", kind="class")
    app.add_module(RECORD["watch"], name="Watcher")
    assert app.run("Watch", case["how"]) == case["inside"]
    assert app.run("Report") == case["after"]
