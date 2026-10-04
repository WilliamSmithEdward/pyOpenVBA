"""Measured event order across document modules and WithEvents listeners."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/excel_event_sinks.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("probe", MEASURED["probes"], ids=lambda probe: probe["name"])
def test_excel_event_sinks(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    sheet = book.sheets_[0]
    app.add_module("Public Log As String\n", name="EventLog")
    app.add_module(MEASURED["sink"], name="ExcelSink", kind="class")
    app.add_module(MEASURED["sheet"], name=sheet.code_name, kind="document")
    app.add_module(MEASURED["book"], name=book.code_name, kind="document")
    app.add_module(probe["source"], name="ProbeModule")
    assert app.run("Probe") == probe["result"]
