"""The order Excel raises Calculate over several sheets, the numbers it saves the sheets with, and the calculation
that ends a pivot refresh, replayed against the in-memory model.

tests/fixtures/calculate_order.json holds what scripts/measure_calculate_order.py
saw in live Excel: the events an application's handlers logged as sheets made
out of their tabs' order were calculated, by Application.Calculate, an edit,
Worksheet.Calculate and pivot refreshes, with and without held notifications;
and the name and sheetId each sheet of a saved workbook is listed with, after
sheets were added before others, deleted, copied or added to a file opened
again.
"""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures" / "calculate_order.json")
                                    .read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """The macro Excel ran names its files under a Windows folder; here they go under this system's."""
    return code.replace('folder & "\\" & name', f'folder & "{os.sep}" & name')


@pytest.mark.parametrize("case", RECORD["cases"], ids=lambda case: case["name"])
def test_sheets_raise_calculate_in_the_order_they_were_made(case: dict[str, Any], tmp_path: Path) -> None:
    """Whatever the tabs' order, and after a save and an open; a refresh ends with a calculation that works out
    formulas reading the table and volatile ones and delivers held notifications, once for RefreshAll, and only in
    automatic calculation."""
    app = ExcelApplication()
    app.add_module("Public Log As String\n", name="EventLog")
    app.add_module(RECORD["sink"], name="CalculateSink", kind="class")
    app.add_module(portable(RECORD["watch"]), name="Watcher")
    assert app.run("Watch", case["how"], str(tmp_path)) == case["log"]


@pytest.mark.parametrize("case", RECORD["numbers"], ids=lambda case: case["name"])
def test_a_saved_sheet_keeps_the_number_it_was_made_with(case: dict[str, Any], tmp_path: Path) -> None:
    """A workbook counts the sheets made in it; one deleted gives its number back to none, and a file opened again
    goes on from its highest."""
    app = ExcelApplication()
    app.add_module(portable(RECORD["make"]), name="Maker")
    assert app.run("Make", case["how"], str(tmp_path)) == "ok"
    with zipfile.ZipFile(tmp_path / "numbers.xlsx") as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
    listed = re.findall(r'<sheet\b[^>]*\bname="([^"]*)"[^>]*\bsheetId="(\d+)"', workbook)
    assert [[name, int(number)] for name, number in listed] == case["sheets"]
