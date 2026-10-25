"""Application settings everyday macros switch, replayed from live Excel (scripts/measure_application_settings.py,
tests/fixtures/application_settings.json).

Excel's macro runs in the model as Excel ran it, with AskToUpdateLinks first turned off as the measuring harness
turns it off when it starts Excel; Excel ships it on, which the model starts with.
"""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "application_settings.json").read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


def answered(text: str) -> list[list[str]]:
    return [entry.split("^") for entry in text.split("~") if entry]


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[list[str]], Path]:
    folder = tmp_path_factory.mktemp("application_settings").resolve()
    app = ExcelApplication()
    app.add_workbook()
    assert app.evaluate("Application.AskToUpdateLinks") is True
    app.add_module("Public Sub AsTheHarnessStarts()\n    Application.AskToUpdateLinks = False\nEnd Sub\n", name="Start")
    app.run("Start.AsTheHarnessStarts")
    app.add_module(RECORD["probe"], name="Settings")
    answer = str(app.run("Settings.Probe"))
    breaks = RECORD["probe"].split("Public Function Probe")[0] + portable(RECORD["breaks"])
    app.add_module(breaks, name="Breaks")
    answer += str(app.run("Breaks.Breaks", str(folder)))
    return answered(answer), folder


def test_each_setting_reads_back_as_excels_does(replayed: tuple[list[list[str]], Path]) -> None:
    found, _folder = replayed
    assert found == RECORD["cases"]


def test_display_page_breaks_is_not_saved(replayed: tuple[list[list[str]], Path]) -> None:
    _found, folder = replayed
    with zipfile.ZipFile(folder / "breaks.xlsx") as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    assert re.findall(r"<sheetView\b[^>]*>", sheet) == RECORD["sheet_views"]


def test_a_save_while_print_communication_is_off_writes_what_was_applied(
        replayed: tuple[list[list[str]], Path]) -> None:
    _found, folder = replayed
    with zipfile.ZipFile(folder / "queued.xlsx") as package:
        sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
    # Excel's file says portrait where the model writes no pageSetup at all, which reads as portrait too.
    assert RECORD["queued_page_setup"] == ['<pageSetup orientation="portrait" r:id="rId1"/>']
    assert 'orientation="landscape"' not in sheet
