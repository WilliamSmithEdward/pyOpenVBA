"""Worksheet settings macros set beside Protect or on dashboards, replayed from live Excel
(scripts/measure_worksheet_settings.py, tests/fixtures/worksheet_settings.json)."""

from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "worksheet_settings.json").read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


def answered(text: str) -> list[list[str]]:
    return [entry.split("^") for entry in text.split("~") if entry]


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[list[str]], Path]:
    folder = tmp_path_factory.mktemp("worksheet_settings").resolve()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Settings")
    answer = str(app.run("Settings.Probe"))
    app.add_module(portable(RECORD["saved"]), name="Saved")
    answer += str(app.run("Saved.Saved", str(folder)))
    return answered(answer), folder


def test_each_setting_reads_back_as_excels_does(replayed: tuple[list[list[str]], Path]) -> None:
    found, _folder = replayed
    assert found == RECORD["cases"]


def test_a_save_writes_the_settings_excel_writes(replayed: tuple[list[list[str]], Path]) -> None:
    _found, folder = replayed
    with zipfile.ZipFile(folder / "saved.xlsx") as package:
        for index in (1, 2):
            sheet = package.read(f"xl/worksheets/sheet{index}.xml").decode("utf-8")
            found = re.findall(r"<sheetPr\b[^>]*>|<sheetView\b[^>]*>|<sheetProtection\b[^>]*>", sheet)
            assert found == RECORD["parts"][f"sheet{index}"], index
            # The frozen sheet's formulas as it kept them, its copy's worked out.
            assert re.findall(r'<c r="[BC]1"[^>]*>.*?<v>(.*?)</v>', sheet) == RECORD["parts"][f"values{index}"], index
