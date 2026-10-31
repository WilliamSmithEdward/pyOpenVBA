"""Workbook.PrecisionAsDisplayed, replayed against live Excel (tests/fixtures/precision_as_displayed.json,
scripts/measure_precision_as_displayed.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.formula._display import precision_shown

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/precision_as_displayed.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_precision_as_displayed_answers_as_live_excel_did(index: int, tmp_path: Path) -> None:
    value, number_format, statements, _read, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, (value, number_format, statements)
    for name, calc in FIXTURE["calc"].items():
        if (tmp_path / name).exists():
            with zipfile.ZipFile(tmp_path / name) as package:
                workbook = package.read("xl/workbook.xml").decode("utf-8")
            assert (re.search(r"<calcPr\b[^>]*/?>", workbook) or [""])[0] == calc, name


def test_a_format_that_shows_a_number_whole_keeps_it() -> None:
    assert precision_shown(1.23456, "General") is None
    assert precision_shown(45293.75, "m/d/yyyy") is None
    assert precision_shown(-0.004, "0.00") == 0.0
