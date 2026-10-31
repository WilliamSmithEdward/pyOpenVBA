"""Workbook.Colors and ResetColors: the palette ColorIndex counts in, replayed against live Excel
(tests/fixtures/workbook_colors.json, scripts/measure_workbook_colors.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/workbook_colors.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: Cases the model reports itself unable to answer: ColorIndex of a colour a changed palette does not hold (8, 24,
#: 42, 45), an array of fewer than 56 colours (18, which crashed Excel, and 22), and Null (39).
UNSUPPORTED = {8, 18, 22, 24, 39, 42, 45}


def _saved_palette(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    found = re.search(r"<colors>.*?</colors>", styles, re.DOTALL)
    return found.group(0) if found else ""


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_the_palette_answers_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, _read, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    if index in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", str(tmp_path), index, index)
        return
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, statements
    if saved:
        assert _saved_palette(tmp_path / saved) == FIXTURE["parts"][saved]


def test_a_saved_palette_comes_back(tmp_path: Path) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub P()\n    ActiveWorkbook.Colors(3) = RGB(1, 2, 3)\n    Range("A1").Interior.ColorIndex = 3\n'
                   'End Sub\n', name="M")
    app.run("P")
    saved = tmp_path / "book.xlsx"
    book.SaveAs(str(saved))
    again = ExcelApplication.open(saved)
    assert again.evaluate("ActiveWorkbook.Colors(3)") == 197121.0
    assert again.evaluate('Range("A1").Interior.Color') == 197121.0
