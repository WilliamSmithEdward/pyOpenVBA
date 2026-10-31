"""Workbook settings, replayed against live Excel (tests/fixtures/workbook_settings.json,
scripts/measure_workbook_settings.py)."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/workbook_settings.json").read_text(
    encoding="utf-8"))
#: Settings whose place in a saved file was not measured report themselves unsupported when set to a value Excel
#: takes; so does the default slicer or timeline style; and the criteria cases, which Excel refused anyway, use one.
UNSUPPORTED_SETTINGS = ("UpdateRemoteReferences", "ChartDataPointTrack", "AccuracyVersion", "DefaultSlicerStyle",
                        "DefaultTimelineStyle")


def _unsupported(statements: list[str], answer: str) -> bool:
    first = statements[0]
    # Workbook.Password, which HasPassword answers after, is not implemented.
    return first.startswith("wbo.Password") or (
        any(f"wbo.{name} = " in first for name in UNSUPPORTED_SETTINGS) and not answer.startswith("Err "))


def _parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        workbook = package.read("xl/workbook.xml").decode("utf-8")
        styles = package.read("xl/styles.xml").decode("utf-8")
    return {name: (re.search(rf"<{name}\b[^>]*/?>", text) or [""])[0]
            for name, text in (("workbookPr", workbook), ("tableStyles", styles))}


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_workbook_settings_answer_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, _read, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    if _unsupported(statements, answer):
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", str(tmp_path), index, index)
        return
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, statements
    if saved and saved in FIXTURE["parts"]:
        mine, theirs = _parts(tmp_path / saved), FIXTURE["parts"][saved]
        assert mine["tableStyles"] == theirs["tableStyles"]
        live_flags = re.sub(r'\s(?:defaultThemeVersion|date1904)="[^"]*"', "", theirs["workbookPr"])
        model_flags = re.sub(r'\s(?:defaultThemeVersion|date1904|codeName)="[^"]*"', "", mine["workbookPr"])
        assert model_flags == live_flags, (mine["workbookPr"], theirs["workbookPr"])
