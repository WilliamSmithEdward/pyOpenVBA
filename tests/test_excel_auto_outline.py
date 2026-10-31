"""Range.AutoOutline, replayed against live Excel (tests/fixtures/auto_outline.json, scripts/measure_auto_outline.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/auto_outline.json").read_text(
    encoding="utf-8"))
#: Summaries of summaries, which Excel nests (1, 2), and groups that overlap (20, 21).
UNSUPPORTED = {1, 2, 20, 21}


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_auto_outline_answers_as_live_excel_did(index: int) -> None:
    cells, call, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    if index in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", index, index)
        return
    assert app.run("Measure.Probe", index, index) == answer, (cells, call)
