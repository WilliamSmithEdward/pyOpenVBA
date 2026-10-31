"""Range.Consolidate and the sheet's consolidation settings, replayed against live Excel
(tests/fixtures/consolidate.json, scripts/measure_consolidate.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/consolidate.json").read_text(encoding="utf-8"))
#: Another workbook's cells (14), CreateLinks (22), a result over its own source (42) and whole columns (53).
UNSUPPORTED = {14, 22, 42, 53}
#: An A1 reference for a source: Excel keeps a part of it as the sheet's sources, which the model does not.
ERROR_ONLY = {12, 50}


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_consolidate_answers_as_live_excel_did(index: int) -> None:
    statements, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    if index in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", index, index)
        return
    got = str(app.run("Measure.Probe", index, index))
    if index in ERROR_ONLY:
        assert got.split(" | ")[0] == answer.split(" | ")[0], statements
        return
    assert got == answer, statements
