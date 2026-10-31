"""Range.ID, AllowEdit and ListHeaderRows, replayed against live Excel (tests/fixtures/range_small_members.json,
scripts/measure_range_small_members.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/range_small_members.json").read_text(
    encoding="utf-8"))


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_small_range_members_answer_as_live_excel_did(index: int) -> None:
    statements, read, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    assert app.run("Measure.Probe", index, index) == answer, (statements, read)
