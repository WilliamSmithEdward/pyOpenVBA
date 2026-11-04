"""Members the checklist had as UNASSESSED that macros reach most, replayed as Excel answered them
(scripts/measure_unassessed_core.py)."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/unassessed_core.json").read_text(encoding="utf-8"))


@pytest.mark.parametrize("row", MEASURED["rows"], ids=lambda row: row["name"])
def test_unassessed_core(row: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["probe"].format(action=row["action"]), name="Measure")
    assert app.run("Probe") == row["result"]
