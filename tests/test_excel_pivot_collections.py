"""A pivot table's small collections (scripts/measure_pivot_collections.py)."""

from __future__ import annotations

import json
from pathlib import Path

from pyopenvba.apps.excel import ExcelApplication

MEASURED = json.loads((Path(__file__).parent / "fixtures/pivot_collections.json").read_text(encoding="utf-8"))


def test_pivot_collections() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURED["source"], name="Measure")
    assert app.run("Probe") == MEASURED["result"]
