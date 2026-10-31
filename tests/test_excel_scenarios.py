"""Worksheet.Scenarios, replayed against live Excel (tests/fixtures/scenarios.json, scripts/measure_scenarios.py)."""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/scenarios.json").read_text(encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: The day the fixture was measured, which Excel wrote into each comment; the model writes today's.
MEASURED = "10/30/2026"


def _today() -> str:
    day = dt.date.today()
    return f"{day.month}/{day.day}/{day.year}"


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_scenarios_answer_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer.replace(MEASURED, _today()), statements
    if saved in FIXTURE["scenarios"]:
        with zipfile.ZipFile(tmp_path / saved) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        found = (re.search(r"<scenarios\b.*?</scenarios>", xml, re.DOTALL) or [""])[0]
        assert found == FIXTURE["scenarios"][saved].replace(MEASURED, _today()), statements
