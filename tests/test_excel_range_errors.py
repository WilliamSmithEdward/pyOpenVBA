"""Range.Errors, replayed against live Excel (tests/fixtures/range_errors.json, scripts/measure_range_errors.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/range_errors.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: Two checks ignored on one cell, which Excel saves as nothing and the model saves the same way.
CASES = range(len(FIXTURE["cases"]))


@pytest.mark.parametrize("index", CASES)
def test_range_errors_answer_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, cell, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, (statements, cell)
    if saved in FIXTURE["ignored"]:
        with zipfile.ZipFile(tmp_path / saved) as package:
            xml = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
        found = (re.search(r"<ignoredErrors>.*?</ignoredErrors>", xml, re.DOTALL) or [""])[0]
        assert found == FIXTURE["ignored"][saved], statements
