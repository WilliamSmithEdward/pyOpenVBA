"""Workbook.DeleteNumberFormat, replayed against live Excel (tests/fixtures/delete_number_format.json,
scripts/measure_delete_number_format.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/delete_number_format.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_delete_number_format_answers_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, _read, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, statements
    if saved:
        with zipfile.ZipFile(tmp_path / saved) as package:
            styles = package.read("xl/styles.xml").decode("utf-8")
        found = re.search(r"<numFmts\b.*?</numFmts>", styles, re.DOTALL)
        assert (found.group(0) if found else "") == FIXTURE["parts"][saved]
