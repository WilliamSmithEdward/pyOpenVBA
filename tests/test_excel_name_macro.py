"""A defined name's MacroType, ShortcutKey, Category and WorkbookParameter, replayed against live Excel
(tests/fixtures/name_macro.json, scripts/measure_name_macro.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/name_macro.json").read_text(encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_name_macro_metadata_answers_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    assert app.run("Measure.Probe", str(tmp_path), index, index) == answer, statements
    if saved in FIXTURE["names"]:
        with zipfile.ZipFile(tmp_path / saved) as package:
            workbook = package.read("xl/workbook.xml").decode("utf-8")
        names = (re.search(r"<definedNames>.*?</definedNames>", workbook, re.DOTALL) or [""])[0]
        assert names == FIXTURE["names"][saved], statements
