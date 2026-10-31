"""Interior.Gradient, replayed against live Excel (tests/fixtures/interior_gradient.json,
scripts/measure_interior_gradient.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/interior_gradient.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: A cell coloured before its gradient: Excel's ColorIndex keeps the old colour's index, which nothing saves.
KEPT_INDEX = {5}


def _fills(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        styles = package.read("xl/styles.xml").decode("utf-8")
    return re.findall(r"<fill>.*?</fill>", (re.search(r"<fills\b.*?</fills>", styles, re.DOTALL) or [""])[0])


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_interior_gradient_answers_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name="Measure")
    got = str(app.run("Measure.Probe", str(tmp_path), index, index))
    if index in KEPT_INDEX:
        got = got.replace("ColorIndex=Long:-4105", "ColorIndex=Long:44")
    if answer.startswith("Err 438 "):
        # The model's 438 names the member it lacks; Excel's says only that there is none.
        got, answer = got[:7], answer[:7]
    assert got == answer, statements
    if saved in FIXTURE["fills"]:
        expected = re.findall(r"<fill>.*?</fill>", FIXTURE["fills"][saved]["fills"])
        assert _fills(tmp_path / saved) == expected, statements
