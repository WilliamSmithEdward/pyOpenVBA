"""Workbook properties of the file and its sharing, replayed against live Excel
(tests/fixtures/workbook_properties.json, scripts/measure_workbook_properties.py)."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/workbook_properties.json").read_text(
    encoding="utf-8"))
SOURCE = FIXTURE["source"].replace('folder & "\\', f'folder & "{os.sep}')
#: UserStatus, which the model does not answer, is left out of each answer.
LEFT_OUT = ("UserStatus=",)
#: EnvelopeVisible, whose set hung Excel; WritePassword and Password, which protect the file.
UNSUPPORTED = {"wbo.EnvelopeVisible = True", 'wbo.WritePassword = "w"', 'wbo.Password = "p"'}


def _kept(answer: str) -> str:
    return "|".join(one for one in answer.split("|") if not one.startswith(LEFT_OUT))


@pytest.mark.parametrize("index", range(len(FIXTURE["cases"])))
def test_workbook_properties_answer_as_live_excel_did(index: int, tmp_path: Path) -> None:
    statements, _saved, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE.replace('Prop(wbo, "UserStatus")', '"x"'), name="Measure")
    if statements and statements[-1] in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", str(tmp_path), index, index)
        return
    assert _kept(str(app.run("Measure.Probe", str(tmp_path), index, index))) == _kept(answer), statements
