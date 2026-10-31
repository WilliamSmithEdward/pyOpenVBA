"""Range.ApplyNames, replayed against live Excel (tests/fixtures/apply_names.json, scripts/measure_apply_names.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads((Path(__file__).parent / "fixtures/apply_names.json").read_text(
    encoding="utf-8"))
#: Two names covering the same cells (10), which the model does not choose between; the probe's own LOW2, a cell's
#: address that Names.Add refused (18); and a name for a constant, after which Excel's automation call failed (19).
UNSUPPORTED = {10}
NOT_REPLAYED = {18, 19}


@pytest.mark.parametrize("index", [one for one in range(len(FIXTURE["cases"])) if one not in NOT_REPLAYED])
def test_apply_names_answers_as_live_excel_did(index: int) -> None:
    _names, _formulas, call, _read, answer = FIXTURE["cases"][index]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(FIXTURE["source"], name="Measure")
    if index in UNSUPPORTED:
        with pytest.raises(VBAUnsupportedError):
            app.run("Measure.Probe", index, index)
        return
    assert app.run("Measure.Probe", index, index) == answer, call


def test_a_defined_name_keeps_its_spelling_though_a_function_has_the_word() -> None:
    """With Rate defined, =RATE*2 reads back =Rate*2 and =SUMX+1 =sumx+1; =Sum*2, naming nothing, reads =SUM*2, and
    a call, =rate(1,2,3), the function, =RATE(1,2,3)."""
    spelling = FIXTURE["names_spelling"]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(spelling["source"], name="Spell")
    assert app.run("Spell.Spelling") == spelling["answer"]
