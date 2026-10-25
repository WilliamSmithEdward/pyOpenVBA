"""Application.ConvertFormula, replayed from live Excel (scripts/measure_convert_formula.py,
tests/fixtures/convert_formula.json): formulas converted each way, the kinds of value each argument takes, names
spelled as Excel first saw them, and the longest formula of each shape it still converts."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "convert_formula.json").read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> dict[str, str]:
    folder = tmp_path_factory.mktemp("convert_formula").resolve()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(portable(RECORD["probe"]), name="Probe")
    answer = str(app.run("Probe.Probe", str(folder)))
    return dict(entry.split("^", 1) for entry in answer.split("~") if entry)


@pytest.mark.parametrize("index", range(len(RECORD["cases"])), ids=lambda index: RECORD["cases"][index][0])
def test_each_answer_is_excels(replayed: dict[str, str], index: int) -> None:
    label, *_, answer = RECORD["cases"][index]
    assert replayed[label] == answer
