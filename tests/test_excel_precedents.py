"""Range.Precedents, DirectPrecedents, Dependents and DirectDependents, replayed from live Excel
(scripts/measure_precedents.py, tests/fixtures/precedents.json)."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD = json.loads((Path(__file__).parent / "fixtures" / "precedents.json").read_text(encoding="utf-8"))


def portable(code: str) -> str:
    """Excel's macro with its paths in this system's separator."""
    return re.sub(r'(folder & ")([^"]*)"', lambda found: found.group(1) + found.group(2).replace("\\", os.sep) + '"',
                  code)


@pytest.fixture(scope="module")
def replayed(tmp_path_factory: pytest.TempPathFactory) -> list[list[str]]:
    folder = tmp_path_factory.mktemp("precedents").resolve()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD["probe"], name="Probe")
    answer = str(app.run("Probe.Probe"))
    app.add_module(portable(RECORD["reopened"]), name="Reopened")
    answer += str(app.run("Reopened.Reopened", str(folder)))
    return [entry.split("^") for entry in answer.split("~") if entry]


@pytest.mark.parametrize("index", range(len(RECORD["cases"])), ids=lambda index: RECORD["cases"][index][0])
def test_each_answer_is_excels(replayed: list[list[str]], index: int) -> None:
    assert replayed[index] == RECORD["cases"][index]
