"""Inserts and deletes beside a chart whose points carry a look or a label and whose error bars read cells, replayed
against live Excel (scripts/measure_chart_shift_losses.py): what the chart reads afterwards, and each series'
references, cached points, numbered points and error bars in the part saved.

A chart made in the session moves its points' looks and labels with their cells. A chart read from a file keeps the
numbers its part gives them, and a reference whose cells all go reads #REF! in its part with its cached points kept,
while the series reads it as it was. An error bar whose cells go reads #REF!, cached as one point of 1 in a chart
made in the session.
"""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('chart_shift_losses', 'chart_shift_losses1', 'chart_shift_losses2')
               if (FIXTURES / f'{stem}.json').exists())


def _script() -> ModuleType:
    path = Path(__file__).parents[1] / 'scripts/measure_chart_shift_losses.py'
    spec = importlib.util.spec_from_file_location('measure_chart_shift_losses', path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


SCRIPT = _script()
CASES: list[tuple[str, list[Any]]] = [(stem, case) for stem in ROUNDS
                                      for case in json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))['cases']]


@pytest.mark.parametrize('stem, case', CASES, ids=[f'{stem}-{case[0]}-{case[1]}' for stem, case in CASES])
def test_each_edit_leaves_the_chart_as_excel_did(stem: str, case: list[Any], tmp_path: Path) -> None:
    del stem
    way, edit, answer, part = case
    assert SCRIPT.model_answer([way, edit], tmp_path) == (answer, part)
