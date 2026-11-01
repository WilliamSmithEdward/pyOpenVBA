"""Window.ActivateNext and Window.ActivatePrevious, replayed from live Excel
(scripts/measure_window_activate_next.py, tests/fixtures/window_activate_next.json): the active workbook after
each step, and the windows they go round."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/window_activate_next.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='WindowProbe')
    return str(app.run('WindowProbe.Probe')).split('@@@')


def test_the_windows_as_excel_counted_them(answers: list[str]) -> None:
    assert answers[0] == RECORD['windows']


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index + 1] == RECORD['cases'][index][1]
