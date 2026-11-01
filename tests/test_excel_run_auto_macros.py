"""Workbook.RunAutoMacros replayed from live Excel (scripts/measure_run_auto_macros.py,
tests/fixtures/run_auto_macros.json): what each step of the sequence runs, and the errors."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/run_auto_macros.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='AutoMacrosProbe')
    return str(app.run('AutoMacrosProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
