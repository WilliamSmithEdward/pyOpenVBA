"""ActiveChart on Application, Workbook and Window, replayed from live Excel (scripts/measure_active_chart.py,
tests/fixtures/active_chart.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/active_chart.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ActiveChartProbe')
    return str(app.run('ActiveChartProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case for case, _ in RECORD['cases']])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
