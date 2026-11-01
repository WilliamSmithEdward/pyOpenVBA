"""PivotTable's ConvertToFormulas, ChangeConnection and ChangePivotCache, replayed from live Excel
(scripts/measure_pivot_table_more.py, tests/fixtures/pivot_table_more.json): what each answers, or the
error it raises, the cache count and the table's cells after it."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/pivot_table_more.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='TableMoreProbe')
    return str(app.run('TableMoreProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
