"""PivotTable.ListFormulas, replayed from live Excel (scripts/measure_pivot_list_formulas.py,
tests/fixtures/pivot_list_formulas.json): the sheets in order, the active sheet, and every cell the listing fills
with its formula, number format and boldness, after each case."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/pivot_list_formulas.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ListProbe')
    return str(app.run('ListProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]
