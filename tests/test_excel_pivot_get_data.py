"""PivotTable.GetData replayed from live Excel (scripts/measure_pivot_get_data.py, tests/fixtures/pivot_get_data.json):
names picking data fields and items, quoted and not, in either order and any case, with field[item], Grand Total,
names the table lacks, and items whose cell shows no value."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/pivot_get_data.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='GetDataProbe')
    return str(app.run('GetDataProbe.Probe')).split('~')


@pytest.mark.parametrize('index', range(len(RECORD['names'])), ids=[name for name, _ in RECORD['names']])
def test_each_name_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['names'][index][1]
