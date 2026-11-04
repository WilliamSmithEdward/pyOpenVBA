"""Range.ColumnWidth and RowHeight given each kind of value, on a cell, a block, a column and a row, replayed from live Excel
(scripts/measure_range_sizes.py, tests/fixtures/range_sizes.json): the error and whether the size changed. The sizes
themselves follow the display's pixels, and tests/fixtures/dimensions.json holds them at 96 DPI."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_sizes.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
IDS = [f"{number}-{case[0]}" for number, case in CASES]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='Sizes')
    made.run('Sizes.Setup')
    return made


@pytest.mark.parametrize(('number', 'case'), CASES, ids=IDS)
def test_each_value_as_excel_answered(app: ExcelApplication, number: int, case: list[str]) -> None:
    try:
        found = str(app.run(f'Sizes.S{number}'))
    except VBAUnsupportedError:
        pytest.skip('hiding every row of a sheet, a gap Range.RowHeight notes')
    assert found == RECORD['early'].get(str(number), case[3])
