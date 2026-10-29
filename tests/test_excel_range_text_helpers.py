"""Range.ListNames, Range.AutoComplete and Range.Parse, replaying what live Excel answered.

Measured by scripts/measure_range_text_helpers.py: the same macro runs here
case by case, each on a new workbook, and reads back the same cells.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/range_text_helpers.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'].replace('Private Function Tried', 'Public Function Tried'), name='Probe')
    return made


@pytest.mark.parametrize('index', range(len(RECORD['cases'])),
                         ids=[case[0][-1][:60] for case in RECORD['cases']])
def test_as_excel_answered(app: ExcelApplication, index: int) -> None:
    want = RECORD['cases'][index][2]
    got = str(app.run('Probe.Tried', index)).replace('\r\n', '\n').replace('\r', '\n')
    assert got == want
