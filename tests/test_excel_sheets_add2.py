"""Sheets.Add2 and Worksheets.Add2 on a workbook's sheets, replayed from live Excel (scripts/measure_sheets_add2.py,
tests/fixtures/sheets_add2.json): error 445 whatever is given, and no sheet added."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/sheets_add2.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Add2Probe')
    return str(app.run('Add2Probe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['calls'])))
def test_each_call_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['calls'][index][1]
