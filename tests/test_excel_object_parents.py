"""Parent and Application of Areas, Filters, AutoFilter, WorksheetFunction and the table collections, and
ListColumn.Total, replayed from live Excel (scripts/measure_object_parents.py, tests/fixtures/object_parents.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/object_parents.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ParentsProbe')
    return str(app.run('ParentsProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['reads'])))
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['reads'][index][1]
