"""Application and Parent of more collections, WorkbookQuery's default member, DisplayFormat's MergeCells and Style,
a colour scale FormatColor's ColorIndex and PivotTables.Add, replayed from live Excel
(scripts/measure_collection_members.py, tests/fixtures/collection_members.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'collection_members.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='MembersProbe')
    return str(app.run('MembersProbe.Probe', str(FIXTURES / RECORD['queries_file']))).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['reads'])))
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['reads'][index][1]
