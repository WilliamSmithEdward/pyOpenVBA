"""Workbook.Connections replayed from live Excel (scripts/measure_connections.py, tests/fixtures/connections.json):
a new workbook's count, and a Power Query workbook's OLE DB connection as the connection and its OLEDBConnection
read it."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'connections.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='ConnectionsProbe')
    return str(app.run('ConnectionsProbe.Probe', str(FIXTURES / RECORD['file']))).split('~')


def test_a_new_workbook_has_none(answers: list[str]) -> None:
    assert answers[0] == RECORD['new_count']


@pytest.mark.parametrize('index', range(len(RECORD['reads'])), ids=[read for read, _ in RECORD['reads']])
def test_each_read_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index + 1] == RECORD['reads'][index][1]
