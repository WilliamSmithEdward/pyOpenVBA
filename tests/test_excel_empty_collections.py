"""Collections a new workbook holds nothing in, replayed from live Excel (scripts/measure_empty_collections.py,
tests/fixtures/empty_collections.json), and refused on a loaded workbook that holds something in one."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/empty_collections.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='EmptyProbe')
    return str(app.run('EmptyProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[name for name, _ in RECORD['cases']])
def test_each_collection_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]


def test_a_workbook_holding_query_tables_is_refused_rather_than_read_empty() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    book.sheets_[0].held_kinds = {'queryTable'}
    app.add_module('Public Function Count() As Long\n    Count = ActiveSheet.QueryTables.Count\nEnd Function\n',
                   name='Held')
    with pytest.raises(VBAUnsupportedError, match='QueryTables'):
        app.run('Held.Count')
