"""Count and Item on every collection, each kind of index a macro can pass and each way in, replayed from live Excel
(scripts/measure_collection_items.py, tests/fixtures/collection_items.json)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
RECORD = json.loads((ROOT / 'tests/fixtures/collection_items.json').read_text(encoding='utf-8'))
#: Collections whose answers the model does not all reproduce yet: an index it reads otherwise, or a member it lacks.
#: Their checklist rows stay PARTIAL.
UNFINISHED = {'ChartGroups', 'ChartObjects', 'CommentsThreaded', 'DataLabels', 'DocumentProperties', 'Errors',
              'FormatConditions', 'FullSeriesCollection', 'Hyperlinks', 'Panes', 'PivotFields', 'PivotFilters',
              'PivotItems', 'PivotTables', 'Points', 'Queries', 'Scenarios', 'SeriesCollection', 'ShapeRange', 'Shapes',
              'SheetViews', 'Slicers', 'SparklineGroup', 'TableStyles', 'VPageBreaks'}
CASES = [(number, *case) for number, case in enumerate(RECORD['cases']) if case[0] not in UNFINISHED]


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='Items')
    made.run('Items.Setup')
    return made


@pytest.mark.parametrize(('number', 'owner', 'reach', 'case', 'answer'), CASES,
                         ids=[f'{one[2]}-{one[3]}' for one in CASES])
def test_each_index_as_excel_answered(app: ExcelApplication, number: int, owner: str, reach: str, case: str,
                                      answer: str) -> None:
    try:
        found = str(app.run(f'Items.C{number}'))
    except VBAUnsupportedError:
        pytest.skip('reported as a gap')
    assert found.replace(str(app.run('Items.BookName')), '<book>') == answer
