"""Interior on a chart's elements, replayed from live Excel (scripts/measure_chart_interior.py, tests/fixtures/chart_interior.json): each
element of a plain chart and an AddChart2 one, read and after each set. What the model refuses passes; nothing reads otherwise."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/chart_interior.json').read_text(encoding='utf-8'))


def label(case: list[Any]) -> str:
    return f"{case[0]} {RECORD['elements'][case[1]]} {RECORD['sets'][case[2]] or '(read)'}"


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(RECORD['source'], name='P')
    return made


@pytest.mark.parametrize('case', RECORD['cases'], ids=label)
def test_each_case_reads_as_excel_did(case: list[Any], app: ExcelApplication) -> None:
    kind, element, which, wanted = case
    try:
        got = str(app.run('P.Probe', kind, element, which))
    except VBAUnsupportedError:
        for book in app.application.workbooks_.books[1:]:
            book.vba_call('Close', [False])
        return
    assert dict(zip(['set', *RECORD['reads']], got.split('|'))) == dict(zip(['set', *RECORD['reads']], wanted.split('|')))
