"""Text's format codes letter by letter, and the arguments the engine read otherwise, as worksheet formulas and through
WorksheetFunction, replayed from live Excel (scripts/measure_worksheet_function_edges.py,
tests/fixtures/worksheet_function_edges.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/worksheet_function_edges.json').read_text(encoding='utf-8'))
#: ZTEST's far tail agrees to 13 digits; erfc is not Excel's series.
KNOWN = {'=ZTEST({3,6,7,8,6,5,4,2,1,9},-1,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},0,1)', '=ZTEST({3,6,7,8,6,5,4,2,1,9},1,1)'}


@pytest.fixture(scope='module')
def found() -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Edges')
    return str(app.run('Edges.Probe')).split('@@')


@pytest.mark.parametrize(('number', 'case'), list(enumerate(RECORD['cases'])), ids=[case[0] for case in RECORD['cases']])
def test_each_as_excel_answered(found: list[str], number: int, case: list[str]) -> None:
    if case[0] in KNOWN:
        assert found[number] != case[1] and found[number][:20] == case[1][:20]
    else:
        assert found[number] == case[1]
