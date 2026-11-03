"""Range.FlashFill, replayed against what live Excel filled (scripts/measure_flash_fill.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/flash_fill.json').read_text(encoding='utf-8'))
CASES: dict[str, dict[str, Any]] = RECORD['cases']


def _joined(values: list[object] | None) -> str:
    return '|'.join('' if value is None else str(value) for value in values or [])


def _filled(case: dict[str, Any]) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='FillProbe')
    first, last = case['part'] or (2, 0)
    return app.run('FillProbe.Fill', _joined(case['inputs']), _joined(case['seconds']), _joined(case['examples']), first, last)


@pytest.mark.parametrize('name', list(CASES))
def test_as_excel_filled(name: str) -> None:
    assert str(_filled(CASES[name])) == CASES[name]['answer']


def test_a_row_without_the_word_its_rule_asks_for_says_it_is_not_implemented() -> None:
    case = {'inputs': ['Ann May Lee', 'Bob'], 'seconds': None, 'examples': ['May'], 'part': None}
    with pytest.raises(VBAUnsupportedError):
        _filled(case)
