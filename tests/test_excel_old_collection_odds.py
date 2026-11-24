"""DrawingObjects.Vertices, Reshape and Ungroup, typed and late, on a sheet with and without a group, replayed from live Excel
(scripts/measure_old_collection_odds.py, tests/fixtures/old_collection_odds.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_odds.json').read_text(encoding='utf-8'))
#: Measured and not built yet, each with why. Formula on a button links its caption to the cell, writing the drawing and the VML, and is
#: not built either.
NOT_YET = {'Ungroup': 'with no group Excel answers a DrawingObjects of 10 of the 22 items, and with one it ungroups it and answers 424; '
                      'which items, and why, is not worked out'}
#: The grouped rows group a rectangle and an oval loaded from the file first, which the model does not do yet.
CASES = [one for one in RECORD['methods'] if one['which'] not in NOT_YET and not one['grouped']]


def _app(binding: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['methods_source'].replace('{declared}', 'Excel.DrawingObjects' if binding == 'Typed' else 'Object'), name='Probe')
    return app


@pytest.mark.parametrize('case', CASES, ids=[f'{one["which"]} {"grouped" if one["grouped"] else "plain"} {one["binding"]}' for one in CASES])
def test_each_member_as_excel_answered(case: dict[str, Any], tmp_path: Path) -> None:
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    assert str(_app(case['binding']).run('Probe.Row', str(book), case['grouped'], case['which'])) == case['excel']


def test_ungroup_says_it_is_not_built(tmp_path: Path) -> None:
    book = tmp_path / 'in.xlsx'
    book.write_bytes((FIXTURES / 'old_collections.xlsx').read_bytes())
    with pytest.raises(VBAUnsupportedError):
        _app('Late').run('Probe.Row', str(book), False, 'Ungroup')
