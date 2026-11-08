"""Which way Range.Insert and Range.Delete with no Shift move cells beside sparklines, replayed against Excel's answers
(scripts/measure_sparkline_directions.py), the whole of A1:L20 read after each edit.

A cell holding only a sparkline counts as empty. Where nothing lies either way that decides -- below the range in its
columns, to its right in its rows -- an Insert pushes cells down and a Delete pulls them left.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

FIXTURES = Path(__file__).parent / 'fixtures'
ROUNDS = tuple(stem for stem in ('sparkline_directions', 'sparkline_directions1', 'sparkline_directions2')
               if (FIXTURES / f'{stem}.json').exists())


def _cases() -> list[tuple[str, list[Any], str, str]]:
    found: list[tuple[str, list[Any], str, str]] = []
    for stem in ROUNDS:
        record = json.loads((FIXTURES / f'{stem}.json').read_text('utf-8'))
        found += [(stem, call, answer, record['source']) for call, answer in zip(record['cases'], record['answers'],
                                                                                   strict=True)]
    return found


CASES = _cases()


@pytest.mark.parametrize('stem, call, answer, source', CASES,
                         ids=[f'{one[0]}-{number}' for number, one in enumerate(CASES)])
def test_each_edit_moves_the_cells_as_excel_did(stem: str, call: list[Any], answer: str, source: str) -> None:
    del stem
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    assert str(app.run('Probe.Go', *call)) == answer
