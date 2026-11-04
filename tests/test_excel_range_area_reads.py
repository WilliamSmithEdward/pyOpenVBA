"""What a range of separate areas reads as, replayed from live Excel (scripts/measure_range_area_reads.py,
tests/fixtures/range_area_reads.json): its value and formulas are the first area's, its text what its cells share or Null."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_area_reads.json').read_text(encoding='utf-8'))
CASES = cast('list[list[str]]', RECORD['cases'])


def _answers(source: str) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Reads')
    return str(app.run('Reads.Probe')).split('@@')


@pytest.mark.parametrize(('binding', 'place'), [('late', 2), ('typed', 3)])
def test_each_read_as_excel_answered(binding: str, place: int) -> None:
    found = _answers(RECORD[f'{binding}_source'])
    assert [(case[0], case[1], case[place], one) for case, one in zip(CASES, found, strict=True) if one != case[place]] == []
