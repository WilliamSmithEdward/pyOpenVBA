"""A reference spanning every row stays whole as rows go in, and one spanning every column as columns go in: a held ws.Cells and a
formula's A1:XFD1048576 or A1:B1048576 keep their lines, where one that only reaches the sheet's edge moves its near end, on a new
book and after Excel's own save and reopen (tests/fixtures/whole_span_inserts.json, scripts/measure_whole_span_inserts.py).
Held-out round 51106 found the model moving ws.Cells to 2:1048576.

Excel also spells a formula's references again as it takes them, A1:B1048576 as A:B and A1:XFD3 as 1:3; the model keeps them as
written. That is a spelling, the same cells either way, and the formulas here are compared by what they count."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures' / 'whole_span_inserts.json').read_text(encoding='utf-8'))
_MEASURED: list[dict[str, Any]] = RECORD['cases']


def _parts(answer: str) -> tuple[list[str], list[str]]:
    """The held ranges' addresses, and the formula's references counted as rows by columns, for each edit."""
    held: list[str] = []
    counted: list[str] = []
    for part in answer.split(' || '):
        addresses, formula = part.split(' | ')
        held.append(addresses)
        counted.append(';'.join(_size(one) for one in re.findall(r'ROWS\(([^)]*)\)', formula)))
    return held, counted


def _size(reference: str) -> str:
    from pyopenvba._a1 import parse_area

    area = parse_area(reference)
    return f'{area.top}-{area.bottom}x{area.left}-{area.right}'


@pytest.mark.parametrize('case', _MEASURED, ids=lambda case: f'reopened {case["reopen"]}')
def test_each_span_as_excel_moved_it(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Spans')
    assert _parts(str(app.run('Spans.Probe', str(tmp_path / 'w.xlsx'), case['reopen']))) == _parts(case['answer'])


def test_the_whole_sheet_stays_whole_both_ways() -> None:
    assert all(case['answer'].startswith('rows 1:1048576 ') and ' || columns 1:1048576 ' in case['answer'] for case in _MEASURED)
