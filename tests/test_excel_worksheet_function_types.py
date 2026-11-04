"""Each worksheet function's parameters the type library declares other than Variant, given cells, several cells, an array,
Null and text, typed and late bound, replayed from live Excel (scripts/measure_worksheet_function_types.py,
tests/fixtures/worksheet_function_types.json).

"known" in the fixture lists the calls the model still answers otherwise, most by the last digits of a statistical
function. Each must go on differing, so one that comes right is taken off the list; those functions' rows stay PARTIAL.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any, cast

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/worksheet_function_types.json').read_text(encoding='utf-8'))
CASES = list(enumerate(cast('list[list[str]]', RECORD['cases'])))
KNOWN: dict[str, list[int]] = RECORD['known']


def _answers(source: str) -> list[str]:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Functions')
    return str(app.run('Functions.Probe')).split('@@')


@pytest.mark.parametrize(('binding', 'place'), [('late', 3), ('typed', 4)])
def test_each_call_as_excel_answered(binding: str, place: int) -> None:
    found = _answers(RECORD[f'{binding}_source'])
    assert len(found) == len(CASES)
    differing = [number for number, case in CASES if found[number] != case[place]]
    assert differing == KNOWN[binding], [CASES[number][1][2] for number in set(differing) ^ set(KNOWN[binding])][:20]
