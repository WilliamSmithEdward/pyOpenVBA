"""Collections taken from a range and held while rows or columns go in: a range's Hyperlinks are the links it held, where they now
stand, and a rule reached through D3 has the moved D3 as Parent, gone once its row is deleted, on a new book and after Excel's own
save and reopen (tests/fixtures/held_range_collections.json, scripts/measure_held_range_collections.py). Held-out round 51103
found the model's held Hyperlinks keeping their old place."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures' / 'held_range_collections.json').read_text(encoding='utf-8'))
_MEASURED: list[dict[str, Any]] = RECORD['cases']
CASES: list[tuple[int, dict[str, Any]]] = list(enumerate(_MEASURED))


@pytest.mark.parametrize(('at', 'case'), CASES, ids=[f'{case["case"]} reopened {case["reopen"]}' for _, case in CASES])
def test_each_held_collection_as_excel_answered(at: int, case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Held')
    assert app.run('Held.Probe', str(tmp_path / 'h.xlsx'), at // 2, case['reopen']) == case['answer']


def test_a_link_added_after_is_not_among_the_held_ones() -> None:
    added = [case['answer'] for case in _MEASURED if case['case'] == 'links held, link added after']
    assert added == ['1/Q4|E0', '1/Q4|E0']
