"""A range's Hyperlinks keep every link they held: taken away, a link is still counted, and Delete takes those still there and then
fails with 1004; a range that held none deletes nothing and raises nothing, on a new book and after Excel's own save and reopen
(tests/fixtures/range_links_delete.json, scripts/measure_range_links_delete.py). Held-out round 51104 found the model refusing a
range with no link."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures' / 'range_links_delete.json').read_text(encoding='utf-8'))
_MEASURED: list[dict[str, Any]] = RECORD['cases']
CASES: list[tuple[int, dict[str, Any]]] = list(enumerate(_MEASURED))


@pytest.mark.parametrize(('at', 'case'), CASES, ids=[f'{case["case"]} reopened {case["reopen"]}' for _, case in CASES])
def test_each_delete_as_excel_answered(at: int, case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Links')
    assert app.run('Links.Probe', str(tmp_path / 'd.xlsx'), at // 2, case['reopen']) == case['answer']


def test_one_gone_link_fails_the_delete_after_the_rest_go() -> None:
    one = [case['answer'] for case in _MEASURED if case['case'] == 'held, one of two taken']
    assert one == ['count 2|delete E1004|left 0'] * 2
