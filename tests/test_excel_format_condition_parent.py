"""A conditional format's Parent is the range its FormatConditions was taken from, not the range it applies to, for every kind of
rule, on a new book and after Excel's own save and reopen (tests/fixtures/format_condition_parent.json,
scripts/measure_format_condition_parent.py). Held-out round 51102 found the model answering AppliesTo."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures' / 'format_condition_parent.json').read_text(encoding='utf-8'))
_MEASURED: list[dict[str, Any]] = RECORD['cases']
CASES: list[tuple[int, dict[str, Any]]] = list(enumerate(_MEASURED))


@pytest.mark.parametrize(('at', 'case'), CASES, ids=[f'{case["how"]} reopened {case["reopen"]}' for _, case in CASES])
def test_each_parent_as_excel_answered(at: int, case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='Parents')
    assert app.run('Parents.Probe', str(tmp_path / 'f.xlsx'), at // 2, case['reopen']) == case['answer']


def test_a_rule_reached_through_another_range_names_that_range() -> None:
    through = {case['how']: case['answer'].split('|')[0] for case in _MEASURED if not case['reopen']}
    assert through['ws.Range("D3").FormatConditions(1)'] == 'Range:D3'
    assert through['ws.Cells.FormatConditions(1)'] == 'Range:1:1048576'
