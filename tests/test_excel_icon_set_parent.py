"""An IconSet's Parent is the IconSetCondition it was taken from, and the workbook's IconSets when reached through wb.IconSets, on
new, reopened and Excel-saved books, late bound and typed (scripts/measure_icon_set_parent.py, tests/fixtures/icon_set_parent.json)."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_kind_rounds', ROOT / 'scripts/measure_kind_rounds.py')
assert _spec is not None and _spec.loader is not None
KINDS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(KINDS)

RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/icon_set_parent.json').read_text(encoding='utf-8'))
ASKED: list[dict[str, str]] = RECORD['cases']
CASES = list(enumerate(ASKED))


@pytest.mark.parametrize(('at', 'case'), CASES, ids=[f'{at}-{case["kind"]}' for at, case in CASES])
def test_each_parent_as_excel_answered(at: int, case: dict[str, str]) -> None:
    assert KINDS.base.model_answer(case, False, RECORD['seed'], KINDS.KINDS) == case['late']
    assert KINDS.base.model_answer(case, True, RECORD['seed'], KINDS.KINDS) == case['typed']


def test_the_rule_written_first_is_the_one_excel_follows() -> None:
    assert all(case['late'] == case['typed'] == case['rule_b'] + ' E0' for case in RECORD['cases'] if 'Key(' not in case['call'])
    assert any(case['late'] != case['rule_a'] + ' E0' for case in RECORD['cases'])
