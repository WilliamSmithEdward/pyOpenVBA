"""Held-out rounds over ListObject's settings and reads, replayed on the model as it stands (scripts/measure_list_object_rounds.py,
tests/fixtures/list_object_rounds/). As in tests/test_excel_table_pivot_rounds.py, each round was predicted before live Excel
answered, and USED_UP names the rounds the model was fixed against."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_list_object_rounds', ROOT / 'scripts/measure_list_object_rounds.py')
assert _spec is not None and _spec.loader is not None
LISTS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(LISTS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8')) for path in sorted((ROOT / 'tests/fixtures/list_object_rounds').glob('round_*.json'))]
USED_UP = {53101: "a table's default member answered DisplayName, not Name (tests/fixtures/list_object_default.json)",
           53102: 'the same default member',
           53103: "TableStyle set to an unknown name on a book with a style of its own took the name, where Excel raises 450 "
                  '(tests/fixtures/table_style_own.json)',
           53104: 'the same TableStyle set',
           53105: "a totals table's AutoFilter.Range took in the totals row with the table's sheet not in front "
                  '(tests/fixtures/table_filter_active.json)'}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert LISTS.base.model_answer(case, False, seed, LISTS.LISTS) == case['late']
    assert LISTS.base.model_answer(case, True, seed, LISTS.LISTS) == case['typed']


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert LISTS.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
