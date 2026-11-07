"""Held-out rounds over ListObject.QueryTable on a table over a database, replayed on the model as it stands
(scripts/measure_list_query_rounds.py, tests/fixtures/list_query_rounds/). Each round was predicted before live Excel answered, and
USED_UP names the rounds the model was fixed against."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_list_query_rounds', ROOT / 'scripts/measure_list_query_rounds.py')
assert _spec is not None and _spec.loader is not None
ROUNDS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ROUNDS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8')) for path in sorted((ROOT / 'tests/fixtures/list_query_rounds').glob('round_*.json'))]
USED_UP = {54201: "TypeName of a ListObject's query table answered ListQueryTable, where Excel answers QueryTable",
           54202: 'the same TypeName'}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert ROUNDS.model_answer(case, False, seed) == case['late']
    assert ROUNDS.model_answer(case, True, seed) == case['typed']


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert ROUNDS.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
