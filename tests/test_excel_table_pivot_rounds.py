"""Held-out rounds over the VERIFIED rows of tables, pivot tables and the held families, replayed on the model as it stands
(scripts/measure_table_pivot_rounds.py, tests/fixtures/table_pivot_rounds/). Each round's calls were seeded and the model's answers
and their digest written before live Excel answered; the replay asks the model again, so an old round that now differs fails.

USED_UP names the rounds the model was fixed against, which prove nothing about a row any more, and what the fix was. A row goes
VERIFIED only on two clean rounds the model was not fixed against, on the same code."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_table_pivot_rounds', ROOT / 'scripts/measure_table_pivot_rounds.py')
assert _spec is not None and _spec.loader is not None
ROUNDS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ROUNDS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8')) for path in sorted((ROOT / 'tests/fixtures/table_pivot_rounds').glob('round_*.json'))]
USED_UP = {51101: 'ListObjects(text) went by DisplayName, not Name (tests/fixtures/table_name_lookup.json); its Author answers were '
                  'read with the model\'s own user name, not the measuring user\'s'}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}-{case["cls"]}' for seed, at, case in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert ROUNDS.model_answer(case, False, seed) == case['late']
    assert ROUNDS.model_answer(case, True, seed) == case['typed']


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert ROUNDS.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
