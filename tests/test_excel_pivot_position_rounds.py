"""Held-out rounds over PivotTable's rows that answer cells, with rows and columns going in and out on the table's sheet, replayed
on the model as it stands (scripts/measure_pivot_position_rounds.py, tests/fixtures/pivot_position_rounds/). As in
tests/test_excel_table_pivot_rounds.py, each round was predicted before live Excel answered, and USED_UP names the rounds the
model was fixed against."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_pivot_position_rounds', ROOT / 'scripts/measure_pivot_position_rounds.py')
assert _spec is not None and _spec.loader is not None
POSITIONS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(POSITIONS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8'))
                                 for path in sorted((ROOT / 'tests/fixtures/pivot_position_rounds').glob('round_*.json'))]
USED_UP: dict[int, str] = {}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert POSITIONS.base.model_answer(case, False, seed, POSITIONS.LISTS) == case['late']
    assert POSITIONS.base.model_answer(case, True, seed, POSITIONS.LISTS) == case['typed']


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert POSITIONS.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}


def test_every_round_moves_the_table() -> None:
    """A row or column goes in or out before the table on its own sheet in each round, which is what these rows wait for."""
    assert all(any(case['edit'] in ('ws2.Rows(1).Insert', 'ws2.Columns(1).Insert', 'ws2.Rows("1:2").Insert', 'ws2.Rows(1).Delete')
                   for case in record['cases']) for record in RECORDS)
