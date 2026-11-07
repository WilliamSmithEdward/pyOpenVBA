"""Held-out rounds over typing another script's digits into a cell, replayed on the model as it stands
(scripts/measure_typing_native_rounds.py, tests/fixtures/typing_native_rounds/): each round was predicted, its digest
committed, before live Excel answered, and USED_UP names the rounds the model was fixed against."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_typing_native_rounds', ROOT / 'scripts/measure_typing_native_rounds.py')
assert _spec is not None and _spec.loader is not None
ROUNDS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ROUNDS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8'))
                                 for path in sorted((ROOT / 'tests/fixtures/typing_native_rounds').glob('round_*.json'))]
USED_UP = {54101: 'Text showed ASCII digits where Excel shows the script the format names; d-mmm-yy, #,##0.00, cents '
                  'and a long number General shows with fewer places were not yet ruled',
           54102: "a cell's own format was replaced by any native one, where Excel replaces it as the ASCII spelling "
                  "would; a euro amount, a General date past its limits and Tamil's zero"}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_write_as_excel_answered(seed: int, at: int, case: dict[str, Any]) -> None:
    assert ROUNDS.model([case]) == [case['excel']]


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert ROUNDS.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('excel' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
