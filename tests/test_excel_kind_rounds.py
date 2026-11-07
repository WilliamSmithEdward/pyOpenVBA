"""Held-out rounds over the conditional format kinds, colour scale criteria, icon sets, AutoFilter and its Filters, pivot cells,
their row items, pivot filters, slicers and threaded comments, replayed on the model as it stands (scripts/measure_kind_rounds.py,
tests/fixtures/kind_rounds/). As in tests/test_excel_table_pivot_rounds.py, each round was predicted before live Excel answered, and
USED_UP names the rounds the model was fixed against.

TYPED_SET holds the typed calls Excel refused with 13 because the insert before them moved a rule of another kind under the range,
so the Set into a variable of the asked class fails. The model refuses such a Set once vba-typed-set is merged; until then they
are expected to differ, and strictly, so the marks come off when it lands."""
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

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8')) for path in sorted((ROOT / 'tests/fixtures/kind_rounds').glob('round_*.json'))]
USED_UP = {52101: "an IconSet reached through its condition answered Parent with the workbook's IconSets, not the condition "
                  '(tests/fixtures/icon_set_parent.json)',
           52102: 'the same IconSet Parent'}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]
TYPED_SET = {(seed, at) for seed, at, case in CASES if case['typed'] == 'reach E13' and case['model_typed'] != 'reach E13'}


def _ids(cases: list[tuple[int, int, dict[str, str]]]) -> list[str]:
    return [f'{seed}-{at}-{case["cls"]}' for seed, at, case in cases]


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=_ids(CASES))
def test_each_late_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert KINDS.base.model_answer(case, False, seed, KINDS.KINDS) == case['late']


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=_ids(CASES))
def test_each_typed_call_as_excel_answered(seed: int, at: int, case: dict[str, str], request: pytest.FixtureRequest) -> None:
    if (seed, at) in TYPED_SET:
        request.applymarker(pytest.mark.xfail(strict=True, reason='a Set of another class into a typed variable: vba-typed-set'))
    assert KINDS.base.model_answer(case, True, seed, KINDS.KINDS) == case['typed']


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert KINDS.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
