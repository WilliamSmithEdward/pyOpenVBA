"""Held-out rounds over the Name and Names rows, replayed on the model as it stands (scripts/measure_name_rounds.py,
tests/fixtures/name_rounds/). As in tests/test_excel_table_pivot_rounds.py, each round was predicted before live Excel
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
_spec = importlib.util.spec_from_file_location('measure_name_rounds', ROOT / 'scripts/measure_name_rounds.py')
assert _spec is not None and _spec.loader is not None
NAMES: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(NAMES)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8')) for path in sorted((ROOT / 'tests/fixtures/name_rounds').glob('round_*.json'))]
USED_UP = {55101: 'a rename to the same spelling in another case took the case, a rename to text holding ! qualified it, and '
                  'ValidWorkbookParameter was True on a book just opened (tests/fixtures/name_rename.json)'}


def _ansi(call: str) -> bool:
    """Whether the call reached Excel as written: 55101 wrote names beyond ASCII as literals, which Excel's module holds in
    its ANSI code page as ?, so those cases asked Excel another question than the model."""
    try:
        call.encode('cp1252')
    except UnicodeEncodeError:
        return False
    return True


CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases']) if _ansi(case['call'])]


#: What the model answers where Excel took a whole text holding ! as a name's bare part (tests/test_excel_name_rename.py).
REPORTED = 'model VBAUnsupportedError: a name renamed to a text holding ! is not implemented'


def _as_answered(model: str, excel: str) -> bool:
    """The model's answer is Excel's, or Excel's name holds a sheet's name twice and the model reported it."""
    return model == excel or (model == REPORTED and excel.split(' ')[0].count('!') == 2)


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert _as_answered(NAMES.base.model_answer(case, False, seed, NAMES.LISTS), case['late'])
    assert _as_answered(NAMES.base.model_answer(case, True, seed, NAMES.LISTS), case['typed'])


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert NAMES.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}


def test_only_the_first_round_wrote_literals_beyond_ansi() -> None:
    assert {record['seed'] for record in RECORDS if not all(_ansi(case['call']) for case in record['cases'])} <= {55101}
