"""Held-out rounds over PivotTable's rows, replayed on the model as it stands (scripts/measure_pivot_table_rounds.py,
tests/fixtures/pivot_table_rounds/). As in tests/test_excel_table_pivot_rounds.py, each round was predicted before live Excel
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
_spec = importlib.util.spec_from_file_location('measure_pivot_table_rounds', ROOT / 'scripts/measure_pivot_table_rounds.py')
assert _spec is not None and _spec.loader is not None
PIVOTS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PIVOTS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8'))
                                 for path in sorted((ROOT / 'tests/fixtures/pivot_table_rounds').glob('round_*.json'))]
USED_UP = {57101: 'ChangeConnection given Nothing raised 13 where Excel raises 1004 (tests/fixtures/change_connection.json)',
           57102: 'a row inserted above the table left it in place where Excel moves it, and ListFormulas with another book in '
                  'front closed its book, where Excel cannot (tests/test_excel_pivot_shift_refusal.py, list_formulas_close.json)'}
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]
#: What the model reports where Excel's answer is one it does not copy: a table moved by a row inserted above it, and the
#: book ListFormulas leaves unclosable with another book in front.
REPORTED = {'ws2.Rows(1).Insert': 'model VBAUnsupportedError: Inserting or deleting cells that would move a pivot table',
            'o.ListFormulas': 'model VBAUnsupportedError: PivotTable.ListFormulas with another workbook in front'}


def _as_answered(model: str, excel: str, case: dict[str, str]) -> bool:
    return model == excel or any(call in case['call'] and model.startswith(report) for call, report in REPORTED.items())


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert _as_answered(PIVOTS.base.model_answer(case, False, seed, PIVOTS.LISTS), case['late'], case)
    assert _as_answered(PIVOTS.base.model_answer(case, True, seed, PIVOTS.LISTS), case['typed'], case)


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert PIVOTS.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
