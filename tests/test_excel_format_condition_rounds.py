"""Held-out rounds over FormatCondition's rows, replayed on the model as it stands (scripts/measure_format_condition_rounds.py,
tests/fixtures/format_condition_rounds/). As in tests/test_excel_table_pivot_rounds.py, each round was predicted before live
Excel answered, and USED_UP names the rounds the model was fixed against."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_format_condition_rounds', ROOT / 'scripts/measure_format_condition_rounds.py')
assert _spec is not None and _spec.loader is not None
RULES: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(RULES)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8'))
                                 for path in sorted((ROOT / 'tests/fixtures/format_condition_rounds').glob('round_*.json'))]
USED_UP = {56101: "a rule's Font.Size set was refused where Excel raises 1004, a refused NumberFormat left the old code where "
                  "Excel leaves none, and Bold set on a rule reached through the sheet's cells took "
                  '(tests/fixtures/format_condition_font.json)',
           56102: "a font set on a rule held while rows went in above it was reported as reached through other cells, and a "
                  "refused NumberFormat took away a format the file gave the rule (tests/fixtures/format_condition_font.json)"}
#: The refusals the model gives where Excel's answer is one it does not copy (tests/test_excel_format_condition_font.py).
REPORTED = ('model VBAUnsupportedError: Font.Bold set on a rule reached through other cells',
            'model VBAUnsupportedError: Font.Italic set on a rule reached through other cells',
            'model VBAUnsupportedError: Font.Strikethrough set on a rule reached through other cells',
            'model VBAUnsupportedError: Font.Underline set on a rule reached through other cells')
CASES = [(record['seed'], at, case) for record in RECORDS for at, case in enumerate(record['cases'])]


def _as_answered(model: str, excel: str, case: dict[str, str]) -> bool:
    """The model's answer is Excel's, or it reports a font set on a rule reached through cells not its own: the sheet's,
    or the rules' cells once rows went in above them before they were reached."""
    moved = case['held'] == 'False' and case['edit'].startswith('wsL.')
    return model == excel or (model.startswith(REPORTED) and (moved or not case['reach'].startswith('wsL.Range(fcells)')))


@pytest.mark.parametrize(('seed', 'at', 'case'), CASES, ids=[f'{seed}-{at}' for seed, at, _ in CASES])
def test_each_call_as_excel_answered(seed: int, at: int, case: dict[str, str]) -> None:
    assert _as_answered(RULES.base.model_answer(case, False, seed, RULES.LISTS), case['late'], case)
    assert _as_answered(RULES.base.model_answer(case, True, seed, RULES.LISTS), case['typed'], case)


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert RULES.base.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}
