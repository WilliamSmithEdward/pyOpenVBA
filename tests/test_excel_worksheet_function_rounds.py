"""Held-out rounds over WorksheetFunction's PARTIAL rows, replayed on the model as it stands
(scripts/measure_worksheet_function_rounds.py, tests/fixtures/worksheet_function_rounds/). Each round varies calls Excel
answered before, by its seed, and was predicted before live Excel answered, typed and late bound. USED_UP names the rounds
the model was fixed against, which prove nothing about a row any more, and what the fix was. A row goes VERIFIED only on two
clean rounds the model was not fixed against, on the same code."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_worksheet_function_rounds', ROOT / 'scripts/measure_worksheet_function_rounds.py')
assert _spec is not None and _spec.loader is not None
ROUNDS: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ROUNDS)

RECORDS: list[dict[str, Any]] = [json.loads(path.read_text(encoding='utf-8'))
                                 for path in sorted((ROOT / 'tests/fixtures/worksheet_function_rounds').glob('round_*.json'))]
USED_UP: dict[int, str] = {
    60101: 'eleven functions answered edge cases otherwise, each rule split from its rivals live '
           '(tests/fixtures/worksheet_function_rivals.json): AmorDegrc, AverageIfs, BahtText, DGet, DStDevP, Frequency, IfNa, '
           'Lcm, Transpose, Trend and ZTest; Confidence_T, F_Dist_RT and Gamma_Dist missed in the last digit and left the rounds',
    60102: 'DProduct passed over a constant under a criteria label that names no field, and LinEst lifted its statistics '
           'over a range (tests/fixtures/worksheet_function_rivals.json), now refused; ChiInv and MultiNomial missed by a few '
           'units in the last digits and left the rounds; Unichar(True) read a space through the harness',
    60103: 'Db took a month of 3.5 as it was, Transpose(Empty) was Empty, and VLookup raised 1004 for a cell of text as its '
           'range_lookup (tests/fixtures/worksheet_function_rivals.json)',
    60104: 'the database functions refused a logical field, which Excel reads as its number, and ValueToText took a format '
           'past 0 and 1 (tests/fixtures/worksheet_function_rivals.json); Confidence_Norm, Growth and LinEst missed in the '
           'last digit and left the rounds',
}
#: The clean rounds, 60105 and 60106, ran on the same code, and promoted each row with no difference in either. These
#: functions differ in them, and their rows stay PARTIAL.
_LAST = 'the model answers a unit away in the fifteenth significant digit'
KNOWN: dict[tuple[int, str], str] = {
    (60105, 'Confidence'): _LAST,
    (60105, 'ErfC_Precise'): _LAST,
    (60105, 'RandArray'): 'its numbers are random draws, which no prediction can match',
    (60106, 'Index'): "Index(range, Null) late bound is error 1004 in Excel, where the model takes Null as a blank row and "
                      "answers the whole column",
    (60106, 'TInv'): _LAST,
    (60106, 'ZTest'): _LAST + ', far out in the tail',
}


def _as_measured(seed: int, answer: str) -> str:
    """The model's answer as a round's live run read it: before ESCAPES_FROM the harness handed each character below a
    space back as a space."""
    return answer if seed >= ROUNDS.ESCAPES_FROM else ''.join(' ' if ord(one) < 32 else one for one in answer)


@pytest.mark.parametrize('typed', [False, True], ids=['late', 'typed'])
@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_call_as_excel_answered(record: dict[str, Any], typed: bool) -> None:
    # The functions a round leaves out are left out here too: those the model answers a unit away in the last digit.
    # A call the model refuses is no answer of its own, and its row stays PARTIAL.
    cases = [case for case in record['cases'] if case['function'] not in ROUNDS.LAST_DIGIT | ROUNDS.REFUSING]
    tag = 'typed' if typed else 'late'
    answers = ROUNDS.model_answers(cases, typed, record['seed'] >= ROUNDS.ESCAPES_FROM)
    wrong = [(case['function'], case['setup'], case['arguments'], case[tag], answer)
             for case, answer in zip(cases, answers, strict=True)
             if not ROUNDS.refused(answer) and _as_measured(record['seed'], answer) != case[tag]]
    assert [one for one in wrong if (record['seed'], one[0]) not in KNOWN] == []


def test_the_known_differences_still_differ() -> None:
    for record in RECORDS:
        cases = [case for case in record['cases'] if (record['seed'], case['function']) in KNOWN]
        answers = ROUNDS.model_answers(cases, False, record['seed'] >= ROUNDS.ESCAPES_FROM)
        found = {case['function'] for case, answer in zip(cases, answers, strict=True)
                 if _as_measured(record['seed'], answer) != case['late'] or case['typed'] != case['model_typed']}
        assert found == {function for seed, function in KNOWN if seed == record['seed']}


@pytest.mark.parametrize('record', RECORDS, ids=lambda record: str(record['seed']))
def test_each_round_was_predicted_before_it_was_measured(record: dict[str, Any]) -> None:
    assert ROUNDS.digest(record['cases']) == record['predictions_sha256_16']
    assert record['predicted_at'] < record['measured_at']
    assert all('late' in case and 'typed' in case for case in record['cases'])


def test_the_used_up_rounds_are_rounds() -> None:
    assert set(USED_UP) <= {record['seed'] for record in RECORDS}


def test_each_round_asks_every_function_it_does_not_leave_out() -> None:
    for record in RECORDS:
        asked = {case['function'] for case in record['cases']}
        assert [name for name in ROUNDS.partial_functions() if name not in asked and ROUNDS.bases().get(name.lower())] == []
