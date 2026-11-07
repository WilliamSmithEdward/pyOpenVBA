"""Counts, positions and numbers past what Power Query's functions take, as live Excel's Power Query answered
(scripts/measure_m_number_arguments.py). A count or a position is a whole 32-bit integer, or Power Query's Expression.Error; arithmetic past a
Double is infinity or NaN; a date past its range is Power Query's own error. None of these may reach a macro as a Python exception."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import MError, evaluate_query

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_m_number_arguments', ROOT / 'scripts/measure_m_number_arguments.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
ANSWERS: dict[str, str] = json.loads((ROOT / 'tests/fixtures/m_number_arguments.json').read_text(encoding='utf-8'))['answers']
#: Answered here, but not as Power Query did. Each is an answer or an M error, never a Python exception, and each leaves when it is fixed. The
#: first three are M's decimal number formatting, the family tests/test_mlang_number_limits.py names; then the type's name in an error ("List",
#: not "list"), and a .NET format string ("F2") that the engine reads as literal text.
GAPS: frozenset[str] = frozenset({'Number.Round(1.5, -1)', 'Number.Round(1e300, 2)', 'Number.RoundUp(1e300, 2)', 'List.Range({1, 2, 3}, 9)',
                                  'Number.ToText(1e300, "F2")'})
#: List.LastN is not built here: a refusal, not an answer.
UNBUILT: frozenset[str] = frozenset(case for case in MEASURE.CASES if case.startswith('List.LastN('))


def _error(answer: str) -> str | None:
    return answer.partition(': ')[2] if answer.startswith('error ') else None


def _answer(case: str) -> str:
    try:
        return evaluate_query(MEASURE.formula(case))
    except MError as failed:
        return f'[{failed.reason}] {failed.message}'


def test_every_case_was_measured() -> None:
    assert sorted(ANSWERS) == sorted(MEASURE.CASES)


@pytest.mark.parametrize('case', [case for case in MEASURE.CASES if case not in GAPS | UNBUILT], ids=lambda case: case[:50])
def test_it_answers_as_power_query_did(case: str) -> None:
    excel = ANSWERS[case]
    assert _answer(case) == (_error(excel) or excel)


@pytest.mark.parametrize('case', sorted(GAPS), ids=lambda case: case[:50])
def test_the_named_gaps_still_differ_and_raise_nothing_of_pythons(case: str) -> None:
    """Should one of these come to answer as Power Query did, it leaves GAPS."""
    excel = ANSWERS[case]
    assert _answer(case) != (_error(excel) or excel)


@pytest.mark.parametrize('case', sorted(UNBUILT), ids=lambda case: case[:50])
def test_what_is_not_built_is_refused_openly(case: str) -> None:
    with pytest.raises(VBAUnsupportedError):
        evaluate_query(MEASURE.formula(case))
