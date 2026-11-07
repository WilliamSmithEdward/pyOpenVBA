"""Lists and tables of a vast size, as live Excel's Power Query answered (scripts/measure_m_vast_lists.py). A list range or Table.Repeat past a
32-bit count is Power Query's Expression.Error. Within it, Power Query builds a vast list lazily, which this engine does not: past MOST_ITEMS it
refuses before any work, where it would otherwise run on, as Excel itself does for List.Count(List.Numbers(1, 1e12))."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import MError, evaluate_query

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_m_vast_lists', ROOT / 'scripts/measure_m_vast_lists.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
ANSWERS: dict[str, str] = json.loads((ROOT / 'tests/fixtures/m_vast_lists.json').read_text(encoding='utf-8'))['answers']
#: Power Query answered these by building lazily, or ran on (List.Count of a trillion); this engine refuses them before any work.
VAST = frozenset({'List.First(List.Numbers(1, 1e12))', 'List.First(List.Numbers(1, 3e9))', 'List.First({1..2147483647})', 'List.First(List.Numbers(1, 1.8e308))',
                  'List.Count(List.Numbers(1, 1e12))'})


def _excel(case: str) -> str:
    return ANSWERS[case].rpartition(' | ')[0] if ' | ' in ANSWERS[case] else ANSWERS[case]


@pytest.mark.parametrize('case', [case for case in ANSWERS if _excel(case).startswith('error')], ids=lambda case: case[:50])
def test_power_querys_error_is_the_same(case: str) -> None:
    with pytest.raises(MError) as raised:
        evaluate_query(MEASURE.formula(case))
    assert _excel(case).endswith(f'[{raised.value.reason}] {raised.value.message}')


@pytest.mark.parametrize('case', [case for case in ANSWERS if not _excel(case).startswith('error') and case not in VAST], ids=lambda case: case[:50])
def test_what_fits_is_built_and_answers_the_same(case: str) -> None:
    assert evaluate_query(MEASURE.formula(case)) == _excel(case)


@pytest.mark.parametrize('case', sorted(VAST), ids=lambda case: case[:50])
def test_what_power_query_builds_lazily_is_refused_before_any_work(case: str) -> None:
    with pytest.raises(VBAUnsupportedError, match='lazily'):
        evaluate_query(MEASURE.formula(case))
