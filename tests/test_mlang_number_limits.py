"""Number literals near and past the largest Double, as live Excel's Power Query evaluates them (scripts/measure_m_number_limits.py).

A hex literal is an exact unsigned 64-bit integer, wrapping past 2**64, never a Double: 0x20000000000001 stays 9007199254740993, and the fuzzer's
0xe111... of 300 digits is its low 64 bits. A decimal literal is a Double: past the largest one it is infinity, which Text.From writes as "∞",
and below the smallest it is 0. Before, a hex literal too large for a float raised OverflowError out of the parser."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.mlang import evaluate_query
from pyopenvba.mlang._parse import MSyntaxError, parse

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_m_number_limits', ROOT / 'scripts/measure_m_number_limits.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
ANSWERS: dict[str, str] = json.loads((ROOT / 'tests/fixtures/m_number_limits.json').read_text(encoding='utf-8'))['answers']
#: Known gaps of a family of their own, not of the crash: Power Query keeps a decimal literal exact past a Double's 2**53 and writes a Double's
#: text in .NET's way, 1E+308 and 1E-05, where the model rounds to a Double and writes Python's.
GAPS = frozenset({'9007199254740993', '12345678901234567', '1e308', '1.7976931348623157e308', '1.7976931348623158e308', '1e-5', '1e-7'})


@pytest.mark.parametrize('literal', [literal for literal in ANSWERS if literal not in GAPS], ids=lambda literal: literal[:40])
def test_a_number_literal_answers_as_power_query_did(literal: str) -> None:
    assert evaluate_query(MEASURE.formula(literal)) == ANSWERS[literal]


@pytest.mark.parametrize('literal', sorted(GAPS))
def test_the_named_gaps_still_differ(literal: str) -> None:
    """Should one of these come to answer as Power Query did, it leaves GAPS."""
    assert evaluate_query(MEASURE.formula(literal)) != ANSWERS[literal]


def test_the_fuzzers_input_is_a_syntax_error_not_an_overflow() -> None:
    """The fuzzer's input, as its ASCII decoding saw it: not M, so a syntax error, where it was an OverflowError from the hex literal in it."""
    seed = (ROOT / 'tests/fuzz_corpus/mformula/regression_hex_too_large.m').read_text(encoding='utf-8')
    with pytest.raises(MSyntaxError):
        parse(seed)
    assert evaluate_query('0x' + 'e' + '1' * 300) == 1229782938247303441
