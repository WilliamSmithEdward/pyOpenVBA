"""What a cell holds when a macro sets it to digits of another script, as live Excel answered
(scripts/measure_typing_native_digits.py, tests/fixtures/typing_native_digits.json): every decimal digit of the Basic
Multilingual Plane alone, and each set's digits as "12", "-12", "1.2" and "1,2"."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel._typing import typed_text

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_typing_native_digits', ROOT / 'scripts/measure_typing_native_digits.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, str]] = json.loads((ROOT / 'tests/fixtures/typing_native_digits.json').read_text(encoding='utf-8'))['cases']


def test_every_digit_types_as_excel_typed_it() -> None:
    texts = [case['text'] for case in CASES]
    answers = [answer for at in range(0, len(texts), PROBE.CHUNK) for answer in PROBE.model(texts[at:at + PROBE.CHUNK])]
    missed = [(case['text'], case['excel'], answer) for case, answer in zip(CASES, answers) if answer != case['excel']]
    assert not missed, missed[:3]
    assert len(answers) == len(CASES)


_shapes = importlib.util.spec_from_file_location('measure_typing_native_shapes', ROOT / 'scripts/measure_typing_native_shapes.py')
assert _shapes is not None and _shapes.loader is not None
SHAPES: Any = importlib.util.module_from_spec(_shapes)
_shapes.loader.exec_module(SHAPES)
SHAPE_CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/typing_native_shapes.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', SHAPE_CASES, ids=[f'{case["zero"]}-{case["shape"]!r}-{at}' for at, case in enumerate(SHAPE_CASES)])
def test_each_shape_types_with_the_value_and_format_excel_gave(case: dict[str, Any]) -> None:
    """Dates, times, percents, currency, exponents, fractions and thousands in another script's digits, alone, mixed
    with ASCII and with a second set, and Arabic's own separators, which keep the text."""
    assert SHAPES.model([case['text']]) == [case['excel']]


@pytest.mark.parametrize('text', ['١٢٣', '١٢', '๙'])
def test_digits_of_one_set_are_a_number(text: str) -> None:
    assert isinstance(typed_text(text).value, float)


@pytest.mark.parametrize('text', ['ب١', '١ب', 'a๑', '１', '߁'],
                         ids=['arabic letter and digit', 'digit and arabic letter', 'latin and thai digit',
                              'full-width digit', 'nko digit'])
def test_text_round_the_digits_stays_text(text: str) -> None:
    assert typed_text(text).value == text


def test_a_formatted_cell_reads_a_fraction_in_another_scripts_digits_as_a_half() -> None:
    assert typed_text('١/٢', fractions=True).value == 0.5
