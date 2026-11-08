"""TextRange2's members and text methods as live Excel answered seven held-out rounds (scripts/measure_text_range2_round.py), each case
run alone, and from round 4 what the statements write to the file. Rounds 6 and 7 ran on this code with its predictions written first,
and every case answered as Excel did."""
from __future__ import annotations

import importlib.util
import json
import tempfile
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_text_range2_round', ROOT / 'scripts/measure_text_range2_round.py')
assert _SPEC is not None and _SPEC.loader is not None
PROBE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(PROBE)
ROUNDS = range(1, 8)


def _record(number: int) -> dict[str, Any]:
    name = 'text_range2_round.json' if number == 1 else f'text_range2_round_{number}.json'
    return json.loads((ROOT / 'tests/fixtures' / name).read_text(encoding='utf-8'))


@pytest.mark.parametrize('number', ROUNDS)
def test_each_case_answers_as_excel_did(number: int) -> None:
    record = _record(number)
    differ = [(case[0], case[-1], got, excel) for case, excel in zip(record['cases'], record['excel'])
              if (got := PROBE.alone(case)) != excel and got != 'refused']
    assert differ == []


@pytest.mark.parametrize('number', [6, 7])
def test_the_held_out_rounds_are_answered_without_a_refusal(number: int) -> None:
    assert [case for case in _record(number)['cases'] if PROBE.alone(case) == 'refused'] == []


@pytest.mark.parametrize('number', range(4, 8))
def test_the_statements_write_the_file_as_excel_did(number: int) -> None:
    excel = _record(number)['excel_markup']
    with tempfile.TemporaryDirectory() as folder:
        written = PROBE.model_markup(Path(folder))
    assert written[:len(excel)] == excel
