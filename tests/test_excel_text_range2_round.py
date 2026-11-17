"""TextRange2's members and text methods as live Excel answered held-out rounds (scripts/measure_text_range2_round.py), each case
run alone, and from round 4 what the statements write to the file. Rounds 6 and 7 ran on this code with its predictions written first,
and every case answered as Excel did. Rounds 8 to 10 measured what the model lacked: For Each, the bounds, InsertChartField and the
text clipboard, rounds 9 and 10 each case from a clipboard it set itself."""
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
ROUNDS = range(1, 11)


def _record(number: int) -> dict[str, Any]:
    name = 'text_range2_round.json' if number == 1 else f'text_range2_round_{number}.json'
    return json.loads((ROOT / 'tests/fixtures' / name).read_text(encoding='utf-8'))


@pytest.mark.parametrize('number', ROUNDS)
def test_each_case_answers_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(PROBE, 'ROUND', number)
    record = _record(number)
    differ = [(case[0], case[-1], got, excel) for case, excel in zip(record['cases'], record['excel'])
              if (got := PROBE.alone(case)) != excel and got != 'refused']
    assert differ == []


@pytest.mark.parametrize('number', [6, 7])
def test_the_held_out_rounds_are_answered_without_a_refusal(number: int) -> None:
    assert [case for case in _record(number)['cases'] if PROBE.alone(case) == 'refused'] == []


@pytest.mark.parametrize('number', range(4, 9))
def test_the_statements_write_the_file_as_excel_did(number: int) -> None:
    excel = _record(number)['excel_markup']
    with tempfile.TemporaryDirectory() as folder:
        written = PROBE.model_markup(Path(folder))
    assert written[:len(excel)] == excel


def test_a_paste_writes_the_file_as_excel_did(monkeypatch: pytest.MonkeyPatch) -> None:
    """Round 10's markup, each statement alone: what the model pastes is written as Excel wrote it, and it refuses the rest by
    name (an HTML paste, copied cells, all the text with its paragraph)."""
    monkeypatch.setattr(PROBE, 'ROUND', 10)
    excel = _record(10)['excel_markup']
    with tempfile.TemporaryDirectory() as folder:
        written = PROBE.model_markup_each(Path(folder))
    assert [index for index, (mine, theirs) in enumerate(zip(written, excel)) if mine not in (theirs, 'refused')] == []
    assert [index for index, mine in enumerate(written) if mine == 'refused'] == [4, 6, 7]
