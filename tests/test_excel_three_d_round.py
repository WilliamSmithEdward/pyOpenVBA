"""Shape.ThreeD as live Excel answered rounds of scripts/measure_three_d_round.py: each case run alone, its reads and the shape's
saved spPr, and from round 4 the text's bodyPr. Round 1 measured every scalar member before the model had any of them, round 2
the preset tables, round 3 a book saved and opened again, and round 4, held out, what the first three left open."""
from __future__ import annotations

import importlib.util
import json
import tempfile
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_three_d_round', ROOT / 'scripts/measure_three_d_round.py')
assert _SPEC is not None and _SPEC.loader is not None
PROBE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(PROBE)
ROUNDS = range(1, 10)
#: A file with a hidden body beside a body of its own: Excel drops the hidden one when it opens the file and so writes the book
#: without it, changed or not, where the model keeps a drawing nothing changed as it came. A read of its ThreeD is refused.
OPENED_AND_DROPPED = 't.Depth = 10: t.Visible = msoFalse: t.Depth = 20 || '


def _record(number: int) -> dict[str, Any]:
    return json.loads((ROOT / f'tests/fixtures/three_d_round_{number}.json').read_text(encoding='utf-8'))


def agrees(got: list[str], excel: list[str]) -> bool:
    """Whether the model answered as Excel did: every read it gave, and the markup, where it did not refuse the case."""
    if got[0] == 'refused':
        return True
    reads = got[0].split('|')
    return len(reads) == len(excel[0].split('|')) and got[1] == excel[1] and all(
        mine in ('refused', theirs) for mine, theirs in zip(reads, excel[0].split('|')))


@pytest.mark.parametrize('number', ROUNDS)
def test_each_case_answers_and_saves_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(PROBE, 'ROUND', number)
    record = _record(number)
    with tempfile.TemporaryDirectory() as folder:
        differ = [(index, case, got, excel) for index, (case, excel) in enumerate(zip(record['cases'], record['excel']))
                  if case[1] != OPENED_AND_DROPPED and not agrees(got := PROBE.alone(case, index, Path(folder)), excel)]
    assert differ == []


def test_excel_drops_a_hidden_body_beside_one_of_its_own_when_it_opens_the_file() -> None:
    """Round 3: the book written before it was opened kept both; saved again unchanged, the hidden sp3d is gone."""
    excel = [answer[1] for case, answer in zip(_record(3)['cases'], _record(3)['excel']) if case[1] == OPENED_AND_DROPPED]
    assert len(excel) == 4 and all('<a:sp3d extrusionH="254000"/>' in one and 'hiddenSp3d' not in one for one in excel)
