"""A freeform's ShapeNodes as live Excel answered the rounds of scripts/measure_shape_nodes_round.py: each case run alone, every
node's SegmentType, EditingType and point, the shape's box, and the saved path. Round 1 measured each editing method and its
errors, round 2 new freeforms of other lengths and line freeforms edited more than once. The model refuses what it has not
measured; a case it answers must answer as Excel did."""
from __future__ import annotations

import importlib.util
import json
import tempfile
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_shape_nodes_round', ROOT / 'scripts/measure_shape_nodes_round.py')
assert _SPEC is not None and _SPEC.loader is not None
PROBE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(PROBE)
ROUNDS = range(1, 7)


def _record(number: int) -> dict[str, Any]:
    return json.loads((ROOT / f'tests/fixtures/shape_nodes_round_{number}.json').read_text(encoding='utf-8'))


def _parts(answer: str) -> list[str]:
    """An answer's reads, the nodes' each split into its SegmentType, EditingType and point."""
    reads = answer.split('|')
    return [reads[0], *(part for node in reads[1].split(';') for part in node.split('/')), *reads[2:]] if len(reads) > 1 else reads


def agrees(got: list[str], excel: list[str]) -> bool:
    """Whether the model answered as Excel did: the markup, and every read but an EditingType it left out."""
    if got[0] == 'refused':
        return True
    mine, theirs = _parts(got[0]), _parts(excel[0])
    return got[1] == excel[1] and len(mine) == len(theirs) and all(a in (PROBE.LEFT_OUT, b) for a, b in zip(mine, theirs))


@pytest.mark.parametrize('number', ROUNDS)
def test_each_case_answers_and_saves_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(PROBE, 'ROUND', number)
    record = _record(number)
    with tempfile.TemporaryDirectory() as folder:
        differ = [(index, case, got, excel) for index, (case, excel) in enumerate(zip(record['cases'], record['excel']))
                  if excel[0] != PROBE.BREAKS_EXCEL and not agrees(got := PROBE.alone(case, index, Path(folder)), excel)]
    assert differ == []


@pytest.mark.parametrize('number', ROUNDS)
def test_a_new_freeform_answers_its_nodes_where_measured(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    """Every new freeform of a family the model knows is answered, not refused: open lines of 2 to 5 nodes, a closed line
    of 3, and one curve."""
    monkeypatch.setattr(PROBE, 'ROUND', number)
    known = {'open', 'closed', 'two', 'four', 'five', 'curve1'}
    with tempfile.TemporaryDirectory() as folder:
        refused = [case for index, case in enumerate(_record(number)['cases'])
                   if case[1] == '' and case[0] in known and PROBE.alone(case, index, Path(folder))[0] == 'refused']
    assert refused == []
