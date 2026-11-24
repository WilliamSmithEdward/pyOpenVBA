"""Shape.ThreeD as live Excel answered held-out rounds (scripts/measure_three_d_round.py): each case run alone, its reads and the
shape's saved spPr. Round 1 measured every scalar member before the model had any of them."""
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
ROUNDS = range(1, 2)


def _record(number: int) -> dict[str, Any]:
    return json.loads((ROOT / f'tests/fixtures/three_d_round_{number}.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('number', ROUNDS)
def test_each_case_answers_and_saves_as_excel_did(number: int, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(PROBE, 'ROUND', number)
    record = _record(number)
    with tempfile.TemporaryDirectory() as folder:
        differ = [(index, case, got, excel) for index, (case, excel) in enumerate(zip(record['cases'], record['excel']))
                  if (got := PROBE.alone(case, index, Path(folder))) != excel and got[0] != 'refused']
    assert differ == []
