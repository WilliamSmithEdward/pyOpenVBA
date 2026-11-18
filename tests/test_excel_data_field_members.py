"""PivotField members read on a pivot table's data field, its row field and its first field, on a new book and one Excel saved,
as live Excel answered (scripts/measure_data_field_members.py, tests/fixtures/data_field_members.json)."""
from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_data_field_members', ROOT / 'scripts/measure_data_field_members.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, str]] = json.loads((ROOT / 'tests/fixtures/data_field_members.json').read_text(encoding='utf-8'))['cases']


@pytest.fixture(scope='module')
def folder() -> Any:
    with tempfile.TemporaryDirectory() as held:
        PROBE.rounds.books(Path(held))
        yield held


@pytest.mark.parametrize('case', CASES, ids=[f'{case["book"]}-{case["field"]}-{case["read"]}' for case in CASES])
def test_each_read_as_excel_answered(case: dict[str, str], folder: str) -> None:
    assert PROBE.model(case, False, folder) == case['late']
    assert PROBE.model(case, True, folder) == case['typed']
