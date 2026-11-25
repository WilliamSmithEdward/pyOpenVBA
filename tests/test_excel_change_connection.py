"""PivotTable.ChangeConnection on a table over a list, given Nothing, other objects and values, as a statement and through
CallByName, as live Excel answered (scripts/measure_change_connection.py, tests/fixtures/change_connection.json)."""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
_spec = importlib.util.spec_from_file_location('measure_change_connection', ROOT / 'scripts/measure_change_connection.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, str]] = json.loads((ROOT / 'tests/fixtures/change_connection.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', CASES, ids=[f'{case["argument"]}-{at % 2}' for at, case in enumerate(CASES)])
def test_each_argument_as_excel_answered(case: dict[str, str]) -> None:
    assert PROBE.model(case['argument'], case['way']) == case['excel']
