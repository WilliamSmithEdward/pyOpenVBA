"""A table's default member is its Name, the name given, where that differs from DisplayName (Tb7 shows as Tb7_), late bound and
typed, new and reopened (scripts/measure_list_object_default.py, tests/fixtures/list_object_default.json)."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_list_object_default', ROOT / 'scripts/measure_list_object_default.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/list_object_default.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', CASES, ids=[f'{case["name"]}-{case["typed"]}-{case["reopen"]}' for case in CASES])
def test_the_default_member_as_excel_answered(case: dict[str, Any], tmp_path: Path) -> None:
    assert PROBE.model(case['name'], case['typed'], case['reopen'], str(tmp_path)) == case['excel']


def test_the_rule_written_first_is_the_one_excel_follows() -> None:
    assert all(case['excel'] == case['rule_b'] for case in CASES)
    assert any(case['excel'] != case['rule_a'] for case in CASES)
