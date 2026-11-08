"""Range.Sort, RemoveDuplicates, a filter's = and > and the Names collection over Hebrew, Arabic, Thai and Hangul, as live
Excel answered (scripts/measure_sort_scripts.py, tests/fixtures/sort_scripts.json): by the Windows word-sort weights
scripts/bake_sort_weights.py bakes for them."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_sort_scripts', ROOT / 'scripts/measure_sort_scripts.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/sort_scripts.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', CASES, ids=[f'{c["kind"]}-{c["group"]}-{c.get("matching", c.get("operator", ""))}' for c in CASES])
def test_each_case_as_excel_answered(case: dict[str, Any]) -> None:
    assert PROBE.model(case) == case['excel']


def test_every_kind_was_measured() -> None:
    assert {case['kind'] for case in CASES} == {'sort', 'duplicates', 'filter', 'names'}
