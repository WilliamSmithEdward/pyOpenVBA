"""ListObject.TableStyle set by name on a book with a style of its own takes a built-in name or the book's own in any case, spelt as
the book keeps it, and is error 450 for any other name (scripts/measure_table_style_own.py, tests/fixtures/table_style_own.json)."""
from __future__ import annotations

import importlib.util
import json
import shutil
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_table_style_own', ROOT / 'scripts/measure_table_style_own.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/table_style_own.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', CASES, ids=[f'{case["name"]!r}-{case["typed"]}' for case in CASES])
def test_each_name_as_excel_takes_it(case: dict[str, Any], tmp_path: Path) -> None:
    shutil.copyfile(PROBE.BOOK, tmp_path / 'made.xlsx')
    assert PROBE.model(case['name'], case['typed'], str(tmp_path)) == case['excel']
