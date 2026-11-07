"""A table's AutoFilter.Range is the whole table, totals row and all, on a new book and once Excel's saved file is opened again,
though the file's filter block leaves the totals row out (scripts/measure_table_filter_range.py,
tests/fixtures/table_filter_range.json)."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

ROOT = Path(__file__).parents[1]
_spec = importlib.util.spec_from_file_location('measure_table_filter_range', ROOT / 'scripts/measure_table_filter_range.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/table_filter_range.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', CASES, ids=[f'{case["order"]}-{case["reopen"]}-{case["typed"]}' for case in CASES])
def test_the_filter_range_as_excel_answered(case: dict[str, Any], tmp_path: Path) -> None:
    assert PROBE.model(PROBE.ORDERS[case['order']], case['reopen'], case['typed'], str(tmp_path)) == case['excel']


def test_excel_wrote_the_block_without_the_totals_row() -> None:
    assert {case['file_ref'] for case in CASES if case['reopen']} == {'H1:I4'}


def _probe(name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, ROOT / 'scripts' / f'{name}.py')
    assert spec is not None and spec.loader is not None
    module: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


ACTIVE = _probe('measure_table_filter_active')
BLOCK = _probe('measure_table_filter_block')
ACTIVE_CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/table_filter_active.json').read_text(encoding='utf-8'))['cases']
BLOCK_CASES: list[dict[str, Any]] = json.loads((ROOT / 'tests/fixtures/table_filter_block.json').read_text(encoding='utf-8'))['cases']


@pytest.mark.parametrize('case', ACTIVE_CASES, ids=[f'{c["where"]}-{c["reopen"]}-{c["totals"]}-{c["criteria"]}-{c["typed"]}' for c in ACTIVE_CASES])
def test_a_totals_row_is_in_the_range_only_while_its_sheet_is_in_front(case: dict[str, Any], tmp_path: Path) -> None:
    assert ACTIVE.model(case['where'], case['reopen'], case['totals'], case['criteria'], case['typed'], str(tmp_path)) == case['excel']


@pytest.mark.parametrize('case', BLOCK_CASES, ids=[f'{c["active"]}-{c["step"]}-{c["typed"]}' for c in BLOCK_CASES])
def test_the_range_after_opening_and_a_step(case: dict[str, Any], tmp_path: Path) -> None:
    assert BLOCK.model(case['active'], case['step'], case['typed'], str(tmp_path)).split(' | ')[0].split(' ')[1] == case['excel'].split(' | ')[0].split(' ')[1]
    assert BLOCK.model(case['active'], case['step'], case['typed'], str(tmp_path)).split(' | ')[1].split(' ')[1] == case['excel'].split(' | ')[1].split(' ')[1]
