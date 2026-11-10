"""PivotTable.ListFormulas and its workbook's Close after it, with and without another workbook in front, as live Excel answered
(scripts/measure_list_formulas_close.py, tests/fixtures/list_formulas_close.json). With another book in front Excel raises
-2147417848, adds the sheet all the same and then cannot close the book; the model reports that rather than copy it."""
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
_spec = importlib.util.spec_from_file_location('measure_list_formulas_close', ROOT / 'scripts/measure_list_formulas_close.py')
assert _spec is not None and _spec.loader is not None
PROBE: Any = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(PROBE)
CASES: list[dict[str, str]] = json.loads((ROOT / 'tests/fixtures/list_formulas_close.json').read_text(encoding='utf-8'))['cases']
REPORTED = 'model VBAUnsupportedError: PivotTable.ListFormulas with another workbook in front is not implemented'


@pytest.fixture(scope='module')
def folder() -> Any:
    with tempfile.TemporaryDirectory() as held:
        PROBE.rounds.books(Path(held))
        yield held


@pytest.mark.parametrize('case', CASES, ids=[f'{case["book"]}-{case["edit"]}-{case["list"]}' for case in CASES])
def test_each_case_as_excel_answered(case: dict[str, str], folder: str) -> None:
    answer = PROBE.model(case, folder)
    if case['edit'] == 'ws2.Columns(1).Insert':
        # The column goes in before the table, which Excel moves and the model reports (tests/test_excel_pivot_shift_refusal.py).
        assert answer.startswith('model VBAUnsupportedError: Inserting or deleting cells that would move a pivot table')
    elif case['book'] == 'file, another in front' and case['list'] == 'True':
        assert answer.startswith(REPORTED) and ' close -2147417848' in case['excel']
    else:
        assert answer == case['excel']
