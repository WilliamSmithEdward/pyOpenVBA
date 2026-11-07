"""A Range given to Workbook settings, typed and late, as live Excel answered (scripts/measure_workbook_setters_range.py), the model's answers
written first. The eight settings with no named gap answer every case as Excel did. Date1904, PrecisionAsDisplayed, SaveLinkValues and
DisplayDrawingObjects keep the gap tests/test_excel_workbook_setters_heldout.py names: given late, one area of two cells is taken by Excel,
where the model refuses it, and two areas Excel refuses with 1004."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_workbook_setters_range', ROOT / 'scripts/measure_workbook_setters_range.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/workbook_setters_range.json').read_text(encoding='utf-8'))
GAPS = frozenset({'Date1904', 'PrecisionAsDisplayed', 'SaveLinkValues', 'DisplayDrawingObjects'})
CASES = [(setting, value) for setting in MEASURE.SETTINGS for value in MEASURE.RANGES]


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
def test_each_case_answers_as_excel_did(declared: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(MEASURE.macro(declared), name='H')
    found = str(app.run('H.Probe')).split(MEASURE.SEPARATOR)
    differ = {case for case, mine, theirs in zip(CASES, found, RECORD['excel'][declared]) if mine != theirs}
    expected: set[tuple[str, str]] = {(setting, 'ws.Range("A1:B1")') for setting in GAPS} if declared == 'Object' else set()
    assert differ == expected
