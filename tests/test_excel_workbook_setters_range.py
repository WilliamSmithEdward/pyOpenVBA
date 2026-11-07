"""A Range given to Workbook settings, typed and late, as live Excel answered (scripts/measure_workbook_setters_range.py), the model's answers
written first. The eight settings with no named gap answer every case as Excel did. Date1904, PrecisionAsDisplayed, SaveLinkValues and
DisplayDrawingObjects, given late one area of two cells, which Excel takes and changes, are refused openly: an object given by Let answers
by its member's measured rule (src/pyopenvba/interpreter/_late_object_lets.py), and these were not measured to leave the member as it was."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).resolve().parents[1]
_SPEC = importlib.util.spec_from_file_location('measure_workbook_setters_range', ROOT / 'scripts/measure_workbook_setters_range.py')
assert _SPEC is not None and _SPEC.loader is not None
MEASURE: Any = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(MEASURE)
RECORD: dict[str, Any] = json.loads((ROOT / 'tests/fixtures/workbook_setters_range.json').read_text(encoding='utf-8'))
GAPS = frozenset({'Date1904', 'PrecisionAsDisplayed', 'SaveLinkValues', 'DisplayDrawingObjects'})
CASES = [(setting, value) for setting in MEASURE.SETTINGS for value in MEASURE.RANGES]


def _alone(declared: str, setting: str, value: str) -> str:
    source = MEASURE.macro(declared)
    head, _, rest = source.partition('    Set wb = Workbooks.Add(xlWBATWorksheet)\n')
    block = '    Set wb = Workbooks.Add(xlWBATWorksheet)\n' + rest.split('    wb.Close SaveChanges:=False\n')[0] + '    wb.Close SaveChanges:=False\n'
    case = block.replace(f'wb.{MEASURE.SETTINGS[0]} = {MEASURE.RANGES[0]}', f'wb.{setting} = {value}')
    case = case.replace(f'v = wb.{MEASURE.SETTINGS[0]}', f'v = wb.{setting}')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(head + case + '    On Error GoTo 0\n    Probe = Mid(out, 3)\nEnd Function\n', name='H')
    try:
        return str(app.run('H.Probe'))
    except VBAUnsupportedError:
        return 'refused'


@pytest.mark.parametrize('declared', ['Workbook', 'Object'])
def test_each_case_answers_as_excel_did_or_is_refused_openly(declared: str) -> None:
    differ = {(setting, value): mine for (setting, value), theirs in zip(CASES, RECORD['excel'][declared])
              if (mine := _alone(declared, setting, value)) != theirs}
    expected: dict[tuple[str, str], str] = ({(setting, 'ws.Range("A1:B1")'): 'refused' for setting in GAPS} if declared == 'Object' else {})
    assert differ == expected
