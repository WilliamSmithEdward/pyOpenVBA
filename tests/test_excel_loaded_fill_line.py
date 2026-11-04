"""Fill and Line set on shapes a file holds, replayed from live Excel (scripts/measure_loaded_fill_line.py, tests/fixtures/loaded_fill_line.json):
what the shape reads after the change, and after the model's save is opened again. Excel read the same saves when it was measured: `model_reopened` there
is Excel's reading of the model's file, and every case that is not a report is the same as Excel's reading of its own."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'loaded_fill_line.json').read_text(encoding='utf-8'))
READER = RECORD['source'] + '''
Public Function Saved(ByVal path As String, ByVal target As String) As String
    Dim wb As Object
    Set wb = Workbooks.Open(path, 0)
    Saved = ReadShape(wb.Worksheets(1).Shapes(target))
    wb.Close False
End Function
'''


@pytest.mark.parametrize('case', RECORD['cases'], ids=[f'{one["target"]} {one["change"]}' for one in RECORD['cases']])
def test_each_change_as_excel_answered(case: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication.open(FIXTURES / 'object_roundtrip.xlsx', with_vba=False)
    app.interpreter.sandbox = False
    app.add_module(READER, name='Probe')
    which = RECORD['changes'].index(case['change'])
    try:
        got = str(app.run('Probe.InModel', case['target'], which))
    except VBAUnsupportedError:
        assert case['model_reopened'].startswith('reports')
        return
    assert got == case['excel']
    assert case['model_reopened'] == case['reopened']
    path = tmp_path / 'saved.xlsx'
    app.save(path)
    assert str(app.run('Probe.Saved', str(path), case['target'])) == case['reopened']
