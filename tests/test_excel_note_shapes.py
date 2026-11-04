"""A note box's Locked, Placement and fill colour, replayed from live Excel (scripts/measure_note_shapes.py, tests/fixtures/note_shapes.json):
what it reads after a change, and after the model's save is opened again. `model_reopened` there is Excel's reading of the model's file, the same as its own."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'note_shapes.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['cases'], ids=[f'{one["target"]} {one["change"]}' for one in RECORD['cases']])
def test_each_change_as_excel_answered(case: dict[str, str], tmp_path: Path) -> None:
    app = ExcelApplication.open(FIXTURES / 'object_roundtrip.xlsx', with_vba=False)
    app.interpreter.sandbox = False
    app.add_module(RECORD['source'], name='Probe')
    assert str(app.run('Probe.InModel', case['target'], RECORD['changes'].index(case['change']))) == case['excel']
    assert case['model_reopened'] == case['reopened']
    path = tmp_path / 'saved.xlsx'
    app.save(path)
    assert str(app.run('Probe.ReadAt', str(path), case['target'])) == case['reopened']


def test_a_shape_in_a_group_from_a_file_refuses_a_change_a_save_would_lose() -> None:
    app = ExcelApplication.open(FIXTURES / 'object_roundtrip.xlsx', with_vba=False)
    app.add_module('Public Sub Paint()\n    ActiveWorkbook.Worksheets(1).Shapes("Grouped").GroupItems(1).Fill.ForeColor.RGB = 255\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError, match='in a group'):
        app.run('Probe.Paint')
