"""MsgBox with no answer queued comes back with the button vbDefaultButton2, 3 or 4 names, as Windows' message box
answers Enter (issue #250; scripts/measure_msgbox_default.py, tests/fixtures/msgbox_default.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/msgbox_default.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module('Public Function Asked(ByVal style As Long) As Long\n'
                    '    Asked = MsgBox("Overwrite?", style)\nEnd Function\n', name='Boxes')
    return made


@pytest.mark.parametrize('style', list(RECORD['answers']))
def test_the_default_button_answers(app: ExcelApplication, style: str) -> None:
    assert app.run('Boxes.Asked', int(style)) == RECORD['answers'][style]
