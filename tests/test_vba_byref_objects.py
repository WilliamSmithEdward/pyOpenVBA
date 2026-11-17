"""A variable passed ByRef to an object parameter, or an object to a value's, compiles where Excel's VBA does, over 92 live compiles."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_byref_objects.json').read_text(encoding='utf-8'))
THING = 'Public Value As Long\n'


def _answer(source: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(THING, name='Thing', kind='class')
        app.add_module(source, name='Harness')
        return str(app.run('Harness.Probe'))
    except VBACompileError as caught:
        return 'compile error: ' + caught.message


@pytest.mark.parametrize(('case', 'excel'), list(zip(RECORD['cases'], RECORD['excel'])),
                         ids=[f"{case['argument']}-{case['parameter']}-{case['form']}" for case in RECORD['cases']])
def test_byref_object_compiles_as_excel(case: dict[str, Any], excel: str) -> None:
    assert _answer(case['source']) == excel
