"""Native-created statistical files retain measured rule properties and displayed matches."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/conditional'
API: list[dict[str, Any]] = [case for case in json.loads((ROOT / 'statistics_api.json').read_text(encoding='utf-8'))['probes']
                            if case.get('file')]
DISPLAY: list[dict[str, Any]] = [case for case in json.loads((ROOT / 'statistics_display.json').read_text(encoding='utf-8'))['probes']
                                if case.get('file')]


@pytest.mark.parametrize('case', API, ids=lambda case: case['name'])
def test_native_statistical_file_properties(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='NativeControl')
    app.add_module('''Function ReadImportedRule() As String
ReadImportedRule=Inspect(Cells.FormatConditions(1))
End Function''', name='ImportInspection')
    expected = 'class=' + case['result'].split('class=', 1)[1].split('|count=', 1)[0]
    assert app.run('ReadImportedRule') == expected


@pytest.mark.parametrize('case', DISPLAY, ids=lambda case: case['name'])
def test_native_statistical_file_display(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='NativeControl')
    if case['name'].startswith('union-six-'):
        target = 'Union(Range("A1:A3"),Range("C5:C7"))'
    elif case['name'].startswith('overlap-'):
        target = 'Range("A1:A6")'
    else:
        target = 'Range("A1:A13")'
    app.add_module('Function ReadImportedDisplay() As String\nReadImportedDisplay=Bits(' + target
                   + ')\nEnd Function', name='ImportInspection')
    assert app.run('ReadImportedDisplay') == case['result'].split('|', 1)[0]
