"""Model-generated statistical rules paint the recorded cells after save and reopen."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/conditional'
CASES: list[dict[str, Any]] = [case for name in ('statistics_display.json', 'statistics_keys.json', 'statistics_criteria.json',
                                               'statistics_unicode.json', 'statistics_patterns.json')
                              for case in json.loads((ROOT / name).read_text(encoding='utf-8'))['probes']]


def saved_rendering(case: dict[str, Any], path: Path) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeControl')
    assert app.run('Probe') == case['result']
    app.add_module('''Function ReadPersistedDisplay() As String
Dim c As Range
For Each c In ''' + case['inspection_range'] + '''
If c.DisplayFormat.Font.Bold Then
ReadPersistedDisplay=ReadPersistedDisplay & "1"
Else
ReadPersistedDisplay=ReadPersistedDisplay & "0"
End If
Next
End Function''', name='PersistedInspection')
    expected = case['result'].split('|', 1)[0]
    assert app.run('ReadPersistedDisplay') == expected
    app.save(path)
    return expected


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_statistical_display_after_model_save(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / 'statistical-display.xlsm'
    expected = saved_rendering(case, path)
    assert ExcelApplication.open(path).run('ReadPersistedDisplay') == expected


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_statistical_display_after_native_open(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    path = tmp_path / 'native-statistical-display.xlsm'
    expected = saved_rendering(case, path)
    native_excel.open_document(path)
    assert native_excel.run_raw('PersistedInspection.ReadPersistedDisplay') == expected
    assert not any(event.get('dialogs') for event in native_excel.events)


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel
