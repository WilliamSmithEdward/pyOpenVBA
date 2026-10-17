"""Native SpreadsheetML escapes and UTF-16 text/formula roundtrips."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/cell_strings'
RECORD: dict[str, Any] = json.loads((ROOT / 'controls.json').read_text(encoding='utf-8'))


def _project(mode: str, path: Path) -> Path:
    if mode == 'native':
        app = ExcelApplication.open(ROOT / 'native.xlsx', with_vba=False)
        app.add_module(RECORD['inspection_source'], name='Harness')
        assert app.run('Snapshot') == RECORD['result']
    else:
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(RECORD['source'], name='Harness')
        assert app.run('Probe') == RECORD['result']
    return app.save(path)


@pytest.mark.parametrize('mode', ['native', 'fresh'])
def test_native_string_roundtrip(mode: str, tmp_path: Path) -> None:
    path = _project(mode, tmp_path / 'strings.xlsm')
    for _ in range(3):
        app = ExcelApplication.open(path)
        assert app.run('Snapshot') == RECORD['result']
        path = app.save(path)


@pytest.mark.parametrize('case', RECORD['unicode_probes'], ids=lambda case: case['name'])
def test_native_utf16_primitives(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / 'native.xlsx', with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('mode', ['native', 'fresh'])
def test_native_recalculates_saved_strings(mode: str, tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = _project(mode, tmp_path / 'strings.xlsm')
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.Snapshot') == RECORD['result']
        result = excel.run_vba('Sub Recalculate()\nApplication.CalculateFull\nEnd Sub',
                               'Recalculate', line_numbers=False)
        assert result.ok and not result.dialogs, result
        assert excel.run_raw('Harness.Snapshot') == RECORD['result']
        assert not any(event.get('dialogs') for event in excel.events)
