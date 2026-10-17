"""Native VBA PivotTable.GetPivotData reference and optional-argument behavior."""
from __future__ import annotations
import json
import importlib.util
import os
from pathlib import Path
from typing import Any
import pytest
from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/pivots'
RECORD: dict[str, Any] = json.loads((ROOT / 'getdata.json').read_text(encoding='utf-8'))
FILES = list(dict.fromkeys(case['file'] for case in RECORD['probes']))
PROJECTS = [next(case for case in reversed(RECORD['probes']) if case['file'] == file and case['result'].startswith('$'))
            for file in FILES]


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['file'] + ':' + case['arguments'])
def test_native_pivot_getdata(case: dict[str, Any]) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    assert app.run(case['source'].split('(')[0].split()[-1]) == case['result']


def _project(case: dict[str, Any], path: Path) -> str:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['source'], name='Harness')
    app.save(path)
    return case['source'].split('(')[0].split()[-1]


@pytest.mark.parametrize('case', PROJECTS, ids=lambda case: case['file'])
def test_getdata_project_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    path = tmp_path / 'getdata.xlsm'
    name = _project(case, path)
    reopened = ExcelApplication.open(path)
    assert reopened.run(name) == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', PROJECTS, ids=lambda case: case['file'])
def test_native_excel_runs_saved_getdata_project(case: dict[str, Any], tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'getdata.xlsm'
    name = _project(case, path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.' + name) == case['result']
        assert not any(event.get('dialogs') for event in excel.events)
