"""Native valid calls beside uncompiled malformed procedure bodies."""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError
from pyopenvba.interpreter._parse import parse_module

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_semantics/body_demand.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_body_demand(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == case['result']
    if 'selected_result' in case:
        assert app.run('Harness.Unused') == case['selected_result']
    else:
        with pytest.raises(VBACompileError):
            app.run('Harness.Unused')
    assert app.run('Harness.Probe') == case['result']


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_explicit_parser_still_rejects_body_errors(case: dict[str, Any]) -> None:
    with pytest.raises(VBACompileError):
        parse_module(case.get('input_source', case['source']), name='Harness')


def _save_project(path: Path) -> None:
    case = next(case for case in RECORD['probes'] if case['name'] == 'missing_expression')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == 'done'
    app.save(path)


def test_body_demand_xlsm_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'body-demand.xlsm'
    _save_project(path)
    app = ExcelApplication.open(path)
    assert app.run('Harness.Probe') == 'done'
    with pytest.raises(VBACompileError):
        app.run('Harness.Unused')


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_native_excel_reads_body_demand_project(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    path = tmp_path / 'body-demand.xlsm'
    _save_project(path)
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        assert excel.run_raw('Harness.Probe') == 'done'
        assert not any(event.get('dialogs') for event in excel.events)
