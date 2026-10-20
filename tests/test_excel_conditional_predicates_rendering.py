"""Independent native typed, Unicode and calendar-boundary predicate controls."""
from __future__ import annotations

import datetime as dt
import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/conditional'
CASES: list[dict[str, Any]] = [case for name in ('display', 'time') for case in
                              json.loads((ROOT / f'predicates_{name}.json').read_text(encoding='utf-8'))['probes']]


def freeze_native_day(app: ExcelApplication, case: dict[str, Any], monkeypatch: pytest.MonkeyPatch) -> None:
    if 'native_clock' in case:
        moment = dt.datetime.fromisoformat(case['native_clock'].split('|')[1])
        monkeypatch.setattr(app.interpreter, 'clock', lambda: moment)


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_predicate_display(case: dict[str, Any], monkeypatch: pytest.MonkeyPatch) -> None:
    app = ExcelApplication()
    app.add_workbook()
    freeze_native_day(app, case, monkeypatch)
    app.add_module(case['source'], name='NativeControl')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', [case for case in CASES if case.get('file')], ids=lambda case: case['name'])
def test_native_predicate_import_and_model_roundtrip(case: dict[str, Any], monkeypatch: pytest.MonkeyPatch,
                                                   tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    freeze_native_day(app, case, monkeypatch)
    app.add_module(case['source'], name='NativeControl')
    app.add_module('Function ReadImported() As String\nReadImported=Bits(Range("' + case['inspection_target']
                   + '"))\nEnd Function', name='ImportInspection')
    assert app.run('ReadImported') == case['result']
    reopened = ExcelApplication.open(app.save(tmp_path / 'imported-predicate.xlsm'))
    freeze_native_day(reopened, case, monkeypatch)
    assert reopened.run('ReadImported') == case['result']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_generated_predicate_display(case: dict[str, Any], tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeControl')
    # Calendar-relative controls deliberately retain their bit vectors on later days.
    assert app.run('Probe') == case['result']
    path = app.save(tmp_path / 'generated-predicate.xlsm')
    with harness.ExcelSession() as excel:
        excel.open_document(path)
        target = case.get('inspection_target', 'A1:A10' if case['name'].startswith('unicode-') else 'A1:A18')
        source = 'Function ReadGenerated() As String\nReadGenerated=Bits(Range("' + target + '"))\nEnd Function'
        result = excel.run_vba(source, 'ReadGenerated', line_numbers=False)
        assert result.ok and result.value == case['result']
        assert not any(event.get('dialogs') for event in excel.events)
