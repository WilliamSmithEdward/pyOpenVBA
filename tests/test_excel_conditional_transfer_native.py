"""Independent native conditional-format transfer and import controls."""
from __future__ import annotations

import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/conditional_transfer'
CASES: list[dict[str, Any]] = json.loads((ROOT / 'controls.json').read_text(encoding='utf-8'))['probes']
PROVENANCE: dict[str, Any] = json.loads((ROOT / 'native52_provenance.json').read_text(encoding='utf-8'))
NATIVE52_IMPORTS: list[dict[str, Any]] = [
    {'name': 'fresh52-' + case['name'], 'file': case['native52_file'],
     'inspection_source': case['inspection_source'], 'expected': case['native52_snapshot']}
    for case in CASES if 'native52_file' in case
] + [
    {'name': 'converted51to52-' + case['name'], 'file': case['file'],
     'inspection_source': case['inspection_source'], 'expected': case['reopened']}
    for case in PROVENANCE['conversions']
]


def generated(case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeControl')
    assert app.run('Probe') == case['result']
    assert app.run('Snapshot') == case['final_snapshot']
    return app


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_transfer_api_display_and_model_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = generated(case)
    reopened = ExcelApplication.open(app.save(tmp_path / 'transferred.xlsm'))
    # Fresh FileFormat52 preserves the independently measured runtime ordering;
    # FileFormat51 inputs below have a separate native provenance oracle.
    assert reopened.run('Snapshot') == case['native52_snapshot']


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_created_transfer_import_and_forced_rewrite(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['inspection_source'], name='TransferInspection')
    # This oracle comes from native reopening the saved sample, not from a
    # guessed normalization of pre-save FormatCondition.NumberFormat strings.
    assert app.run('Snapshot') == case['sample_snapshot']
    for sheet in app.workbook.sheets_:
        if sheet.conditional_formats:
            sheet.conditional_changed = True
            sheet.touched()
    reopened = ExcelApplication.open(app.save(tmp_path / 'rewritten-transfer.xlsm'))
    assert reopened.run('Snapshot') == case['sample_snapshot']


@pytest.mark.parametrize('case', NATIVE52_IMPORTS, ids=lambda case: case['name'])
def test_native_macro_enabled_transfer_import_provenance(case: dict[str, Any], tmp_path: Path) -> None:
    # These independently native-created projects have no auto-run entrypoint;
    # exclude their stored VBA and inject only the read-only inspection helpers.
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['inspection_source'], name='TransferInspection')
    assert app.run('Snapshot') == case['expected']
    for sheet in app.workbook.sheets_:
        if sheet.conditional_formats:
            sheet.conditional_changed = True
            sheet.touched()
    reopened = ExcelApplication.open(app.save(tmp_path / 'rewritten-native52.xlsm'))
    assert reopened.run('Snapshot') == case['expected']


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_reopens_model_transferred_workbook(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = generated(case)
    path = app.save(tmp_path / 'native-transferred.xlsm')
    native_excel.open_document(path)
    # Compare to the independently measured native FileFormat52 reopen oracle,
    # not the distinct FileFormat51 input's API projection.
    assert native_excel.run_raw('NativeControl.Snapshot') == case['native52_snapshot']
    assert not any(event.get('dialogs') for event in native_excel.events)
