"""Independent native Pivot refresh controls and explicit model admission limits."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from hashlib import sha256
from pathlib import Path, PureWindowsPath
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/pivot_refresh'
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf8'))
RECORDS: dict[str, Any] = {name: json.loads((ROOT / name).read_text(encoding='utf8'))
                           for name in MANIFEST['raw_files']}
CASES = [(filename, case) for filename, record in RECORDS.items() for case in record['probes']
         if MANIFEST['cases'][filename + ':' + case['name']]['admission'] == 'runtime']
EVENTS = [case for case in RECORDS['events_save.json']['probes'] if 'file' in case]
IMPORTED = RECORDS['imported.json']['probes']
COLLISIONS = [case for case in RECORDS['growth.json']['probes'] if case['name'] != 'blank-growth']
SCHEMA = [case for case in RECORDS['followup.json']['probes'] if case['name'].startswith('changed-header-')]


def configured(case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    for name, source in case.get('classes', {}).items():
        app.add_module(source, name=name, kind='class')
    app.add_module(case['source'], name='RefreshProbe')
    return app


def without_host_build(value: str) -> str:
    """Explicit host-only projection, preserving Version and all behavior fields."""
    version, pipe, rest = value.partition('|')
    build, semicolon, behavior = rest.partition(';')
    assert version.startswith('version=') and pipe and semicolon and build.isdecimal()
    return version + ';' + behavior


@pytest.mark.parametrize('filename', RECORDS)
def test_native_recorded_json_original_bytes(filename: str) -> None:
    assert sha256((ROOT / filename).read_bytes()).hexdigest() == MANIFEST['raw_files'][filename]


@pytest.mark.parametrize('filename,case', CASES, ids=[case['name'] for _name, case in CASES])
def test_native_refresh_behavior(filename: str, case: dict[str, Any]) -> None:
    admission = MANIFEST['cases'][filename + ':' + case['name']]
    assert sha256(case['source'].encode()).hexdigest() == admission['source_sha256']
    app = configured(case)
    actual = app.run('Probe')
    assert isinstance(actual, str)
    if admission['build_projection']:
        assert without_host_build(actual) == without_host_build(case['raw_result'])
    else:
        assert actual == case['raw_result']


@pytest.mark.parametrize('sample', MANIFEST['samples'], ids=lambda sample: sample['case'])
def test_original_native_refresh_package_bytes(sample: dict[str, Any]) -> None:
    path = ROOT / sample['filename']
    assert sha256(path.read_bytes()).hexdigest() == sample['sha256']
    with ZipFile(path) as package:
        assert sha256(package.read('xl/vbaProject.bin')).hexdigest() == sample['vba_sha256']


@pytest.mark.parametrize('enabled', [False, True])
@pytest.mark.parametrize('case', EVENTS, ids=lambda case: case['name'])
def test_native_refresh_input_public_enable_refresh_save_preserves_records(
        case: dict[str, Any], enabled: bool, tmp_path: Path) -> None:
    # Original native52 package is opened without loading its recorded macros.
    original = ROOT / (case['name'] + '.xlsm')
    with ZipFile(original) as package:
        records = {name: package.read(name) for name in package.namelist()
                   if name.startswith('xl/pivotCache/pivotCacheRecords') and name.endswith('.xml')}
    assert records
    app = ExcelApplication.open(original, with_vba=False)
    app.add_module(case['inspection_source'], name='RefreshRead')
    assert app.run('Snapshot') == case['reopened']
    first = next(table for sheet in app.workbook.sheets_ for table in sheet.pivot_tables
                 if table.root.get('name') == 'First')
    first.cache.set_enable_refresh(enabled)
    saved = app.save(tmp_path / 'metadata-edited.xlsm')
    with ZipFile(saved) as package:
        assert {name: package.read(name) for name in records} == records
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='RefreshRead')
    again = next(table for sheet in reopened.workbook.sheets_ for table in sheet.pivot_tables
                 if table.root.get('name') == 'First')
    assert bool(again.cache.EnableRefresh()) is enabled
    assert reopened.run('Snapshot') == case['reopened']


@pytest.mark.parametrize('case', EVENTS, ids=lambda case: case['name'])
def test_generated_refresh_save_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(case)
    assert app.run('Probe') == case['raw_result']
    saved = app.save(tmp_path / 'generated.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='RefreshRead')
    assert reopened.run('Snapshot') == case['reopened']


def refresh_state(app: ExcelApplication) -> object:
    book = app.workbook
    return (book.saved, tuple((id(cache), id(cache.root), cache.index, cache.cache_id, cache.part,
                              cache.changed, xml_bytes(cache.root),
                              xml_bytes(cache.records) if cache.records is not None else None)
                             for cache in book.pivot_caches),
            tuple((sheet.dirty, tuple((key, cell.value, cell.formula, cell.style, cell.stale)
                                     for key, cell in sheet.cells_.items()),
                   tuple((id(table), id(table.root), xml_bytes(table.root), table.changed)
                         for table in sheet.pivot_tables)) for sheet in book.sheets_),
            book.stylesheet.dxf_count)


@pytest.mark.parametrize('case', SCHEMA, ids=lambda case: case['name'])
def test_native_schema_changes_are_retained_as_deferred_model_admission(case: dict[str, Any]) -> None:
    # Native controls rebuild fields/drop measures and invalidate held proxies.
    # This increment deliberately refuses those changes, without claiming equality.
    assert ';first=completed;' in case['raw_result']
    prefix, marker, _rest = case['source'].partition('stage="first-action"\n')
    assert marker
    before_probe, probe, body = prefix.partition('Function Probe() As String\n')
    assert probe
    source = before_probe + probe + body.replace('On Error GoTo Failed\n', 'On Error GoTo 0\n', 1)
    source += 'Probe=s\nEnd Function\n'
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='DeferredSchemaSetup')
    app.run('Probe')
    before = refresh_state(app)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.pivot_caches[0].Refresh()
    assert refresh_state(app) == before


def test_handler_error_control_is_diagnostic_without_a_state_oracle_claim() -> None:
    case = next(case for case in RECORDS['events_save.json']['probes'] if case['name'] == 'cache-handler-error')
    assert MANIFEST['cases']['events_save.json:cache-handler-error']['admission'] == 'handler-error-diagnostic'
    assert case['native_status']['ok'] is False
    assert 'file' not in case and 'reopened' not in case


@pytest.mark.parametrize('case', IMPORTED, ids=lambda case: case['name'])
def test_native_post_import_refresh_and_generated_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    assert sha256(case['source'].encode()).hexdigest() == MANIFEST['cases'][
        'imported.json:' + case['name']]['source_sha256']
    app = ExcelApplication.open(ROOT / PureWindowsPath(case['input_file']).name, with_vba=False)
    app.add_module(case['inspection_source'], name='RefreshProbe')
    app.add_module(case['source'], name='ImportedRefresh')
    sheet = app.workbook.sheets_[0]
    first = next(table for table in sheet.pivot_tables if table.root.get('name') == 'First')
    unrelated = next(table for table in sheet.pivot_tables if table.root.get('name') == 'Unrelated')
    field = first.PivotFields('Amount')
    roots = first.root, first.cache.root, field.root
    unrelated_before = xml_bytes(unrelated.root), xml_bytes(unrelated.cache.root)
    cache_id, part = first.cache.cache_id, first.cache.part
    assert app.run('Probe') == case['raw_result']
    assert roots == (first.root, first.cache.root, field.root)
    assert (cache_id, part) == (first.cache.cache_id, first.cache.part)
    assert field.Name() == 'Amount'
    assert (xml_bytes(unrelated.root), xml_bytes(unrelated.cache.root)) == unrelated_before
    saved = app.save(tmp_path / 'imported-refreshed.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='RefreshProbe')
    assert reopened.run('Snapshot') == case['reopened']


@pytest.mark.parametrize('case', IMPORTED, ids=lambda case: case['name'])
def test_native_post_import_refreshed_input_rewrite(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / PureWindowsPath(case['file']).name, with_vba=False)
    app.add_module(case['inspection_source'], name='RefreshProbe')
    assert app.run('Snapshot') == case['reopened']
    first = next(table for sheet in app.workbook.sheets_ for table in sheet.pivot_tables
                 if table.root.get('name') == 'First')
    first.cache.set_enable_refresh(True)
    saved = app.save(tmp_path / 'refreshed-input-rewritten.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='RefreshProbe')
    assert reopened.run('Snapshot') == case['reopened']


@pytest.fixture(scope='module')
def native_pivot_refresh_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    # One session serves all seven controlled files; context cleanup closes Excel.
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native Excel generated-file gate')
@pytest.mark.parametrize('kind,case', [('generated', case) for case in EVENTS] +
                         [('imported', case) for case in IMPORTED],
                         ids=[case['name'] for case in EVENTS] + [case['name'] for case in IMPORTED])
def test_generated_pivot_refresh_native_excel_reopen(
        kind: str, case: dict[str, Any], tmp_path: Path, native_pivot_refresh_excel: Any) -> None:
    if kind == 'generated':
        app = configured(case)
    else:
        app = ExcelApplication.open(ROOT / PureWindowsPath(case['input_file']).name, with_vba=False)
        app.add_module(case['inspection_source'], name='RefreshProbe')
        app.add_module(case['source'], name='ImportedRefresh')
    assert app.run('Probe') == case['raw_result']
    saved = app.save(tmp_path / 'generated-native-reopen.xlsm')
    native_pivot_refresh_excel.open_document(saved)
    # The persisted module is controlled source; no native code/line injection.
    assert native_pivot_refresh_excel.run_raw('RefreshProbe.Snapshot') == case['reopened']


@pytest.mark.parametrize('case', COLLISIONS, ids=lambda case: case['name'])
def test_native_collision_error_state_is_diagnostic_and_model_refuses_atomically(case: dict[str, Any]) -> None:
    # Native completed without a raised error but produced a collapsed Error cell.
    # Its raw classifier is retained; the explicit manifest disposition governs admission.
    assert case['native_status']['ok'] is True
    assert ';refresh=completed;' in case['raw_result'] and ':Error:Error 2045|' in case['raw_result']
    assert MANIFEST['cases']['growth.json:' + case['name']]['admission'] == 'native-error-diagnostic-model-refusal'
    prefix, marker, _rest = case['source'].partition('stage="refresh"\n')
    assert marker
    helpers, declaration, body = prefix.partition('Function Probe() As String\n')
    assert declaration
    source = helpers + declaration + body.replace('On Error GoTo Failed\n', 'On Error GoTo 0\n', 1)
    source += 'Probe=s\nEnd Function\n'
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='CollisionSetup')
    app.run('Probe')
    before = refresh_state(app)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.pivot_caches[0].Refresh()
    assert refresh_state(app) == before
