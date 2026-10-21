"""Independent native Table.Sort runtime and macrofree-package oracles.

All raw controls remain unchanged. Eleven complete runtime controls are admitted;
eight whole controls remain deferred. Reading an imported result does not admit
the mutation that originally produced that result. No expected strings are
computed from the model and no source/result projections are applied.
"""
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
from pyopenvba.apps.excel._charts_file import xml_tree

ROOT = Path(__file__).parent / 'fixtures/table_sort'
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf-8'))
RECORDS: dict[str, Any] = {
    name: json.loads((ROOT / name).read_text(encoding='utf-8'))
    for name in MANIFEST['raw_files']
}
CASES = [(filename, case) for filename, record in RECORDS.items() for case in record['probes']]
ADMITTED = [(filename, case) for filename, case in CASES
            if MANIFEST['cases'][filename + ':' + case['name']]['runtime'] == 'admitted-full']
INPUTS = [(filename, case) for filename, case in CASES if 'file' in case]
GENERATED = [(filename, case) for filename, case in ADMITTED if 'file' in case]


def configured(filename: str, case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    for name, source in RECORDS[filename].get('classes', {}).items():
        app.add_module(source, name=name, kind='class')
    app.add_module(case['source'], name='SortProbe')
    return app


@pytest.mark.parametrize('filename', RECORDS)
def test_native_table_sort_record_bytes_and_admission(filename: str) -> None:
    record = RECORDS[filename]
    assert sha256((ROOT / filename).read_bytes()).hexdigest() == MANIFEST['raw_files'][filename]
    assert record['completed'] and record['closed_utc']
    for case in record['probes']:
        entry = MANIFEST['cases'][filename + ':' + case['name']]
        assert entry['runtime'] in ('admitted-full', 'deferred-whole')
        assert case['status']['ok'] and not case['status']['dialogs']
        assert sha256(case['source'].encode()).hexdigest() == entry['source_sha256']
        assert sha256(case['result'].encode()).hexdigest() == entry['result_sha256']


@pytest.mark.parametrize('filename,case', ADMITTED, ids=[case['name'] for _, case in ADMITTED])
def test_native_table_sort_complete_runtime(filename: str, case: dict[str, Any]) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('filename,case', GENERATED, ids=[case['name'] for _, case in GENERATED])
def test_generated_table_sort_save_reopen_matches_native(
        filename: str, case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / 'generated-sort.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='SortRead')
    assert reopened.run('Snapshot') == case['reopened']


@pytest.mark.parametrize('filename,case', INPUTS, ids=[case['name'] for _, case in INPUTS])
def test_native_macrofree_table_sort_input_matches_snapshot(filename: str, case: dict[str, Any]) -> None:
    entry = MANIFEST['cases'][filename + ':' + case['name']]
    sample = ROOT / PureWindowsPath(case['file']).name
    assert sample.name == entry['sample']
    assert sha256(sample.read_bytes()).hexdigest() == entry['sample_sha256']
    assert case['save_status']['ok'] and not case['save_status']['dialogs']
    assert case['reopen_status']['ok'] and not case['reopen_status']['dialogs']
    assert sha256(case['reopened'].encode()).hexdigest() == entry['reopened_sha256']
    assert sha256(case['inspection_source'].encode()).hexdigest() == entry['inspection_source_sha256']
    app = ExcelApplication.open(sample, with_vba=False)
    app.add_module(case['inspection_source'], name='SortRead')
    assert app.run('Snapshot') == case['reopened']


@pytest.mark.parametrize('filename,case', INPUTS, ids=[case['name'] for _, case in INPUTS])
def test_imported_table_sort_public_metadata_rewrite_preserves_native_snapshot(
        filename: str, case: dict[str, Any], tmp_path: Path) -> None:
    sample = ROOT / PureWindowsPath(case['file']).name
    app = ExcelApplication.open(sample, with_vba=False)
    app.add_module(case['inspection_source'], name='SortRead')
    assert app.run('Snapshot') == case['reopened']
    app.add_module('''Sub RewriteTableMetadata()
Dim t As ListObject
For Each t In Worksheets("Data").ListObjects
 t.Name=t.Name
Next
End Sub''', name='MetadataRewrite')
    app.run('RewriteTableMetadata')
    assert all(table.changed for table in app.workbook.sheets_[0].tables)
    saved = app.save(tmp_path / 'imported-sort-rewritten.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='SortRead')
    assert reopened.run('Snapshot') == case['reopened']


@pytest.mark.parametrize('filename,case', INPUTS, ids=[case['name'] for _, case in INPUTS])
def test_recorded_table_sort_packages_are_macrofree_and_keep_native_metadata(
        filename: str, case: dict[str, Any]) -> None:
    sample = ROOT / PureWindowsPath(case['file']).name
    with ZipFile(sample) as package:
        assert not any('vbaproject' in part.casefold() for part in package.namelist())
        for part, original_xml in case['xml'].items():
            assert package.read(part).decode('utf-8') == original_xml
        for part in package.namelist():
            if part.endswith('.rels'):
                relations = xml_tree(package.read(part))
                assert all(node.get('TargetMode', '').casefold() != 'external' for node in relations)


@pytest.fixture(scope='module')
def native_table_sort_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native Excel generated table-sort gate')
@pytest.mark.parametrize('filename,case', GENERATED, ids=[case['name'] for _, case in GENERATED])
def test_generated_table_sort_native_excel_reopen(
        filename: str, case: dict[str, Any], tmp_path: Path, native_table_sort_excel: Any) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / 'native-generated-sort.xlsm')
    native_table_sort_excel.open_document(saved)
    assert native_table_sort_excel.run_raw('SortProbe.Snapshot') == case['reopened']


POST_APPLY: dict[str, Any] = json.loads((ROOT / 'post_apply/native.json').read_text(encoding='utf8'))


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native Excel post-Apply table-sort gate')
@pytest.mark.parametrize('case', POST_APPLY['probes'], ids=lambda case: case['name'])
def test_post_apply_table_sort_native_excel_reopen(
        case: dict[str, Any], tmp_path: Path, native_table_sort_excel: Any) -> None:
    if case['loaded']:
        app = ExcelApplication.open(ROOT / PureWindowsPath(case['input_file']).name, with_vba=False)
    else:
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(case['setup_source'], name='Setup')
        assert app.run('Setup.Probe') == case['setup']['value']
    app.add_module(case['source'], name='Mutation')
    assert app.run('Mutation.Probe') == case['result']['value']
    saved = app.save(tmp_path / 'native-post-apply-sort.xlsm')
    native_table_sort_excel.open_document(saved)
    assert native_table_sort_excel.run_raw('Mutation.Snapshot') == case['reopened']['value']
