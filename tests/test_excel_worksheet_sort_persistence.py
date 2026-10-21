"""Exact native worksheet-sort persistence controls and bounded import refusal."""
from __future__ import annotations

import json
from hashlib import sha256
from pathlib import Path, PureWindowsPath
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._sort import SortObject
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/table_sort/persistence'
RECORD: dict[str, Any] = json.loads((ROOT / 'native.json').read_text(encoding='utf8'))
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_worksheet_configuration_and_save(case: dict[str, Any], tmp_path: Path) -> None:
    if case['loaded']:
        app = ExcelApplication.open(ROOT / PureWindowsPath(case['input_file']).name, with_vba=False)
        app.add_module(case['inspection_source'], name='ReadBefore')
        assert app.run('SortSnapshot') == case['loaded_before']['value']
    else:
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(case['setup_source'], name='Setup')
        assert app.run('Prepare') == case['setup']['value']
    app.add_module(case['source'], name='Mutation')
    assert app.run('Mutation.Probe') == case['result']['value']
    if MANIFEST['cases'][case['name']]['admission'] != 'runtime-and-reopen':
        with pytest.raises(VBAUnsupportedError, match='methods/orientations'):
            app.save(tmp_path / 'unsupported.xlsm')
        return
    saved = app.save(tmp_path / 'configuration.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case['inspection_source'], name='Inspect')
    assert reopened.run('SortSnapshot') == case['reopened']['value']


def test_raw_and_macrofree_package_provenance() -> None:
    assert RECORD['completed'] and RECORD['closed_utc']
    assert sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['raw_sha256']
    for case in RECORD['probes']:
        assert sha256(case['source'].encode()).hexdigest() == MANIFEST['cases'][case['name']]['source_sha256']
        for phase in ('setup', 'result', 'save_status', 'reopened'):
            assert case[phase]['ok'] and not case[phase]['dialogs']
        if case['loaded']:
            assert case['loaded_before']['ok'] and not case['loaded_before']['dialogs']
    for name, digest in MANIFEST['samples'].items():
        assert sha256((ROOT / name).read_bytes()).hexdigest() == digest
        with ZipFile(ROOT / name) as package:
            assert not any('vbaproject' in part.casefold() or 'embeddings/' in part for part in package.namelist())


@pytest.mark.parametrize('attrs,child', [
    ('custom="opaque"', ''), ('sortMethod="stroke"', ''), ('columnSort="1"', ''),
    ('caseSensitive="invalid"', ''), ('', '<sortCondition ref="J2:J5" sortBy="cellColor"/>'),
    ('', '<sortCondition ref="J2:J5"><extLst/></sortCondition>'),
    ('', '<sortCondition ref="Other!J2:J5"/>'),
    ('', '<sortCondition ref="J2:J5"/><!--opaque-->'),
    ('', '<sortCondition ref="J2:J5"/><?opaque retain?>'),
    ('', '<sortCondition ref="J2:J5"><![CDATA[ ]]></sortCondition>'),
])
def test_unknown_loaded_worksheet_sort_is_refused_without_hydrating_or_rewriting(attrs: str, child: str) -> None:
    app = ExcelApplication.open(ROOT / 'loaded-options-before.xlsx', with_vba=False)
    sheet, package = app.workbook.sheets_[0], app.workbook.package
    assert package is not None
    xml = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>'
           + '<sortState ref="J1:K5" ' + attrs + '>' + (child or '<sortCondition ref="J2:J5"/>')
           + '</sortState></worksheet>')
    package.write(sheet.part_name, xml.encode())
    with pytest.raises(VBAUnsupportedError):
        SortObject(sheet)
    assert sheet.sort_state is None
    assert package.read(sheet.part_name) == xml.encode()


@pytest.mark.parametrize('has_sort_state', [False, True])
def test_large_worksheet_does_not_limit_small_or_absent_sort_metadata(has_sort_state: bool, tmp_path: Path) -> None:
    native = ROOT / 'loaded-field-order-option-before.xlsx'
    derived = tmp_path / 'large-worksheet.xlsx'
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    with ZipFile(native) as source, ZipFile(derived, 'w') as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == 'xl/worksheets/sheet1.xml':
                part = opc_part(data, namespace, 'worksheet')
                if not has_sort_state:
                    spans = [(start, end) for name, _attrs, start, end in part.children
                             if name == namespace + '}sortState']
                    assert len(spans) == 1
                    start, end = spans[0]
                    data = data[:start] + data[end:]
                    part = opc_part(data, namespace, 'worksheet')
                data = part.insert(part.closing_start, b'<!--' + b'x' * 4_100_000 + b'-->')
                assert len(data) > 4_000_000
            target.writestr(info, data)
    app = ExcelApplication.open(derived, with_vba=False)
    sheet = app.workbook.sheets_[0]
    sort = SortObject(sheet)
    assert len(sort.state.fields) == int(has_sort_state)
    if has_sort_state:
        assert sort.state.fields[0].key.first.address() == '$J$2:$J$5'
    else:
        sort.SetRange(sheet.Range('J2:K5'))
        sort.SortFields().Add(sheet.Range('J2:J5'))
    sort.Apply()


@pytest.mark.parametrize('extension', ['xlsx', 'xlsm'])
def test_worksheet_sort_field_persistence_budget_is_atomic(extension: str, tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'loaded-field-order-option-before.xlsx', with_vba=False)
    sheet, package = app.workbook.sheets_[0], app.workbook.package
    assert package is not None
    sort = SortObject(sheet)
    sort.SortFields().Clear()
    sort.SetRange(sheet.Range('J1:K5'))
    for _ in range(65):
        sort.SortFields().Add(sheet.Range('J2:J5'))
    assert len(sort.state.fields) == 65  # Runtime Add remains available.
    sheet.Range('Z1').vba_set('Value', 'pending change')
    before = package.serialize()
    output = tmp_path / ('too-many-fields.' + extension)
    with pytest.raises(VBAUnsupportedError, match='64 worksheet sort fields'):
        app.save(output)
    assert not output.exists()
    assert app.workbook.package is package and package.serialize() == before
    assert len(sort.state.fields) == 65
    assert sheet.Range('Z1').vba_get('Value') == 'pending change'


def test_worksheet_sort_64_fields_roundtrip(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'loaded-field-order-option-before.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    sort = SortObject(sheet)
    sort.SortFields().Clear()
    sort.SetRange(sheet.Range('J1:K5'))
    for _ in range(64):
        sort.SortFields().Add(sheet.Range('J2:J5'))
    saved = app.save(tmp_path / '64-fields.xlsx')
    reopened = ExcelApplication.open(saved, with_vba=False)
    restored = SortObject(reopened.workbook.sheets_[0])
    assert len(restored.state.fields) == 64
    assert all(field.key.first.address() == '$J$2:$J$5' for field in restored.state.fields)
