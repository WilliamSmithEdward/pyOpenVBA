"""Owned filter/print-name migration and inactive AdvancedFilter copy admission."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from hashlib import sha256
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba._a1 import split_sheet
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._model import Range
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/owned_name_rename'
RECORD: dict[str, Any] = json.loads((ROOT / 'native.json').read_text(encoding='utf8'))
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text(encoding='utf8'))
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_owned_name_rename_matches_native_and_persists(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='RenameControl')
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / 'renamed.xlsm')
    reopened = ExcelApplication.open(saved)
    assert reopened.run('Snapshot') == case['reopened']
    with ZipFile(saved) as package:
        root = xml_tree(package.read('xl/workbook.xml'))
    names = root.findall(f'{NS}definedNames/{NS}definedName')
    assert all(entry.get('localSheetId') == '0' for entry in names)
    assert all(entry.get('name', '').startswith('_xlnm.') for entry in names)
    assert all("'O''Brien renamed'!" in (entry.text or '') for entry in names)
    imported = ExcelApplication.open(ROOT / MANIFEST['samples'][case['name']]['file'], with_vba=False)
    imported.add_module(case['inspection_source'], name='RenameRead')
    assert imported.run('Snapshot') == case['reopened']


def test_native_name_rename_evidence_is_byte_exact_and_macrofree() -> None:
    assert sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['raw_sha256']
    assert RECORD['completed'] and RECORD['closed_utc']
    for case in RECORD['probes']:
        assert case['status']['ok'] and not case['status']['dialogs']
        assert case['save_status']['ok'] and not case['save_status']['dialogs']
        assert case['reopen_status']['ok'] and not case['reopen_status']['dialogs']
        sample = MANIFEST['samples'][case['name']]
        path = ROOT / sample['file']
        assert sha256(path.read_bytes()).hexdigest() == sample['sha256']
        with ZipFile(path) as package:
            assert not any('vbaproject' in name.casefold() for name in package.namelist())


def test_malformed_later_filter_name_leaves_print_and_filter_names_unchanged() -> None:
    app = ExcelApplication()
    app.add_workbook()
    case = next(case for case in RECORD['probes'] if case['name'] == 'extract')
    app.add_module(case['source'], name='RenameControl')
    assert app.run('Probe') == case['result']
    sheet = app.workbook.sheets_[0]
    entry = next(entry for entry in app.workbook.names_.entries if split_sheet(entry.name)[1] == 'Extract')
    entry.refers_to = '=D1+"unterminated'
    before = sheet.name, [(id(entry), entry.name, entry.refers_to, entry.attributes) for entry in app.workbook.names_.entries], app.workbook.names_.changed, app.workbook.saved
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Another name')
    assert (sheet.name, [(id(entry), entry.name, entry.refers_to, entry.attributes) for entry in app.workbook.names_.entries], app.workbook.names_.changed, app.workbook.saved) == before


def filter_prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Prepare()
Range("A1:B1").Value=Array("Name","Qty")
Range("A2:B2").Value=Array("apple",1)
Range("A3:B3").Value=Array("apple",1)
Range("D1:E1").Value=Array("Name","Qty")
Range("G1").Value="Name"
Range("G2").Value="a*"
End Sub''', name='FilterSetup')
    app.run('Prepare')
    return app


@pytest.mark.parametrize('kind', ['extract', 'showall', 'unique-showall'])
@pytest.mark.parametrize('reopen', [False, True])
def test_retained_filter_metadata_copy_is_atomic(kind: str, reopen: bool, tmp_path: Path) -> None:
    app = filter_prepared()
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    if kind == 'extract':
        source.AdvancedFilter(2, sheet.Range('G1:G2'), sheet.Range('D1:E1'))
    elif kind == 'showall':
        source.AdvancedFilter(1, sheet.Range('G1:G2'))
        sheet.ShowAllData()
    else:
        source.AdvancedFilter(1, Unique=True)
        sheet.ShowAllData()
    assert sheet.advanced_filter is None
    if reopen:
        app = ExcelApplication.open(app.save(tmp_path / 'retained.xlsm'), with_vba=False)
        sheet = app.workbook.sheets_[0]
    before = len(app.workbook.application.workbooks_.books), tuple(app.workbook.sheets_), app.workbook.active_sheet, [(entry.name, entry.refers_to, entry.attributes) for entry in app.workbook.names_.entries]
    with pytest.raises(VBAUnsupportedError):
        sheet.Copy(After=sheet)
    assert (len(app.workbook.application.workbooks_.books), tuple(app.workbook.sheets_), app.workbook.active_sheet, [(entry.name, entry.refers_to, entry.attributes) for entry in app.workbook.names_.entries]) == before


def test_ordinary_autofilter_copy_remains_supported() -> None:
    app = filter_prepared()
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    source.AutoFilter(1, 'apple')
    sheet.Copy(After=sheet)
    assert len(app.workbook.sheets_) == 2
    assert sheet.auto_filter is not None
    assert app.workbook.sheets_[1].cells_[2, 1].value == 'apple'



def test_owned_name_rename_preserves_unrelated_names_and_opaque_attributes(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._advanced_filter import attribute_values

    app = filter_prepared()
    app.add_module('''Sub ExtraNames()
ActiveSheet.Names.Add Name:="UserLocal",RefersTo:="=1"
ActiveWorkbook.Names.Add Name:="Criteria",RefersTo:="=2"
End Sub''', name='OrdinaryNames')
    app.run('ExtraNames')
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    source.AdvancedFilter(2, sheet.Range('G1:G2'), sheet.Range('D1:E1'))
    unrelated = [(entry, entry.name, entry.refers_to, entry.attributes) for entry in app.workbook.names_.entries
                 if split_sheet(entry.name)[1] == 'UserLocal' or not split_sheet(entry.name)[0]]
    for entry in app.workbook.names_.entries:
        if split_sheet(entry.name)[1] in ('Criteria', 'Extract') and split_sheet(entry.name)[0]:
            entry.attributes += ' xmlns:v="urn:vendor" v:opaque="value > &quot;quoted&quot;"'
    sheet.vba_set('Name', "O'Brien renamed")
    for entry, name, refers, attrs in unrelated:
        owner, bare = split_sheet(name)
        expected = "'O''Brien renamed'!" + bare if owner else name
        assert (entry.name, entry.refers_to, entry.attributes) == (expected, refers, attrs)
    # Ordinary local scope now migrates too; keep it through persistence rather
    # than removing the intentionally deferred owner as the older test did.
    reopened = ExcelApplication.open(app.save(tmp_path / 'opaque-renamed.xlsm'), with_vba=False)
    assert any(entry.name == "'O''Brien renamed'!UserLocal" for entry in reopened.workbook.names_.entries)
    assert not any(entry.name == 'UserLocal' for entry in reopened.workbook.names_.entries)
    actual = [attribute_values(entry.attributes) for entry in reopened.workbook.names_.entries
              if split_sheet(entry.name)[1] in ('Criteria', 'Extract') and split_sheet(entry.name)[0]]
    assert len(actual) == 2
    assert all(attrs['xmlns:v'] == 'urn:vendor' and attrs['v:opaque'] == 'value > "quoted"' for attrs in actual)



@pytest.fixture(scope='module')
def native_owned_name_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native owned-name persistence gate')
@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_engine_generated_owned_rename_reopened_by_native(
        case: dict[str, Any], tmp_path: Path, native_owned_name_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='RenameControl')
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / (case['name'] + '-renamed.xlsm'))
    native_owned_name_excel.open_document(saved)
    result = native_owned_name_excel.run_vba(case['inspection_source'], 'Snapshot', module_name='NativeNameRead',
                                             line_numbers=False, timeout=30)
    assert result.ok and not result.dialogs, result
    assert result.value == case['reopened']
