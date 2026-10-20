"""Native persisted worksheet PageSetup configuration and explicit scope guards."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._page_setup import PageSetup, with_setup
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL

ROOT = Path(__file__).parent / 'fixtures/page_setup'
RECORD: dict[str, Any] = json.loads((ROOT / 'configuration.json').read_text(encoding='utf8'))
CONTROLS = [case for case in RECORD['controls'] if case['runtime_admitted']]
SAMPLES = RECORD['samples']
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


@pytest.mark.parametrize('case', CONTROLS, ids=lambda case: case['name'])
def test_native_configuration_control(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='ConfigurationControl')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_native_configuration_import_and_forced_rewrite(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['fixture_file'], with_vba=False)
    app.add_module(case['inspection_source'], name='ConfigurationRead')
    assert app.run('ConfigurationState') == case['reopened']
    app.workbook.sheets_[0].touched()
    result = app.save(tmp_path / 'rewritten.xlsm')
    with ZipFile(ROOT / case['fixture_file']) as original, ZipFile(result) as rewritten:
        for part in original.namelist():
            if part.startswith('xl/printerSettings/'):
                assert original.read(part) == rewritten.read(part)
        rel = 'xl/worksheets/_rels/sheet1.xml.rels'
        if rel in original.namelist():
            assert original.read(rel) == rewritten.read(rel)
    assert ExcelApplication.open(result).run('ConfigurationState') == case['reopened']


@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_generated_configuration_save_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    app = generated(case)
    result = app.save(tmp_path / 'generated.xlsm')
    assert ExcelApplication.open(result).run('ConfigurationState') == case['reopened']


def generated(case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['generated_source'], name='GeneratedConfiguration')
    app.run('Generate' if case.get('generator_is_derivative') else 'Probe')
    assert app.run('ConfigurationState') == case['reopened']
    return app


@pytest.fixture(scope='module')
def native_configuration_excel() -> Any:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native configuration gate')
@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_generated_configuration_reopened_by_native(case: dict[str, Any], tmp_path: Path,
                                                    native_configuration_excel: Any) -> None:
    app = generated(case)
    path = app.save(tmp_path / 'native-generated.xlsm')
    native_configuration_excel.open_document(path)
    assert native_configuration_excel.run_raw('GeneratedConfiguration.ConfigurationState') == case['reopened']


def sheet_state(app: ExcelApplication) -> object:
    sheet = app.workbook.sheets_[0]
    current = sheet.page_setup_state
    return (sheet.dirty, app.workbook.saved,
            None if current is None else (dict(current.attributes), dict(current.changes)),
            [(entry.name, entry.refers_to, entry.visible, entry.attributes) for entry in app.workbook.names_.entries],
            frozenset(sheet.page_print_names_owned))


@pytest.mark.parametrize('value', [-1, -18.5, True, False, NULL, '18.5', float('inf'), float('nan')])
def test_unadmitted_margin_values_fail_before_mutation(value: object) -> None:
    app = ExcelApplication()
    app.add_workbook()
    page = PageSetup(app.workbook.sheets_[0])
    page.LeftMargin()  # Lazy read state is established before the snapshot.
    before = sheet_state(app)
    with pytest.raises(VBAUnsupportedError):
        page.vba_set('LeftMargin', value)
    assert sheet_state(app) == before


def test_empty_margin_is_common_zero_and_points_are_preserved(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    page = PageSetup(app.workbook.sheets_[0])
    page.vba_set('LeftMargin', EMPTY)
    page.vba_set('RightMargin', 18.125)
    assert page.LeftMargin() == 0 and page.RightMargin() == 18.125
    result = app.save(tmp_path / 'fractional.xlsm')
    restored = PageSetup(ExcelApplication.open(result).workbook.sheets_[0])
    assert restored.LeftMargin() == 0 and restored.RightMargin() == 18.125


@pytest.mark.parametrize('kind', ['PrintArea', 'PrintTitleRows', 'PrintTitleColumns'])
def test_print_names_reject_foreign_or_unmeasured_references_atomically(kind: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    page = PageSetup(app.workbook.sheets_[0])
    before = sheet_state(app)
    with pytest.raises(VBAUnsupportedError):
        page.vba_set(kind, 'Other!A1:B2')
    assert sheet_state(app) == before


def test_print_area_owned_name_update_is_idempotent() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    page = PageSetup(sheet)
    page.vba_set('PrintArea', 'A1:B9')
    entry = app.workbook.names_.entries[0]
    attrs = entry.attributes
    for _ in range(20):
        page.vba_set('PrintArea', 'A1:B9')
    assert app.workbook.names_.entries == [entry] and entry.attributes == attrs


@pytest.mark.parametrize('prefix', ['', 's:'])
def test_narrow_xml_writes_preserve_opaque_attributes_and_nested_metadata(prefix: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    page = PageSetup(sheet)
    page.vba_set('LeftMargin', 18)
    page.vba_set('Zoom', False)
    declaration = 'xmlns:s' if prefix else 'xmlns'
    xml = (f'<{prefix}worksheet {declaration}="{NS[1:-1]}" xmlns:r="urn:relationship" xmlns:x="urn:extra">'
           f'<{prefix}sheetPr x:opaque=\'keep fitToPage="0" > value\'><{prefix}outlinePr summaryBelow="0"/>'
           f'<{prefix}pageSetUpPr x:opaque=\'fitToPage="0"\' autoPageBreaks="0"/></{prefix}sheetPr>'
           f'<{prefix}sheetData/><{prefix}pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" '
           f'header="0.3" footer="0.3" x:opaque=\'left="99"\'/>'
           f'<{prefix}pageSetup r:id="printer1" x:opaque="untouched"/>'
           f'<{prefix}headerFooter><{prefix}oddHeader>&amp;RPrivate</{prefix}oddHeader></{prefix}headerFooter>'
           f'</{prefix}worksheet>')
    result = with_setup(sheet, xml)
    root = xml_tree(result)
    margins = root.find(NS + 'pageMargins')
    assert margins is not None and margins.get('left') == '0.25' and margins.get('{urn:extra}opaque') == 'left="99"'
    assert f'<{prefix}outlinePr summaryBelow="0"/>' in result
    assert f'<{prefix}pageSetup r:id="printer1" x:opaque="untouched"/>' in result
    assert f'<{prefix}headerFooter><{prefix}oddHeader>&amp;RPrivate</{prefix}oddHeader></{prefix}headerFooter>' in result
    nested = root.find(NS + 'sheetPr/' + NS + 'pageSetUpPr')
    assert nested is not None and nested.get('fitToPage') == '1' and nested.get('autoPageBreaks') == '0'
    assert nested.get('{urn:extra}opaque') == 'fitToPage="0"'


def test_fixture_projection_and_macrofree_package_provenance() -> None:
    manifest = json.loads((ROOT / 'provenance.json').read_text())
    assert manifest['projection']['excluded_environment_fields'] == ['Application.ActivePrinter']
    for name, digest in manifest['files'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == digest
    for case in SAMPLES:
        with ZipFile(ROOT / case['fixture_file']) as package:
            assert not any('vbaProject' in part or 'externalLink' in part for part in package.namelist())


LIFECYCLE = json.loads((ROOT / 'name_lifecycle.json').read_text(encoding='utf8'))['controls']


@pytest.mark.parametrize('case', LIFECYCLE, ids=lambda case: case['name'])
def test_native_print_name_lifecycle_and_generated_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='PageNameLifecycle')
    assert app.run('Probe') == case['result']
    result = app.save(tmp_path / 'name-lifecycle.xlsm')
    assert ExcelApplication.open(result).run('ConfigurationState') == case['reopened']
    imported = ExcelApplication.open(ROOT / case['fixture_file'], with_vba=False)
    imported.add_module(case['inspection_source'], name='ReadLifecycle')
    assert imported.run('ConfigurationState') == case['reopened']


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native configuration gate')
@pytest.mark.parametrize('case', LIFECYCLE, ids=lambda case: case['name'])
def test_generated_print_name_lifecycle_reopened_by_native(case: dict[str, Any], tmp_path: Path,
                                                          native_configuration_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='PageNameLifecycle')
    assert app.run('Probe') == case['result']
    path = app.save(tmp_path / 'native-lifecycle.xlsm')
    native_configuration_excel.open_document(path)
    assert native_configuration_excel.run_raw('PageNameLifecycle.ConfigurationState') == case['reopened']


def test_configuration_reads_cache_bounded_metadata_and_setters_validate_only_their_paths(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _page_setup

    app = ExcelApplication.open(ROOT / 'margins-points.xlsx', with_vba=False)
    page = PageSetup(app.workbook.sheets_[0])
    original = _page_setup.opc_part
    calls: list[tuple[str, ...] | None] = []

    def counted(data: bytes, namespace: str, root_name: str, **kwargs: Any) -> Any:
        calls.append(kwargs.get('container_path'))
        return original(data, namespace, root_name, **kwargs)

    monkeypatch.setattr(_page_setup, 'opc_part', counted)
    for _ in range(100):
        assert page.LeftMargin() == 18
        assert page.Orientation() == 1
    assert calls == [None, ('sheetPr',)]
    page.vba_set('LeftMargin', 36)
    assert calls[-1] == () and len(calls) == 3
    page.vba_set('Zoom', False)
    assert calls[-1] == ('sheetPr',) and len(calls) == 4


def test_configuration_only_edits_preserve_raw_sheet_data_and_repeat_save_bytes(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'margins-points.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    package = app.workbook.package
    assert package is not None
    xml = package.read(sheet.part_name).decode('utf8')
    opaque = '<sheetData><row r="8" spans="1:1" xmlns:x="urn:row" x:opaque="retain"><c r="A8" x:opaque="retain"><v>23</v></c></row></sheetData>'
    xml = xml.replace('<sheetData/>', opaque)
    assert opaque in xml
    package.write(sheet.part_name, xml.encode('utf8'))
    page = PageSetup(sheet)
    page.vba_set('LeftMargin', 36)
    page.vba_set('Zoom', False)
    page.vba_set('PrintArea', 'A1:B9')
    assert not sheet.dirty
    first = app.save(tmp_path / 'configuration-only.xlsm')
    with ZipFile(first) as saved:
        first_xml = saved.read(sheet.part_name)
        assert opaque.encode() in first_xml
        printer_parts = {name: saved.read(name) for name in saved.namelist() if name.startswith('xl/printerSettings/')}
    second = app.save(tmp_path / 'repeat.xlsm')
    with ZipFile(second) as saved:
        assert saved.read(sheet.part_name) == first_xml
        assert all(saved.read(name) == value for name, value in printer_parts.items())


@pytest.mark.parametrize('property_name,value', [('Zoom', 75), ('Zoom', False), ('FitToPagesWide', 2),
                                              ('FirstPageNumber', 3), ('BlackAndWhite', True),
                                              ('Orientation', 2)])
def test_new_page_setup_preserves_orientation_while_materializing_other_settings(
        tmp_path: Path, property_name: str, value: object) -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    page = PageSetup(sheet)
    assert page.Orientation() == 1
    page.vba_set(property_name, value)
    expected = 'landscape' if property_name == 'Orientation' else 'portrait'
    assert page.Orientation() == (2 if property_name == 'Orientation' else 1)
    first = app.save(tmp_path / 'first.xlsm')
    with ZipFile(first) as package:
        root = xml_tree(package.read(sheet.part_name))
        element = root.find(NS + 'pageSetup')
        if property_name == 'Zoom' and value is False:
            assert element is None  # fitToPage alone lives in sheetPr.
        else:
            assert element is not None and element.get('orientation') == expected
    page.vba_set('PrintErrors', 2)
    second = app.save(tmp_path / 'second.xlsm')
    with ZipFile(second) as package:
        root = xml_tree(package.read(sheet.part_name))
        element = root.find(NS + 'pageSetup')
        assert element is not None and element.get('orientation') == expected


@pytest.mark.parametrize('stored_orientation', [None, 'default'])
def test_imported_page_setup_missing_or_unmeasured_orientation_is_not_reported_as_portrait(
        stored_orientation: str | None) -> None:
    app = ExcelApplication.open(ROOT / 'margins-points.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    package = app.workbook.package
    assert package is not None
    attribute = '' if stored_orientation is None else ' orientation="default"'
    xml = f'<worksheet xmlns="{NS[1:-1]}"><sheetData/><pageSetup scale="75"{attribute}/></worksheet>'
    package.write(sheet.part_name, xml.encode())
    page = PageSetup(sheet)
    if stored_orientation is None:
        # The five independently reopened generated controls reported Double 0.
        assert page.Orientation() == 0.0
        page.vba_set('LeftMargin', 36)
        assert page.Orientation() == 0.0
    else:
        with pytest.raises(VBAUnsupportedError, match='stored enumeration'):
            page.Orientation()


@pytest.mark.parametrize('prefix', ['', 's:'])
def test_foreign_no_namespace_configuration_is_not_owned_and_survives_writes(prefix: str) -> None:
    app = ExcelApplication.open(ROOT / 'margins-points.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    package = app.workbook.package
    assert package is not None
    declaration = 'xmlns:s' if prefix else 'xmlns'
    foreign = ('<pageSetup xmlns="" orientation="landscape" opaque="retain-setup"/>'
               '<printOptions xmlns="" horizontalCentered="1" opaque="retain-options"/>'
               '<pageMargins xmlns="" left="9" opaque="retain-margins"/>')
    xml = (f'<{prefix}worksheet {declaration}="{NS[1:-1]}"><{prefix}sheetData/>'
           + foreign + f'</{prefix}worksheet>')
    package.write(sheet.part_name, xml.encode())
    page = PageSetup(sheet)
    assert page.Orientation() == 1
    assert page.CenterHorizontally() is False
    assert page.LeftMargin() == 50.4
    page.vba_set('LeftMargin', 18)
    page.vba_set('PrintHeadings', True)
    page.vba_set('FirstPageNumber', 3)
    result = with_setup(sheet, xml)
    assert foreign in result
    root = xml_tree(result)
    owned_setup = root.find(NS + 'pageSetup')
    owned_options = root.find(NS + 'printOptions')
    owned_margins = root.find(NS + 'pageMargins')
    assert owned_setup is not None and owned_setup.get('orientation') == 'portrait'
    assert owned_options is not None and owned_options.get('headings') == '1'
    assert owned_margins is not None and owned_margins.get('left') == '0.25'
