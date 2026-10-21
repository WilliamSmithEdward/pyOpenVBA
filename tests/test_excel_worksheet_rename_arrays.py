"""Independent native shared/CSE worksheet rename controls and atomic admission."""
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

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import NameEntry
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/worksheet_rename_arrays'
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text())
RECORD: dict[str, Any] = json.loads((ROOT / 'native.json').read_text())
CASES: list[dict[str, Any]] = RECORD['controls']


def context(app: ExcelApplication, case: dict[str, Any]) -> None:
    # CalculateBeforeSave belongs to the native host, not the opened file.
    # Set the independently recorded host context without replacing any oracle.
    app.workbook.application.calculate_before_save = bool(case['configuration'].get('calculate_before_save'))


def configured(case: dict[str, Any], tmp_path: Path) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.workbook.name = MANIFEST['cases'][case['name']]['initial_book_name']
    app.add_module(case.get('clear_source', case['source']), name='RenameArrayControl')
    assert app.run('Prepare') == case['prepare_result']
    if 'seed_package' in case:
        app.save(tmp_path / PureWindowsPath(case['before_clear_package']['file']).name)
        assert app.run('ClearSeed') == case['clear_result']
        saved = app.save(tmp_path / PureWindowsPath(case['seed_package']['file']).name)
        app = ExcelApplication.open(saved, with_vba=False)
        app.add_module(case['source'], name='RenameArrayControl')
    return app


def fingerprint(app: ExcelApplication) -> object:
    book = app.workbook
    return (tuple((id(s), s.name, s.dirty, tuple(s.shared_groups.items()), tuple(s.array_formulas.items()),
                   tuple((position, id(c), c.formula, c.value, c.shared, c.stale, c.save_dirty)
                         for position, c in s.cells_.items())) for s in book.sheets_),
            tuple((id(n), n.name, n.refers_to) for n in book.names_.entries), book.saved, book.names_.changed,
            tuple(book.calculator.compiled), tuple(book.calculator._readers),
            tuple((e.name, e.body) for e in book.package.entries) if book.package else None)


def prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.sheet(1).set_value('A1', 1)
    app.sheet(1).set_value('A2', 2)
    app.workbook.sheets_[0].Range('D1:E2').vba_set('FormulaArray', '=A1:A2*{1,2}')
    app.workbook.application.Calculate()
    return app


def test_native_raw_source_and_package_provenance() -> None:
    assert sha256((ROOT / 'native.json').read_bytes()).hexdigest() == MANIFEST['raw_sha256']
    assert RECORD['completed'] and RECORD['closed_utc']
    for case in CASES:
        assert sha256(case['source'].encode()).hexdigest() == case['source_sha256']
        for key, value in case.items():
            if key == 'status' or key.endswith('_status'):
                assert value['ok'] and not value['dialogs'], (case['name'], key)
    for name, record in MANIFEST['packages'].items():
        path = ROOT / name
        assert sha256(path.read_bytes()).hexdigest() == record['sha256']
        with ZipFile(path) as package:
            assert not any('vbaproject' in n.casefold() for n in package.namelist())
            assert {n: sha256(package.read(n)).hexdigest() for n in package.namelist()} == record['parts']


@pytest.mark.parametrize('case', CASES, ids=[case['name'] for case in CASES])
def test_native_shared_cse_full_runtime(case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(case, tmp_path)
    sheets = tuple(app.workbook.sheets_)
    arrays = [sheet.array_formulas.copy() for sheet in sheets]
    shared = [set(sheet.shared_groups) for sheet in sheets]
    cells = {id(c) for sheet in sheets for c in sheet.cells_.values()}
    assert app.run('Probe') == case['result']
    assert tuple(app.workbook.sheets_) == sheets
    assert [sheet.array_formulas for sheet in sheets] == arrays
    assert [set(sheet.shared_groups) for sheet in sheets] == shared
    assert cells <= {id(c) for sheet in sheets for c in sheet.cells_.values()}


@pytest.mark.parametrize('case', CASES, ids=[case['name'] for case in CASES])
def test_generated_shared_cse_save_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(case, tmp_path)
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / PureWindowsPath(case['save_package']['file']).name)
    assert app.run('Snapshot') == case['after_save_result']
    reopened = ExcelApplication.open(saved, with_vba=False)
    context(reopened, case)
    reopened.add_module(case['source'], name='RenameArrayControl')
    assert reopened.run('Snapshot') == case['reopened_result']


@pytest.mark.parametrize('case', CASES, ids=[case['name'] for case in CASES])
def test_native_input_shared_cse_public_rewrite(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / PureWindowsPath(case['save_package']['file']).name, with_vba=False)
    context(app, case)
    app.add_module(case['source'], name='RenameArrayControl')
    assert app.run('Snapshot') == case['reopened_result']
    sheet = app.workbook.sheets_[0]
    sheet.Range('A1').vba_set('Value2', sheet.Range('A1').vba_get('Value2'))
    saved = app.save(tmp_path / PureWindowsPath(case['save_package']['file']).name)
    reopened = ExcelApplication.open(saved, with_vba=False)
    context(reopened, case)
    reopened.add_module(case['source'], name='RenameArrayControl')
    assert reopened.run('Snapshot') == case['reopened_result']


def test_shared_xml_evidence_includes_real_master_and_cleared_hole() -> None:
    fresh, imported = CASES[:2]
    assert fresh['save_package']['observed_shared_formula_nodes'] == 3
    assert imported['before_clear_package']['observed_shared_formula_nodes'] == 3
    assert imported['seed_package']['observed_shared_formula_nodes'] == 2
    assert imported['save_package']['observed_shared_formula_nodes'] == 2
    seed = imported['seed_package']['formula_cells']['xl/worksheets/sheet1.xml'][0]
    final = imported['save_package']['formula_cells']['xl/worksheets/sheet1.xml'][0]
    assert (seed['cell'], seed['attributes']['ref']) == ('B2', 'B1:B3')
    assert (final['cell'], final['attributes']['ref']) == ('B2', 'B2:B3')
    # This is observed native file provenance, not a claim that rename caused
    # contraction rather than native's preceding load/save canonicalization.


@pytest.mark.parametrize('kind', ['array-overlap', 'array-orphan', 'array-member-formula',
                                  'shared-orphan', 'shared-outside', 'shared-inconsistent',
                                  'shared-qualified', 'late-name'])
def test_malformed_or_late_failed_storage_rename_is_atomic(kind: str) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    if kind == 'array-overlap':
        sheet.array_formulas[(2, 5)] = Area(2, 5, 3, 6)
        anchor = sheet.cell(2, 5, create=True)
        assert anchor is not None
        anchor.formula = '=1'
    elif kind == 'array-orphan':
        sheet.cells_.pop((1, 4))
    elif kind == 'array-member-formula':
        member = sheet.cell(1, 5, create=True)
        assert member is not None
        member.formula = '=5'
    elif kind == 'late-name':
        app.workbook.names_.entries.append(NameEntry('Last', '=1+"unterminated', app.workbook))
    else:
        sheet.Range('H1:H2').vba_set('Formula', '=A1*2')
        member = sheet.cells_[(2, 8)]
        if kind == 'shared-orphan':
            member.shared = 999
        elif kind == 'shared-outside':
            sheet.shared_groups[member.shared if member.shared is not None else -1] = Area(1, 8, 1, 8)
        elif kind == 'shared-qualified':
            for row in (1, 2):
                sheet.cells_[(row, 8)].formula = '=Sheet1!A' + str(row)
        else:
            member.formula = '=A2*3'
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError, match='overlapping' if kind == 'array-overlap' else None):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.mark.parametrize('bound', ['regions', 'cells'])
def test_storage_limits_admit_before_region_population(bound: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename_arrays

    app = prepared()
    before = fingerprint(app)
    monkeypatch.setattr(_worksheet_rename_arrays, 'MAX_REGIONS' if bound == 'regions' else 'MAX_REGION_CELLS', 0)
    with pytest.raises(VBAUnsupportedError, match='budget'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.mark.parametrize('formula', [
    '<f t="array" ref="A1:A2"/>',
    '<f t="dataTable" ref="A1:A2">A1</f>',
    '<f t="shared" si="5"/>',
    '<f t="array" ref="A1:A2" unknown="1">A1:A2</f>',
    '<f>A1<opaque/></f>',
])
def test_unowned_stored_formula_metadata_refuses_before_changes(formula: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    sheet = app.workbook.sheets_[0]
    app.workbook.package.write(sheet.part_name, (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<sheetData><row r="1"><c r="A1">' + formula + '</c></row></sheetData></worksheet>').encode())
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.fixture(scope='module')
def native_array_rename_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native generated shared/CSE rename gate')
@pytest.mark.parametrize('case', CASES, ids=[case['name'] for case in CASES])
def test_generated_shared_cse_reopened_by_native(
        case: dict[str, Any], tmp_path: Path, native_array_rename_excel: Any) -> None:
    app = configured(case, tmp_path)
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / PureWindowsPath(case['save_package']['file']).name)
    # Restore the recorded host setting before opening/injecting into the
    # generated workbook. It is not persisted workbook state or an oracle edit.
    native_array_rename_excel.new_document()
    setting = 'True' if case['configuration'].get('calculate_before_save') else 'False'
    host_source = ('Sub RestoreHostContext()\nApplication.EnableEvents=False\n'
                   'Application.CalculateBeforeSave=' + setting + '\nEnd Sub')
    host = native_array_rename_excel.run_vba(host_source, 'RestoreHostContext', module_name='NativeHostContext',
                                             line_numbers=False, timeout=30)
    assert host.ok and not host.dialogs, host
    native_array_rename_excel.open_document(saved)
    result = native_array_rename_excel.run_vba(case['source'], 'Snapshot', module_name='RenameArrayControl',
                                              line_numbers=False, timeout=30)
    assert result.ok and not result.dialogs, result
    assert result.value == case['reopened_result']


@pytest.mark.parametrize('body', [
    '<sheetData><row r="1"><c r="A1"><f><![CDATA["</f>"]]></f></c></row></sheetData>',
    '<sheetData><row r="1"><c r="A1"><!--<f>A1</f>--></c></row></sheetData>',
    '<sheetData><row:row xmlns:row="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="1"><c r="A1"><f>A1</f></c></row:row></sheetData>',
    '<sheetData><row r="1"><c:c xmlns:c="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A1"><f>A1</f></c:c></row></sheetData>',
    '<s:sheetData xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><row r="1"><c r="A1"><f>A1</f></c></row></s:sheetData>',
])
def test_lexical_formula_ownership_refuses_comment_cdata_and_qualified_storage(body: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    sheet = app.workbook.sheets_[0]
    app.workbook.package.write(sheet.part_name, (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + body + '</worksheet>').encode())
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


def test_sparse_huge_array_refuses_before_member_iteration() -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    sheet.array_formulas[(1, 4)] = Area(1, 4, 1_048_576, 16_384)
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError, match='budget'):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


def test_shared_text_budget_precedes_master_shift(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename, _worksheet_rename_arrays

    app = ExcelApplication()
    app.add_workbook()
    app.workbook.sheets_[0].Range('B1:B2').vba_set('Formula', '=A1*2')
    before = fingerprint(app)
    # Two source formula charges exhaust the budget before the first shared
    # master shift. No large adversarial formula or stress workload is needed.
    monkeypatch.setattr(_worksheet_rename, 'MAX_TEXT', 10)

    def forbidden(*_args: object) -> str:
        pytest.fail('shared master shift entered after exhausted text budget')

    monkeypatch.setattr(_worksheet_rename_arrays, 'shift_text', forbidden)
    with pytest.raises(VBAUnsupportedError, match='budget'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.mark.parametrize('kind', ['master', 'member', 'array-anchor'])
def test_storage_helper_whole_expression_preflight_before_any_lexer(
        kind: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _shared, _worksheet_rename, _worksheet_rename_arrays

    app = ExcelApplication()
    app.add_workbook()
    first = app.workbook.sheets_[0]
    first.Range('B1:B2').vba_set('Formula', '=A1*2')
    later = app.workbook.add_sheet('Later')
    if kind == 'array-anchor':
        later.Range('D1:E2').vba_set('FormulaArray', '=A1:A2*{1,2}')
        cell = later.cells_[(1, 4)]
    else:
        later.Range('B1:B2').vba_set('Formula', '=A1*2')
        cell = later.cells_[(1 if kind == 'master' else 2, 2)]
    cell.formula = '=' + '1+' * (_worksheet_rename.MAX_EXPRESSION_TEXT // 2) + '1'
    before = fingerprint(app)

    def forbidden(*_args: object) -> Any:
        pytest.fail('a lexer entered before whole-population expression admission')

    monkeypatch.setattr(_shared, 'names_a_sheet', forbidden)
    monkeypatch.setattr(_worksheet_rename_arrays, 'shift_text', forbidden)
    with pytest.raises(VBAUnsupportedError, match='expression text length'):
        _worksheet_rename_arrays.admit(app.workbook, first, 'Renamed', _worksheet_rename.Budget())
    assert fingerprint(app) == before


@pytest.mark.parametrize('fragment', [
    '<!--<row r="1"><c r="F1"><v>99</v></c></row>-->',
    '<![CDATA[<row r="1"><c r="F1"><v>99</v></c></row>]]>',
    '<!--<row r="9"/>-->',
])
def test_native_derived_row_shadow_refuses_before_rename(fragment: str, tmp_path: Path) -> None:
    original = ROOT / 'shared-owner-auto.xlsx'
    target = tmp_path / 'shadow.xlsx'
    with ZipFile(original) as source, ZipFile(target, 'w') as derived:
        for entry in source.infolist():
            raw = source.read(entry.filename)
            if entry.filename == 'xl/worksheets/sheet2.xml':
                raw = raw.replace(b'</sheetData>', fragment.encode() + b'</sheetData>')
            derived.writestr(entry, raw)
    app = ExcelApplication.open(target, with_vba=False)
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError, match='lexical'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.mark.parametrize('body', [
    '<sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f><v>1</v></c><!--<c r="A1"><v>99</v></c>--></row></sheetData>',
    '<sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f><v>1</v></c><![CDATA[<c r="A1"><v>99</v></c>]]></row></sheetData>',
    '<!--<sheetData><row r="1"><c r="A1"><v>99</v></c></row></sheetData>--><sheetData><row r="1"><c r="A1"><f>Sheet1!A1</f></c></row></sheetData>',
    '<sheetData><row r="1"><wrapped><c r="A1"><f>Sheet1!A1</f></c></wrapped></row></sheetData>',
])
def test_all_reader_layers_reject_unowned_cells_and_wrappers(body: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    sheet = app.workbook.sheets_[0]
    app.workbook.package.write(sheet.part_name, (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + body + '</worksheet>').encode())
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


@pytest.mark.parametrize('body', [
    "<sheetData><row r='1'><c r=\"A1\"><f>Sheet1!A1</f></c></row></sheetData>",
    "<sheetData><row r=\"1\"><c r='A1'><f>Sheet1!A1</f></c></row></sheetData>",
    "<sheetData><row r=\"1\"><c r=\"A1\"><f t='array' ref=\"A1:A2\">Sheet1!A1:A2</f></c></row></sheetData>",
    "<sheetData><row r=\"1\"><c r=\"A1\" opaque='keep r=\"B1\" tail'><f>Sheet1!A1</f></c></row></sheetData>",
    "<sheetData><row r=\"1\"><c opaque='>' r=\"A1\"><f>Sheet1!A1</f></c></row></sheetData>",
    "<sheetData><row r=\"1\"><c r=\"A1\"><f ca='1'>Sheet1!A1</f></c></row></sheetData>",
])
def test_owned_key_attributes_must_agree_with_legacy_reader(body: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    sheet = app.workbook.sheets_[0]
    app.workbook.package.write(sheet.part_name, (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + body + '</worksheet>').encode())
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError, match='key attributes'):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before


def test_storage_guard_leaves_other_formula_carriers_and_dirty_live_edits_alone(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._worksheet_rename_arrays import stored

    raw = (b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           b'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" '
           b'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
           b'<sheetData><row r="1"/><row r="2"><c r="A2"/></row></sheetData>'
           b'<extLst><ext uri="urn:test"><x14:dataValidations><x14:dataValidation>'
           b'<x14:formula1><xm:f>Other!A1</xm:f></x14:formula1></x14:dataValidation>'
           b'</x14:dataValidations></ext></extLst></worksheet>')
    stored(raw)
    app = ExcelApplication.open(ROOT / 'cse-other-source-auto.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    # Model changes may legitimately differ from the original package values.
    sheet.Range('A1').vba_set('Value2', 12)
    app.workbook.sheets_[1].Range('F1').vba_set('Formula', '=SUM(B2:C3)+1')
    sheet.vba_set('Name', 'Renamed')
    assert app.workbook.sheets_[1].Range('F1').vba_get('Value2') == 43
    assert ExcelApplication.open(app.save(tmp_path / 'edited.xlsx'), with_vba=False).sheet('Calc').value('F1') == 43


@pytest.mark.parametrize('local', ['sheetData', 'row', 'c', 'f'])
def test_unowned_openings_refuse_before_legacy_closing_tag_regexes(
        local: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _io, _worksheet_rename_arrays

    raw = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           '<sheetData><row r="1"><c r="A1"><f>1</f></c></row><!--' +
           ('<' + local + '>') * 20 + '--></sheetData></worksheet>').encode()

    def forbidden(*_args: object) -> Any:
        pytest.fail('legacy closing-tag regex entered before opening ownership admission')

    monkeypatch.setattr(_io, 'formula_spans', forbidden)
    with pytest.raises(VBAUnsupportedError, match='lexical'):
        _worksheet_rename_arrays.stored(raw)


@pytest.mark.parametrize('carrier', [
    '<sheetPr><f>Sheet1!A1</f></sheetPr>',
    '<sheetViews><sheetView workbookViewId="0"><f>Sheet1!A1</f></sheetView></sheetViews>',
    '<sheetPr><xm:f xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">Sheet1!A1</xm:f></sheetPr>',
])
def test_unknown_outside_cell_formula_carrier_refuses_atomically(carrier: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    sheet = app.workbook.sheets_[0]
    app.workbook.package.write(sheet.part_name, (
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + carrier +
        '<sheetData/></worksheet>').encode())
    before = fingerprint(app)
    with pytest.raises(VBAUnsupportedError, match='outside cell storage'):
        sheet.vba_set('Name', 'Renamed')
    assert fingerprint(app) == before
