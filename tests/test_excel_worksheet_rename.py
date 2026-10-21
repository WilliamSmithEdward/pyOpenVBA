"""Native scalar/name worksheet-rename observations and admission boundaries."""
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
from pyopenvba.apps.excel._model import NameEntry
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/worksheet_rename'
MANIFEST: dict[str, Any] = json.loads((ROOT / 'manifest.json').read_text())
RECORDS = {name: json.loads((ROOT / name).read_text()) for name in MANIFEST['raw_files']}
CASES = [(filename, case) for filename, record in RECORDS.items() for case in record['controls']]
FULL = [(filename, case) for filename, case in CASES
        if MANIFEST['cases'][filename + ':' + case['name']]['runtime'] == 'full']
INSPECTION = (ROOT / 'inspection.bas').read_text()


def configured(filename: str, case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    # Native's serial session allocated Book1…Book6. This is explicit initial
    # host context, not a replacement in an expected behavior string.
    context = MANIFEST['cases'][filename + ':' + case['name']]['initial_book_name']
    if context is not None:
        app.workbook.name = context
    app.add_module(case['source'], name='RenameControl')
    assert app.workbook.sheets_[0].code_name == 'Sheet1'
    return app


@pytest.mark.parametrize('filename', RECORDS)
def test_native_rename_record_and_source_bytes(filename: str) -> None:
    assert sha256((ROOT / filename).read_bytes()).hexdigest() == MANIFEST['raw_files'][filename]
    record = RECORDS[filename]
    assert record['completed'] and record['closed_utc']
    for case in record['controls']:
        assert case['status']['ok'] and not case['status']['dialogs']
        assert case['save_status']['ok'] and not case['save_status']['dialogs']
        assert case['reopened_status']['ok'] and not case['reopened_status']['dialogs']
        assert sha256(case['source'].encode()).hexdigest() == case['source_sha256']
    assert sha256(INSPECTION.encode()).hexdigest() == MANIFEST['inspection_source_sha256']


@pytest.mark.parametrize('filename,case', FULL, ids=[case['name'] for _, case in FULL])
def test_native_full_scalar_name_rename_runtime(filename: str, case: dict[str, Any]) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('filename,case', FULL, ids=[case['name'] for _, case in FULL])
def test_generated_scalar_rename_save_reopen(filename: str, case: dict[str, Any], tmp_path: Path) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']
    path = tmp_path / PureWindowsPath(case['file']).name
    saved = app.save(path)
    reopened = ExcelApplication.open(saved, with_vba=False)
    reopened.add_module(case.get('inspection_source', INSPECTION), name='InspectControl')
    assert reopened.run('Snapshot' if 'inspection_source' in case else 'Reopened') == case['reopened_result']


@pytest.mark.parametrize('filename,case', CASES, ids=[case['name'] for _, case in CASES])
def test_native_scalar_rename_input_code_identity_and_snapshot(filename: str, case: dict[str, Any]) -> None:
    sample = ROOT / PureWindowsPath(case['file']).name
    assert sha256(sample.read_bytes()).hexdigest() == case['sha256']
    with ZipFile(sample) as package:
        assert not any('vbaproject' in name.casefold() for name in package.namelist())
        for name, xml in case['xml'].items():
            assert package.read(name).decode('utf8') == xml
    app = ExcelApplication.open(sample, with_vba=False)
    owner = app.workbook.sheets_[0]
    identity = owner.code_name
    if filename in ('native-core.json', 'native-save-cache.json'):
        assert identity == 'Sheet1'
    app.add_module(case.get('inspection_source', INSPECTION), name='InspectControl')
    actual = app.run('Snapshot' if 'inspection_source' in case else 'Reopened')
    if case['name'] == 'strings-indirect-local-reference':
        # Native creates the hidden SINGLE name while opening/evaluating; it
        # is absent even from native's saved workbook XML. Only this explicitly
        # projected supported observation is compared, never a full-name claim.
        assert isinstance(actual, str)
        assert actual.partition(';name=[')[0] == case['reopened_result'].partition(';name=[')[0]
    else:
        assert actual == case['reopened_result']
    assert app.workbook.sheets_[0] is owner
    if identity:
        assert owner.code_name == identity


def test_native_string_text_projection_without_hidden_name_oracle_claim() -> None:
    filename, case = next((filename, case) for filename, case in CASES
                          if case['name'] == 'strings-indirect-local-reference')
    app = configured(filename, case)
    actual = app.run('Probe')
    assert isinstance(actual, str)
    # Explicit projection: retain only the measured literal/INDIRECT/implicit
    # formula observations. The complete raw result includes native's hidden
    # _xlfn.SINGLE creation, which is not implemented or normalized here.
    assert actual.partition(';final=[')[0] == case['result'].partition(';final=[')[0]
    assert '_xlfn.SINGLE' in case['result']


def prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.workbook.add_sheet('Calc')
    app.sheet(1).set_value('A1', 3)
    app.sheet(1).set_value('B1', '=A1*2')
    app.sheet('Calc').set_value('A1', '=Sheet1!A1*2')
    app.add_named_range('Global', '=Sheet1!$A$1')
    return app


def snapshot(app: ExcelApplication) -> object:
    book = app.workbook
    return (tuple((id(sheet), sheet.name, sheet.code_name, sheet.dirty,
                   tuple((key, id(cell), cell.formula, cell.value, cell.stale, cell.save_dirty)
                         for key, cell in sheet.cells_.items())) for sheet in book.sheets_),
            tuple((id(entry), entry.name, entry.refers_to, entry.attributes) for entry in book.names_.entries),
            book.names_.changed, book.saved,
            tuple(book.calculator.compiled), tuple(book.calculator._readers),
            tuple((entry.name, entry.body) for entry in book.package.entries) if book.package else None)


@pytest.mark.parametrize('kind', ['malformed-name', 'name-collision', '3d', 'external', 'cf', 'validation', 'array'])
def test_unsupported_rename_admission_is_atomic(kind: str) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    if kind == 'malformed-name':
        app.workbook.names_.entries.append(NameEntry('Last', '=1+"unterminated', app.workbook))
    elif kind == 'name-collision':
        app.workbook.names_.entries.append(NameEntry('Sheet1!Choice', '=1', app.workbook))
        app.workbook.names_.entries.append(NameEntry('Renamed!Choice', '=2', app.workbook))
    elif kind == '3d':
        app.workbook.names_.entries.append(NameEntry('Span', '=SUM(Sheet1:Calc!A1)', app.workbook))
    elif kind == 'external':
        app.workbook.names_.entries.append(NameEntry('Foreign', "='[Other.xlsx]Sheet1'!A1", app.workbook))
    elif kind == 'cf':
        sheet.Range('A1').FormatConditions().Add(2, Formula1='=Sheet1!A1>0')
    elif kind == 'validation':
        sheet.Range('A1').Validation().Add(1, Formula1='=Sheet1!A1', Formula2='10')
    else:
        sheet.Range('D1:D2').vba_set('FormulaArray', '=A1:A2')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_text_budget_admits_before_formula_parser(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename

    app = prepared()
    before = snapshot(app)
    monkeypatch.setattr(_worksheet_rename, 'MAX_TEXT', 1)

    def forbidden(*_args: object) -> str:
        pytest.fail('budget must reject before scalar formula parsing')

    monkeypatch.setattr(_worksheet_rename, '_formula', forbidden)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_enabled_document_handler_is_deferred_but_disabled_events_admit() -> None:
    app = prepared()
    app.add_module('Private Sub Worksheet_Calculate()\nEnd Sub', name='Sheet1', kind='document')
    sheet = app.workbook.sheets_[0]
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before
    app.workbook.application.enable_events = False
    sheet.vba_set('Name', 'Renamed')
    assert app.sheet('Calc').formula('A1') == '=Renamed!A1*2'


def test_held_scalar_and_name_identities_and_graph_follow_source_edit() -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    held = sheet.Range('A1')
    name = app.workbook.names_.entries[0]
    assert app.sheet('Calc').value('A1') == 6
    sheet.vba_set('Name', "O'Brien renamed")
    assert held.sheet is sheet and app.workbook.names_.entries[0] is name
    assert name.refers_to == "='O''Brien renamed'!$A$1"
    held.vba_set('Value2', 9)
    assert app.sheet('Calc').value('A1') == 18
    assert all(key[0] != 'sheet1' for key in app.workbook.calculator.compiled)



def test_enabled_bound_sink_is_deferred_without_running_handlers() -> None:
    app = prepared()
    app.add_module("Public WithEvents Watched As Worksheet\nPrivate Sub Watched_Calculate()\nErr.Raise 5\nEnd Sub", name='RenameSink', kind='class')
    app.add_module("Public Sink As RenameSink\nSub Attach()\nSet Sink=New RenameSink\nSet Sink.Watched=Worksheets(1)\nEnd Sub", name='AttachSink')
    app.run('Attach')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('xml', [
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst/></worksheet>',
    '<!DOCTYPE worksheet [<!ENTITY e "value">]><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
])
def test_persisted_unmodeled_or_declared_entity_metadata_refuses_before_changes(xml: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsm')
    part = app.workbook.sheets_[0].part_name
    app.workbook.package.write(part, xml.encode())
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_rename_metadata_budget_limits_actual_decompression() -> None:
    import zlib

    from pyopenvba.apps.excel._worksheet_rename import Budget, metadata_bytes
    from pyopenvba.powerquery._opc import Entry

    data = b'x' * 1000
    compressor = zlib.compressobj(wbits=-15)
    compressed = compressor.compress(data) + compressor.flush()
    # The declared size cannot authorize reading more than the transaction cap.
    entry = Entry('custom.rels', compressed, 8, 0, 0, 0, 0, 1)
    budget = Budget()
    budget.remaining = 64
    with pytest.raises(VBAUnsupportedError):
        metadata_bytes(entry, budget)



@pytest.mark.parametrize('metadata', [
    '<definedNames><definedName name="Bad!Scope">Sheet1!A1</definedName></definedNames>',
    '<definedNames><definedName name="Bad" localSheetId="999">Sheet1!A1</definedName></definedNames>',
    '<extLst><ext uri="urn:test"><definedNames><definedName name="Hidden">Sheet1!A1</definedName></definedNames></ext></extLst>',
])
def test_unowned_persisted_name_metadata_is_refused_atomically(metadata: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsm')
    package = app.workbook.package
    original = package.read('xl/workbook.xml').decode()
    import re

    original = re.sub(r'<definedNames>.*?</definedNames>', '', original)
    package.write('xl/workbook.xml', original.replace('</workbook>', metadata + '</workbook>').encode())
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before



def test_name_created_from_held_range_after_rename_uses_its_current_owner(tmp_path: Path) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    held = sheet.Range('A1')
    sheet.vba_set('Name', "O'Brien renamed")
    made = app.workbook.names_.Add('HeldAlias', RefersTo=held)
    assert made.entry.refers_to == "='O''Brien renamed'!$A$1"
    saved = app.save(tmp_path / 'held-name.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    assert reopened.named_range('HeldAlias').refers_to == "='O''Brien renamed'!$A$1"
    assert reopened.sheet("O'Brien renamed").value('HeldAlias') == 3



@pytest.fixture(scope='module')
def native_rename_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native generated worksheet-rename gate')
@pytest.mark.parametrize('filename,case', FULL, ids=[case['name'] for _, case in FULL])
def test_generated_worksheet_rename_reopened_by_native(
        filename: str, case: dict[str, Any], tmp_path: Path, native_rename_excel: Any) -> None:
    app = configured(filename, case)
    assert app.run('Probe') == case['result']
    saved = app.save(tmp_path / PureWindowsPath(case['file']).name)
    native_rename_excel.open_document(saved)
    result = native_rename_excel.run_vba(case.get('inspection_source', INSPECTION),
                                         'Snapshot' if 'inspection_source' in case else 'Reopened',
                                         module_name='NativeRead', line_numbers=False, timeout=30)
    assert result.ok and not result.dialogs, result
    assert result.value == case['reopened_result']


@pytest.mark.parametrize('fragment', [
    '<!--<definedName name="Phantom">Renamed!A1</definedName>-->',
    '<![CDATA[<definedName name="Phantom">Renamed!A1</definedName>]]>',
    '<!--<definedNames/>-->',
    '<definedNames><definedName name="Empty"/></definedNames>',
])
def test_native_derived_unowned_lexical_names_refuse_before_rename(fragment: str, tmp_path: Path) -> None:
    sample = ROOT / 'automatic-scalar-identity.xlsx'
    target = tmp_path / 'derived.xlsx'
    with ZipFile(sample) as original, ZipFile(target, 'w') as derived:
        for entry in original.infolist():
            raw = original.read(entry.filename)
            if entry.filename == 'xl/workbook.xml':
                raw = raw.replace(b'</workbook>', fragment.encode() + b'</workbook>')
            derived.writestr(entry, raw)
    app = ExcelApplication.open(target, with_vba=False)
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='name metadata'):
        app.workbook.sheets_[0].vba_set('Name', 'Next')
    assert snapshot(app) == before


def test_native_real_name_metadata_and_noop_remain_admitted(tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / 'ordinary-local-global-names.xlsx', with_vba=False)
    sheet = app.workbook.sheets_[0]
    identities = [id(entry) for entry in app.workbook.names_.entries]
    sheet.vba_set('Name', sheet.name)
    sheet.vba_set('Name', 'Next')
    assert [id(entry) for entry in app.workbook.names_.entries] == identities
    reopened = ExcelApplication.open(app.save(tmp_path / 'real.xlsx'), with_vba=False)
    assert len(reopened.workbook.names_.entries) == len(identities)


@pytest.mark.parametrize('kind', ['cell', 'name'])
def test_expression_bound_precedes_lexer_and_parser(kind: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename

    app = prepared()
    expression = '=1' + '+1' * (_worksheet_rename.MAX_EXPRESSION_TEXT // 2)
    if kind == 'cell':
        cell = app.workbook.sheets_[0].cell(9, 9, create=True)
        assert cell is not None
        cell.formula = expression
    else:
        app.workbook.names_.entries.append(NameEntry('Long', expression, app.workbook))
    before = snapshot(app)

    def forbidden(*_args: object) -> Any:
        pytest.fail('oversized expression reached lexer/parser')

    # The preflight scans every expression length before any lexer enters.
    monkeypatch.setattr(_worksheet_rename, 'tokenize', forbidden)
    monkeypatch.setattr(_worksheet_rename, 'parse', forbidden)
    with pytest.raises(VBAUnsupportedError, match='expression text length'):
        app.workbook.sheets_[0].vba_set('Name', 'Next')
    assert snapshot(app) == before


def test_removed_worksheet_noop_refuses() -> None:
    app = prepared()
    sheet = app.workbook.sheets_.pop()
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='removed worksheet'):
        sheet.vba_set('Name', sheet.name)
    assert snapshot(app) == before


@pytest.mark.parametrize('relationship', ['table', 'chart', 'pivotTable', 'unrecognized'])
def test_reference_relationships_cannot_hide_in_arbitrary_part_paths(relationship: str, tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsm')
    package = app.workbook.package
    raw = package.read('xl/_rels/workbook.xml.rels')
    relation = ('<Relationship Id="Hidden" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/' + relationship + '" Target="Unusual/opaque.xml"/>')
    package.write('xl/_rels/workbook.xml.rels', raw.replace(b'</Relationships>', relation.encode() + b'</Relationships>'))
    package.write('xl/Unusual/opaque.xml', b'<opaque/>')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='relationships'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_relocated_worksheet_metadata_uses_its_declared_part_identity(tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsm')
    sheet = app.workbook.sheets_[0]
    sheet.part_name = 'xl/Unusual/opaque.xml'
    app.workbook.package.write(sheet.part_name, (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<pivotTableParts/></worksheet>'))
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='reference metadata'):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('header', ['Private Sub _\nWorksheet_Calculate()',
                                    'Private _\nSub Worksheet_Calculate()',
                                    'Private Sub Worksheet_Calculate()'])
def test_pending_document_continuation_observer_refuses_atomically(header: str) -> None:
    app = prepared()
    sheet = app.workbook.sheets_[0]
    sheet.code_name = 'Sheet1'
    app.workbook.pending_document_sources['Sheet1'] = header + '\nEnd Sub'
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='pending document event observers'):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before
    assert sheet.vba_document is None
    app.workbook.application.enable_events = False
    sheet.vba_set('Name', 'Renamed')
    assert sheet.name == 'Renamed'
    assert app.sheet('Calc').formula('A1') == '=Renamed!A1*2'


def test_unrelated_pending_document_headers_and_string_literals_admit() -> None:
    app = prepared()
    app.workbook.pending_document_sources['Sheet1'] = (
        'Private Sub _\nOrdinaryHelper()\n'
        'Debug.Print "Private Sub Worksheet_Calculate()"\n'
        "'Private Sub Worksheet_Calculate()\nEnd Sub")
    app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert app.sheet('Calc').formula('A1') == '=Renamed!A1*2'


def test_pending_document_text_bound_precedes_vba_lexer(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename

    app = prepared()
    app.workbook.pending_document_sources['Sheet1'] = "'" + 'x' * _worksheet_rename.MAX_EXPRESSION_TEXT
    before = snapshot(app)

    def forbidden(*_args: object) -> Any:
        pytest.fail('oversized pending source reached VBA lexer')

    monkeypatch.setattr(_worksheet_rename, 'vba_tokens', forbidden)
    with pytest.raises(VBAUnsupportedError, match='pending-document text length'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_pending_document_collection_count_precedes_copy() -> None:
    class OversizedSources(dict[str, str]):
        def values(self) -> Any:
            pytest.fail('oversized pending collection was copied')

    app = prepared()
    app.workbook.pending_document_sources = OversizedSources({str(index): '' for index in range(1001)})
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='pending-document count'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('loaded', [False, True])
def test_unaffected_conditional_rename_preserves_roots_and_registry(loaded: bool, tmp_path: Path) -> None:
    from pyopenvba.apps.excel._charts_file import xml_bytes

    app = ExcelApplication()
    app.add_workbook()
    owner = app.workbook.sheets_[0]
    owner.Range('A1').vba_set('Value', 3)
    owner.Range('A1:A3').FormatConditions().Add(2, Formula1='=A1>0')
    if loaded:
        app = ExcelApplication.open(app.save(tmp_path / 'conditional.xlsx'), with_vba=False)
        owner = app.workbook.sheets_[0]
    rule = owner.conditional_formats[0]
    before = (id(rule), id(rule.root), xml_bytes(rule.root), dict(owner.conditional_color_slots),
              owner.conditional_color_next, owner.conditional_color_groups)
    owner.vba_set('Name', "O'Brien")
    assert (id(rule), id(rule.root), xml_bytes(rule.root), dict(owner.conditional_color_slots),
            owner.conditional_color_next, owner.conditional_color_groups) == before
    reopened = ExcelApplication.open(app.save(tmp_path / 'renamed.xlsx'), with_vba=False)
    assert reopened.workbook.sheets_[0].name == "O'Brien"
    formula = reopened.workbook.sheets_[0].conditional_formats[0].root.find(
        '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}formula')
    assert formula is not None and formula.text == 'A1>0'


@pytest.mark.parametrize('kind', ['conditional', 'validation', 'control-link', 'control-list'])
def test_secondary_affected_reference_refusal_preserves_entire_scalar_plan(kind: str) -> None:
    app = prepared()
    owner = app.workbook.sheets_[1]
    if kind == 'conditional':
        owner.Range('A1').FormatConditions().Add(2, Formula1="='Sheet1'!A1>0")
    elif kind == 'validation':
        owner.Range('A1').Validation().Add(1, Formula1="='Sheet1'!A1", Formula2='10')
    else:
        app.sheet('Calc').add_form_control(2, name='Linked')
        control = owner.shapes_[0].control
        assert control is not None
        if kind == 'control-link':
            control.linked_cell = "'Sheet1'!A1"
        else:
            control.list_range = "'Sheet1'!A1:A3"
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='affected secondary'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


def test_unrelated_local_secondary_features_do_not_block_fresh_sheet_rename() -> None:
    app = prepared()
    owner = app.workbook.sheets_[0]
    owner.Range('A1').FormatConditions().Add(2, Formula1='=A1>0')
    owner.Range('B1').Validation().Add(1, Formula1='0', Formula2='10')
    app.sheet(1).add_form_control(2, name='Linked')
    control = owner.shapes_[0].control
    assert control is not None
    control.linked_cell = 'A1'
    control.list_range = 'A1:A3'
    fresh = app.workbook.add_sheet('Fresh')
    fresh.vba_set('Name', 'Witness')
    assert fresh.name == 'Witness'
    assert control.linked_cell == 'A1' and control.list_range == 'A1:A3'


@pytest.mark.parametrize('payload', [
    '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>Sheet1!A1&gt;0</formula></cfRule></conditionalFormatting>',
    '<extLst><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting></extLst>',
    '<conditionalFormatting sqref="A1" unknown="Sheet1!A1"><cfRule type="expression" priority="1"/></conditionalFormatting>',
])
def test_unmaterialized_or_opaque_conditional_package_references_are_atomic(payload: str) -> None:
    app = ExcelApplication.open(ROOT / 'automatic-scalar-identity.xlsx', with_vba=False)
    sheet, package = app.workbook.sheets_[0], app.workbook.package
    assert package is not None
    raw = package.read(sheet.part_name)
    package.write(sheet.part_name, raw.replace(b'</worksheet>', payload.replace('Sheet1', sheet.name).encode() + b'</worksheet>'))
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('kind', ['shared', 'array'])
def test_unrelated_structured_formulas_remain_unchanged_when_blank_sheet_is_renamed(kind: str) -> None:
    app = prepared()
    owner = app.workbook.sheets_[0]
    target = owner.Range('D1:D2')
    target.vba_set('Formula' if kind == 'shared' else 'FormulaArray', '=A1*2' if kind == 'shared' else '=A1:A2')
    assert owner.shared_groups if kind == 'shared' else owner.array_formulas
    before = (dict(owner.shared_groups), dict(owner.array_formulas),
              tuple((key, id(cell), cell.formula) for key, cell in owner.cells_.items()))
    fresh = app.workbook.add_sheet('Fresh')
    fresh.vba_set('Name', 'Other')
    assert (dict(owner.shared_groups), dict(owner.array_formulas),
            tuple((key, id(cell), cell.formula) for key, cell in owner.cells_.items())) == before


def test_affected_foreign_cse_formula_is_refused_before_any_scalar_change() -> None:
    app = prepared()
    owner = app.workbook.sheets_[1]
    owner.Range('D1:D2').vba_set('FormulaArray', '=Sheet1!A1:A2')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='affected shared/array/spill'):
        app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('kind', ['shared', 'array'])
def test_unrelated_formula_storage_manual_caches_and_members_survive_rename(kind: str, tmp_path: Path) -> None:
    app = prepared()
    owner = app.workbook.sheets_[0]
    owner.Range('D1:D2').vba_set('Formula' if kind == 'shared' else 'FormulaArray', '=A1*2' if kind == 'shared' else '=A1:A2')
    app.workbook.application.CalculateFull()
    app.workbook.application.calculation = -4135
    app.workbook.application.calculate_before_save = False
    owner.Range('A1').vba_set('Value', 9)
    groups = (dict(owner.shared_groups), dict(owner.array_formulas))
    before = {key: (id(cell), cell.value, cell.formula, cell.shared, cell.save_dirty if cell.save_dirty is not None else cell.stale)
              for key, cell in owner.cells_.items()}
    fresh = app.workbook.add_sheet('Fresh')
    fresh.vba_set('Name', 'Other')
    assert groups == (dict(owner.shared_groups), dict(owner.array_formulas))
    assert before == {key: (id(cell), cell.value, cell.formula, cell.shared, cell.save_dirty if cell.save_dirty is not None else cell.stale)
                      for key, cell in owner.cells_.items()}
    saved = app.save(tmp_path / ('unchanged-' + kind + '.xlsx'))
    reopened = ExcelApplication.open(saved, with_vba=False)
    restored = reopened.workbook.sheets_[0]
    assert restored.name == owner.name
    assert tuple({key: area.address() for key, area in regions.items()}
                 for regions in (restored.shared_groups, restored.array_formulas)) == tuple(
                     {key: area.address() for key, area in regions.items()} for regions in groups)
    assert tuple(restored.cells_[key].value for key in ((1, 4), (2, 4))) == tuple(
        owner.cells_[key].value for key in ((1, 4), (2, 4)))


@pytest.mark.parametrize('damage', ['orphan', 'overlap', 'budget'])
def test_unrelated_malformed_formula_storage_is_refused_atomically(damage: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _worksheet_rename_arrays

    app = prepared()
    owner = app.workbook.sheets_[0]
    owner.Range('D1:D2').vba_set('Formula', '=A1*2')
    fresh = app.workbook.add_sheet('Fresh')
    if damage == 'orphan':
        owner.cells_[(2, 4)].shared = 999
    elif damage == 'overlap':
        owner.array_formulas[(1, 4)] = next(iter(owner.shared_groups.values()))
    else:
        monkeypatch.setattr(_worksheet_rename_arrays, 'MAX_REGION_CELLS', 1)
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        fresh.vba_set('Name', 'Other')
    assert snapshot(app) == before


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
    before = snapshot(app)

    def forbidden(*_args: object) -> Any:
        pytest.fail('a lexer entered before whole-population expression admission')

    monkeypatch.setattr(_shared, 'names_a_sheet', forbidden)
    monkeypatch.setattr(_worksheet_rename_arrays, 'shift_text', forbidden)
    with pytest.raises(VBAUnsupportedError, match='expression text length'):
        _worksheet_rename_arrays.admit(app.workbook, first, 'Renamed', _worksheet_rename.Budget())
    assert snapshot(app) == before



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
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before


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
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='key attributes'):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before



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
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='outside cell storage'):
        sheet.vba_set('Name', 'Renamed')
    assert snapshot(app) == before


@pytest.mark.parametrize('fragment', [
    '<!--<row r="1"><c r="A1"><v>99</v></c></row>-->',
    '<![CDATA[<row r="1"><c r="A1"><v>99</v></c></row>]]>',
])
def test_real_imported_cell_shadow_refuses_before_public_rename(fragment: str, tmp_path: Path) -> None:
    app = prepared()
    original = app.save(tmp_path / 'original.xlsx')
    derived = tmp_path / 'shadow.xlsx'
    with ZipFile(original) as source, ZipFile(derived, 'w') as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == 'xl/worksheets/sheet2.xml':
                data = data.replace(b'</sheetData>', fragment.encode() + b'</sheetData>')
            target.writestr(entry, data)
    imported = ExcelApplication.open(derived, with_vba=False)
    assert imported.workbook.sheets_[1].cells_[(1, 1)].formula == ''
    assert imported.sheet('Calc').value('A1') == 99
    before = snapshot(imported)
    with pytest.raises(VBAUnsupportedError, match='lexical'):
        imported.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(imported) == before


def test_storage_guard_accepts_current_dirty_values_and_formulas(tmp_path: Path) -> None:
    app = prepared()
    saved = app.save(tmp_path / 'before.xlsx')
    app = ExcelApplication.open(saved, with_vba=False)
    app.sheet(1).set_value('A1', 12)
    app.sheet('Calc').set_value('A1', '=Sheet1!A1*3')
    app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert app.sheet('Calc').formula('A1') == '=Renamed!A1*3'
    assert app.sheet('Calc').value('A1') == 36
    reopened = ExcelApplication.open(app.save(tmp_path / 'after.xlsx'), with_vba=False)
    assert reopened.sheet('Calc').value('A1') == 36


@pytest.mark.parametrize('damage', ['none', 'affected', 'cache-attribute', 'cache-child', 'field-attribute', 'report-child'])
def test_known_pivot_source_admission_is_read_only_and_atomic(damage: str) -> None:
    from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes
    from pyopenvba.apps.excel._pivots import PivotCache, PivotTable

    app = prepared()
    owner = app.workbook.sheets_[0]
    for address, value in (('G1', 'Header'), ('G2', 1), ('G3', 2)):
        owner.Range(address).vba_set('Value', value)
    cache = app.workbook.PivotCaches().Create(1, owner.Range('G1:G3'))
    assert isinstance(cache, PivotCache)
    report = cache.CreatePivotTable(owner.Range('J1'), 'Report')
    assert isinstance(report, PivotTable)
    if damage == 'cache-attribute':
        cache.root.set('unknown', 'Fresh!A1')
    elif damage == 'cache-child':
        cache.root.append(XmlElement('{urn:unknown}reference', {'sheet': 'Fresh'}))
    elif damage == 'field-attribute':
        fields = cache.root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheFields')
        assert fields is not None
        fields[0].set('formula', 'Fresh!A1')
    elif damage == 'report-child':
        report.root.append(XmlElement('{urn:unknown}reference', {'sheet': 'Fresh'}))
    fresh = app.workbook.add_sheet('Fresh')
    target = owner if damage == 'affected' else fresh
    before = snapshot(app), xml_bytes(cache.root), xml_bytes(report.root)
    if damage == 'none':
        target.vba_set('Name', 'Other')
        assert xml_bytes(cache.root) == before[1] and xml_bytes(report.root) == before[2]
    else:
        with pytest.raises(VBAUnsupportedError):
            target.vba_set('Name', 'Other')
        assert (snapshot(app), xml_bytes(cache.root), xml_bytes(report.root)) == before


@pytest.mark.parametrize('payload', [
    '<cfRule type="expression" priority="1"><formula>Sheet1!A1&gt;0</formula></cfRule>',
    '<sheetPr><formula>Sheet1!A1</formula></sheetPr>',
    '<dataValidation type="custom"><formula1>Sheet1!A1</formula1></dataValidation>',
])
def test_imported_secondary_formula_parent_ownership_is_atomic(payload: str, tmp_path: Path) -> None:
    app = prepared()
    original = app.save(tmp_path / 'original.xlsx')
    derived = tmp_path / 'secondary.xlsx'
    with ZipFile(original) as source, ZipFile(derived, 'w') as target:
        for entry in source.infolist():
            data = source.read(entry.filename)
            if entry.filename == 'xl/worksheets/sheet2.xml':
                data = data.replace(b'</worksheet>', payload.encode() + b'</worksheet>')
            target.writestr(entry, data)
    imported = ExcelApplication.open(derived, with_vba=False)
    before = snapshot(imported)
    with pytest.raises(VBAUnsupportedError, match='unowned secondary'):
        imported.workbook.sheets_[0].vba_set('Name', 'Renamed')
    assert snapshot(imported) == before


def test_unchanged_opaque_metadata_is_preserved_without_generic_tag_refusal(tmp_path: Path) -> None:
    app = prepared()
    app.save(tmp_path / 'before.xlsx')
    package = app.workbook.package
    assert package is not None
    sheet = app.workbook.sheets_[1]
    opaque = b'<customAnnotation token="keep">opaque text</customAnnotation>'
    raw = package.read(sheet.part_name)
    package.write(sheet.part_name, raw.replace(b'</worksheet>', opaque + b'</worksheet>'))
    app.workbook.sheets_[0].vba_set('Name', 'Renamed')
    saved = app.save(tmp_path / 'opaque.xlsx')
    with ZipFile(saved) as result:
        assert opaque in result.read(sheet.part_name)
