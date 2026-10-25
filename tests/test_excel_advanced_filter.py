"""Independent native AdvancedFilter API and workbook persistence controls."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from hashlib import sha256
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).parent / 'fixtures/advanced_filter'
RECORD: dict[str, Any] = json.loads((ROOT / 'native.json').read_text(encoding='utf8'))
CASES = [case for case in RECORD['probes'] if case['admitted']]
SAMPLES = [case for case in CASES if 'fixture_file' in case]
CALCULATION: dict[str, Any] = json.loads((ROOT / 'calculation.json').read_text(encoding='utf8'))
PROJECTION: dict[str, Any] = json.loads((ROOT / 'calculation_projection.json').read_text(encoding='utf8'))


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_advanced_filter_runtime(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='AdvancedControl')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_native_advanced_filter_input_and_forced_rewrite(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['fixture_file'], with_vba=False)
    app.add_module(case['inspection_source'], name='AdvancedRead')
    assert app.run('Snapshot') == case['reopened']
    app.workbook.sheets_[0].touched()
    result = app.save(tmp_path / 'rewritten.xlsm')
    assert ExcelApplication.open(result).run('Snapshot') == case['reopened']


@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_generated_advanced_filter_save_reopen(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='AdvancedControl')
    assert app.run('Probe') == case['result']
    result = app.save(tmp_path / 'generated.xlsm')
    assert ExcelApplication.open(result).run('Snapshot') == case['reopened']


def prepared() -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Sub Prepare()
Range("A1:B1").Value=Array("Name","Qty")
Range("A2:B2").Value=Array("apple",1)
Range("A3:B3").Value=Array("pear",2)
Range("D1:E1").Value=Array("Name","Qty")
Range("D2:E5").Value="old"
Range("G1").Value="Name"
Range("G2").Value="a*"
End Sub''', name='PrepareData')
    app.run('Prepare')
    return app


def state(app: ExcelApplication) -> object:
    from pyopenvba.apps.excel._charts_file import xml_bytes

    sheet = app.workbook.sheets_[0]
    return ([(key, cell.value, cell.formula, cell.style, cell.stale) for key, cell in sheet.cells_.items()],
            [(entry.name, entry.refers_to, entry.visible, entry.attributes) for entry in app.workbook.names_.entries],
            [(id(rule), tuple(rule.areas), xml_bytes(rule.root)) for rule in sheet.conditional_formats],
            sheet.advanced_filter, sheet.advanced_filter_changed, sheet.auto_filter, sheet.filter_changed, sheet.dirty,
            app.workbook.saved, app.workbook.stylesheet.dxf_count,
            tuple((row, record.hidden) for row, record in sheet.dims.rows.items()))


@pytest.mark.parametrize('setup,source,criteria,destination', [
    ('Range("G1").Value="Unmatched"', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('Range("G2").Value=String(255,"a")', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('Range("D1").Value="Unmatched"', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('ActiveSheet.Protect', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('Range("A2:B2").Merge', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('Range("A2").FormatConditions.Add 2,Formula1:="=TRUE"', 'A1:B3', 'G1:G2', 'D1:E1'),
    ('', 'A1:B3', 'G1:G2', 'A1:B1'),
    ('', 'A1:B3', 'G:G', 'D1:E1'),
])
def test_unsupported_advanced_filter_is_atomic(setup: str, source: str, criteria: str, destination: str) -> None:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    if setup:
        app.add_module('Sub Configure()\n' + setup + '\nEnd Sub', name='ConfigureData')
        app.run('Configure')
    sheet = app.workbook.sheets_[0]
    before = state(app)
    asked = sheet.Range(source)
    assert isinstance(asked, Range)
    with pytest.raises(VBAUnsupportedError):
        asked.AdvancedFilter(2, sheet.Range(criteria), sheet.Range(destination))
    assert state(app) == before


def test_advanced_filter_leaves_unrelated_user_criteria_names_ordinary(tmp_path: Path) -> None:
    from zipfile import ZipFile

    app = prepared()
    app.add_module('Sub AddNames()\nActiveWorkbook.Names.Add Name:="Criteria",RefersTo:="=1"\nActiveWorkbook.Names.Add Name:="Extract",RefersTo:="=2"\nEnd Sub', name='UserNames')
    app.run('AddNames')
    saved = app.save(tmp_path / 'user-names.xlsm')
    with ZipFile(saved) as package:
        xml = package.read('xl/workbook.xml').decode()
        assert 'name="Criteria"' in xml and 'name="Extract"' in xml
        assert 'name="_xlnm.Criteria"' not in xml and 'name="_xlnm.Extract"' not in xml


def test_advanced_filter_wildcards_do_not_use_backtracking_regex() -> None:
    app = prepared()
    app.add_module('''Function LongPattern() As String
Range("A2").Value=String(2000,"a")
Range("G2").Value=String(1,"*")
Dim i As Long
For i=1 To 60
Range("G2").Value=Range("G2").Value & "a*"
Next
Range("G2").Value=Range("G2").Value & "b"
Range("A1:B3").AdvancedFilter 2,Range("G1:G2"),Range("D1:E1")
LongPattern=CStr(IsEmpty(Range("D2").Value))
End Function''', name='BoundedWildcard')
    assert app.run('LongPattern') == 'True'


@pytest.mark.parametrize('prefix', ['', 's:'])
@pytest.mark.parametrize('existing', [False, True])
def test_advanced_filter_flag_preserves_qualified_unknown_xml(prefix: str, existing: bool) -> None:
    from pyopenvba.apps.excel._advanced_filter import with_filter_mode
    from pyopenvba.apps.excel._charts_file import xml_tree

    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    declaration = 'xmlns:s' if prefix else 'xmlns'
    properties = ('<' + prefix + 'sheetPr filterMode = \'0\' xmlns:x="urn:extra" x:opaque="a&gt;b"/>') if existing else ''
    raw = ('<' + prefix + 'worksheet ' + declaration + '="' + namespace + '">' + properties +
           '<' + prefix + 'sheetData/><!--preserve--></' + prefix + 'worksheet>')
    saved = with_filter_mode(raw, True)
    root = xml_tree(saved)
    element = root.find('{' + namespace + '}sheetPr')
    assert element is not None and element.get('filterMode') == '1'
    assert '<!--preserve-->' in saved
    if existing:
        assert element.get('{urn:extra}opaque') == 'a>b'
    cleared = with_filter_mode(saved, False)
    cleared_properties = xml_tree(cleared).find('{' + namespace + '}sheetPr')
    assert cleared_properties is not None and cleared_properties.get('filterMode') is None


@pytest.mark.parametrize('operation', [
    'Rows(2).Insert', 'Rows(2).Delete', 'Range("A2").Delete Shift:=xlShiftUp',
    'Range("A1:B3").Cut Destination:=Range("J1")', 'ActiveSheet.Copy', 'Range("A1:B3").AutoFilter',
])
def test_unimplemented_advanced_filter_state_migration_is_atomic(operation: str) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    app.add_module('Sub Operate()\n' + operation + '\nEnd Sub\nSub FilterFirst()\nRange("A1:B3").AdvancedFilter 1,Range("G1:G2")\nEnd Sub', name='StateMigration')
    app.run('FilterFirst')
    before, books, clipboard = state(app), tuple(app.workbook.application.workbooks_.books), app.workbook.application.clipboard
    with pytest.raises(VBAUnsupportedError):
        app.run('Operate')
    assert state(app) == before and tuple(app.workbook.application.workbooks_.books) == books and app.workbook.application.clipboard is clipboard


@pytest.mark.parametrize('unique', [False, True])
def test_advanced_filter_rejects_large_duplicate_projection_before_allocating_output(unique: bool, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBAUnsupportedError
    from pyopenvba.interpreter._values import MISSING

    app = prepared()
    app.add_module('Sub WideHeaders()\nRange("D1:IV1").Value="Name"\nEnd Sub', name='ProjectionSetup')
    app.run('WideHeaders')
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:A5000')
    assert isinstance(source, Range)
    before = state(app)
    from pyopenvba.apps.excel import _advanced_filter

    def forbidden_read(*_args: object) -> object:
        pytest.fail('source or projection was read before geometry admission')

    monkeypatch.setattr(_advanced_filter, '_value', forbidden_read)
    with pytest.raises(VBAUnsupportedError, match='million cells'):
        source.AdvancedFilter(2, MISSING, sheet.Range('D1:IV1'), unique)
    assert state(app) == before


@pytest.fixture(scope='module')
def native_advanced_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native Excel gate')
@pytest.mark.parametrize('case', SAMPLES, ids=lambda case: case['name'])
def test_generated_advanced_filter_reopened_by_native_excel(
        case: dict[str, Any], tmp_path: Path, native_advanced_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='AdvancedControl')
    assert app.run('Probe') == case['result']
    result = app.save(tmp_path / 'native-generated.xlsm')
    native_advanced_excel.open_document(result)
    assert native_advanced_excel.run_raw('AdvancedControl.Snapshot') == case['reopened']


@pytest.mark.parametrize('reference', [None, '=#REF!', '=Other!$A$1:$B$3', '=Sheet1!$A$1:$A$3,$C$1:$C$3'])
def test_unsupported_imported_filter_list_is_not_silently_activated(reference: str | None) -> None:
    from pyopenvba.apps.excel._advanced_filter import read_state
    from pyopenvba.apps.excel._model import NameEntry
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    sheet = app.workbook.sheets_[0]
    if reference is not None:
        app.workbook.names_.entries.append(NameEntry('Sheet1!_FilterDatabase', reference, app.workbook))
    before = state(app)
    with pytest.raises(VBAUnsupportedError):
        read_state(sheet, b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr filterMode="1"/></worksheet>')
    assert state(app) == before


@pytest.mark.parametrize('unique', [False, True])
def test_advanced_filter_large_criteria_rejected_before_any_cell_read(unique: bool, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _advanced_filter
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    sheet = app.workbook.sheets_[0]
    before = state(app)

    def forbidden_read(*_args: object) -> object:
        pytest.fail('a cell was read before criteria geometry admission')

    monkeypatch.setattr(_advanced_filter, '_value', forbidden_read)
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    with pytest.raises(VBAUnsupportedError, match='4096 criteria cells'):
        source.AdvancedFilter(2, sheet.Range('G1:H1000000'), sheet.Range('D1:E1'), unique)
    assert state(app) == before


@pytest.mark.parametrize('self_closing', [False, True])
def test_filter_flag_does_not_patch_inside_other_attribute_values(self_closing: bool) -> None:
    from pyopenvba.apps.excel._advanced_filter import with_filter_mode
    from pyopenvba.apps.excel._charts_file import xml_tree

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    attrs = "xmlns:x=\"urn:filterMode='0'\" x:opaque=\"keep filterMode='0' tail\" x:filterMode='0'"
    raw = '<worksheet xmlns="' + ns + '"><sheetPr ' + attrs + ('/>' if self_closing else '></sheetPr>') + '<sheetData/></worksheet>'
    for filtered in (True, False):
        saved = with_filter_mode(raw, filtered)
        assert attrs in saved
        props = xml_tree(saved).find('{' + ns + '}sheetPr')
        assert props is not None
        assert props.get('filterMode') == ('1' if filtered else None)
        assert props.get("{urn:filterMode='0'}opaque") == "keep filterMode='0' tail"
        assert props.get("{urn:filterMode='0'}filterMode") == '0'


def test_filter_owned_names_preserve_opaque_quoted_attributes_through_save(tmp_path: Path) -> None:
    from zipfile import ZipFile

    from pyopenvba.apps.excel._charts_file import xml_tree
    from pyopenvba.apps.excel._model import NameEntry, Range

    app = prepared()
    sheet = app.workbook.sheets_[0]
    opaque = "xmlns:x='urn:extra' x:opaque='keep name=\"other\" > tail'"
    app.workbook.names_.entries.append(NameEntry('Sheet1!Criteria', '=1', app.workbook,
        attributes=opaque + ' name="_xlnm.Criteria" localSheetId="0"'))
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    assert source.AdvancedFilter(2, sheet.Range('G1:G2'), sheet.Range('D1:E1')) is True
    for filename in ('names.xlsm', 'names-reopened.xlsm'):
        saved = app.save(tmp_path / filename)
        with ZipFile(saved) as package:
            xml = package.read('xl/workbook.xml').decode()
        assert opaque in xml
        root = xml_tree(xml)
        entries = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName')
        entry = next(element for element in entries if element.get('name') == '_xlnm.Criteria')
        assert entry.get('{urn:extra}opaque') == 'keep name="other" > tail'
        assert entry.text == 'Sheet1!$G$1:$G$2'
        app = ExcelApplication.open(saved)
        assert any(item.name == 'Sheet1!Criteria' for item in app.workbook.names_.entries)


def test_filter_malformed_later_name_is_rejected_before_any_name_or_output_mutation() -> None:
    from pyopenvba.apps.excel._model import NameEntry, Range
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    sheet = app.workbook.sheets_[0]
    app.workbook.names_.entries.append(NameEntry('Sheet1!Criteria', '=1', app.workbook,
                                               attributes='x:opaque="unterminated'))
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    before = state(app)
    with pytest.raises(VBAUnsupportedError, match='unterminated attribute'):
        source.AdvancedFilter(2, sheet.Range('G1:G2'), sheet.Range('D1:E1'))
    assert state(app) == before


def test_filtered_visibility_calculations_survive_generated_roundtrip(tmp_path: Path) -> None:
    case = CALCULATION['probes'][0]
    source, result = projected_calculation_control(case)
    # Stop at the independently measured filtered phase, preserving its formulas.
    source = source[:source.index('Application.Calculate\ns=s & ";calculated=')]
    source += 'Probe=Snap()\nEnd Function\n'
    expected = result.split(';filtered=[', 1)[1].split('];calculated=', 1)[0]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='FilterCalculation')
    assert app.run('Probe') == expected
    saved = app.save(tmp_path / 'filtered-calculations.xlsm')
    assert ExcelApplication.open(saved).run('Snap') == expected


def projected_calculation_control(case: dict[str, Any]) -> tuple[str, str]:
    """A labeled native projection: unsupported CalculationState alone is omitted."""
    assert sha256((ROOT / 'calculation.json').read_bytes()).hexdigest() == PROJECTION['raw_fixture_sha256']
    assert sha256(case['source'].encode()).hexdigest() == PROJECTION['source_sha256'][case['name']]
    expression = ' & ";state=" & CStr(Application.CalculationState)'
    assert case['source'].count(expression) == 1
    source = case['source'].replace(expression, '')
    result = case['result'].replace(';state=0;', ';').replace(';state=2;', ';')
    assert 'CalculationState' not in source and ';state=' not in result
    return source, result


@pytest.mark.parametrize('case', CALCULATION['probes'], ids=lambda case: case['name'])
def test_native_calculation_values_and_visibility_projection(case: dict[str, Any]) -> None:
    source, expected = projected_calculation_control(case)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='ProjectedCalculation')
    assert app.run('Probe') == expected


@pytest.mark.parametrize('criterion', ['*b', '2'])
def test_aggregate_text_work_is_rejected_before_an_expensive_match_or_numeric_parse(
        criterion: str, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _advanced_filter, _typing
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    app.add_module('Sub LongValue()\nRange("A2").Value=String(100,"a")\nEnd Sub', name='BudgetSetup')
    app.run('LongValue')
    criteria_cell = sheet.cells_[(2, 7)]
    criteria_cell.value = criterion
    before = state(app)
    original_work = getattr(_advanced_filter, '_Work')
    monkeypatch.setattr(_advanced_filter, '_Work', lambda: original_work(50))
    original_typing = _typing.typed_text

    def check_typing(value: str, *args: Any, **kwargs: Any) -> Any:
        assert len(value) < 50, 'long numeric text scanned before aggregate admission'
        return original_typing(value, *args, **kwargs)

    def check_match(text: str, _pattern: str, *, is_text: bool) -> bool:
        pytest.fail('long text scanned before aggregate admission')

    monkeypatch.setattr(_typing, 'typed_text', check_typing)
    monkeypatch.setattr(_advanced_filter, '_matches', check_match)
    with pytest.raises(VBAUnsupportedError, match='text characters of comparison work'):
        source.AdvancedFilter(1, sheet.Range('G1:G2'))
    assert state(app) == before


def test_unique_string_keys_share_the_aggregate_text_budget(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _advanced_filter
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.exceptions import VBAUnsupportedError
    from pyopenvba.interpreter._values import MISSING

    app = prepared()
    sheet = app.workbook.sheets_[0]
    source = sheet.Range('A1:B3')
    assert isinstance(source, Range)
    sheet.cells_[(2, 1)].value = 'a' * 100
    before = state(app)
    original_work = getattr(_advanced_filter, '_Work')
    monkeypatch.setattr(_advanced_filter, '_Work', lambda: original_work(50))
    with pytest.raises(VBAUnsupportedError, match='text characters of comparison work'):
        source.AdvancedFilter(2, MISSING, sheet.Range('D1:E1'), True)
    assert state(app) == before



def test_comparison_text_work_accumulates_across_rows_before_the_next_scan(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _advanced_filter
    from pyopenvba.apps.excel._model import Cell, Range
    from pyopenvba.exceptions import VBAUnsupportedError

    app = prepared()
    sheet = app.workbook.sheets_[0]
    for row in range(2, 6):
        sheet.cells_[(row, 1)] = Cell(value='a' * 20)
        sheet.cells_[(row, 2)] = Cell(value=1)
    sheet.cells_[(2, 7)].value = '*b'
    source = sheet.Range('A1:B5')
    assert isinstance(source, Range)
    before = state(app)
    original_work, original_like = getattr(_advanced_filter, '_Work'), getattr(_advanced_filter, '_like')
    monkeypatch.setattr(_advanced_filter, '_Work', lambda: original_work(65))
    scanned: list[str] = []

    def checked_like(text: str, pattern: str, *, prefix: bool, is_text: bool) -> bool:
        scanned.append(text)
        assert len(scanned) <= 2, 'third scan started after aggregate budget was exhausted'
        return original_like(text, pattern, prefix=prefix, is_text=is_text)

    monkeypatch.setattr(_advanced_filter, '_like', checked_like)
    with pytest.raises(VBAUnsupportedError, match='text characters of comparison work'):
        source.AdvancedFilter(1, sheet.Range('G1:G2'))
    assert len(scanned) == 2 and state(app) == before


@pytest.mark.parametrize('owner_name', ['Owner', "O'Brien owner"])
@pytest.mark.parametrize('filter_action', [1, 2])
@pytest.mark.parametrize('show_all_first', [False, True])
def test_deleting_advanced_filter_owner_removes_only_its_scoped_names(
        owner_name: str, filter_action: int, show_all_first: bool, tmp_path: Path) -> None:
    from zipfile import ZipFile

    from pyopenvba._a1 import split_sheet
    from pyopenvba.apps.excel._charts_file import xml_tree
    from pyopenvba.apps.excel._model import Range

    app = prepared()
    owner = app.workbook.sheets_[0]
    owner.vba_set("Name", owner_name)
    app.add_module('''Sub NamesAndSheets()
ActiveSheet.Names.Add Name:="LocalValue",RefersTo:="=1"
ActiveWorkbook.Names.Add Name:="WorkbookValue",RefersTo:="=1"
ActiveWorkbook.Names.Add Name:="WorkbookReference",RefersTo:=ActiveSheet.Range("A2")
Worksheets.Add(After:=Worksheets(1)).Name="Survivor"
ActiveSheet.Names.Add Name:="OtherLocalValue",RefersTo:="=2"
End Sub''', name='DeleteNamesSetup')
    app.run('NamesAndSheets')
    source = owner.Range('A1:B3')
    assert isinstance(source, Range)
    source.AdvancedFilter(filter_action, owner.Range('G1:G2'), owner.Range('D1:E1'))
    if show_all_first and filter_action == 1:
        owner.ShowAllData()
    expected_local = {'Criteria', '_FilterDatabase', 'LocalValue'}
    if filter_action == 2:
        expected_local.add('Extract')
    assert {split_sheet(entry.name)[1] for entry in app.workbook.names_.entries
            if split_sheet(entry.name)[0] == owner_name} == expected_local
    preserved = [(entry.name, entry.refers_to) for entry in app.workbook.names_.entries
                 if split_sheet(entry.name)[0] != owner_name]
    owner.Delete()
    assert [(entry.name, entry.refers_to) for entry in app.workbook.names_.entries] == preserved
    assert {entry.name for entry in app.workbook.names_.entries} == {
        'WorkbookValue', 'WorkbookReference', 'Survivor!OtherLocalValue'}
    saved = app.save(tmp_path / 'deleted-owner.xlsm')
    reopened = ExcelApplication.open(saved, with_vba=False)
    assert [(entry.name, entry.refers_to) for entry in reopened.workbook.names_.entries] == sorted(
        preserved, key=lambda pair: pair[0].split('!')[-1].casefold())
    with ZipFile(saved) as package:
        root = xml_tree(package.read('xl/workbook.xml'))
    names = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName')
    assert not any(element.get('name', '').startswith('_xlnm.') for element in names)
    assert next(element for element in names if element.get('name') == 'OtherLocalValue').get('localSheetId') == '0'


@pytest.mark.parametrize('case', [case for case in SAMPLES if case['name'].startswith('inplace-')],
                         ids=lambda case: case['name'])
def test_native_inplace_filter_input_owner_delete_does_not_globalize_builtins(
        case: dict[str, Any], tmp_path: Path) -> None:
    from zipfile import ZipFile

    from pyopenvba._a1 import split_sheet
    from pyopenvba.apps.excel._charts_file import xml_tree

    app = ExcelApplication.open(ROOT / case['fixture_file'], with_vba=False)
    app.add_module(case['inspection_source'], name='NativeFilterRead')
    assert app.run('Snapshot') == case['reopened']
    owner = app.workbook.sheets_[0]
    app.add_module('Sub AddSurvivor()\nWorksheets.Add(After:=Worksheets(1)).Name="Survivor"\nEnd Sub', name='DeleteOwner')
    app.run('AddSurvivor')
    owner.Delete()
    assert not any(split_sheet(entry.name)[1] in {'Criteria', 'Extract', '_FilterDatabase'}
                   for entry in app.workbook.names_.entries)
    saved = app.save(tmp_path / 'native-filter-owner-deleted.xlsm')
    with ZipFile(saved) as package:
        root = xml_tree(package.read('xl/workbook.xml'))
    assert not any(element.get('name', '').startswith('_xlnm.')
                   for element in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}definedName'))
    reopened = ExcelApplication.open(saved, with_vba=False)
    assert not reopened.workbook.names_.entries


@pytest.mark.parametrize('properties,active', [
    ('<sheetPr filterMode="1"/>', True),
    ('<s:sheetPr filterMode="true"/>', True),
    ('<sheetPr x:filterMode="1"/>', False),
    ('<x:sheetPr filterMode="1"/>', False),
    ('<x:wrapper><sheetPr filterMode="1"/></x:wrapper>', False),
    ('<sheetPr filterMode="0"/><sheetPr filterMode="1"/>', False),
    ('<x:sheetPr filterMode="0"/><s:sheetPr filterMode="1"/>', True),
])
def test_filter_state_reads_only_first_direct_main_namespace_flag(properties: str, active: bool) -> None:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._advanced_filter import read_state
    from pyopenvba.apps.excel._model import NameEntry

    app = prepared()
    sheet = app.workbook.sheets_[0]
    app.workbook.names_.entries.append(NameEntry('Sheet1!_FilterDatabase', '=Sheet1!$A$1:$B$3', app.workbook))
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    xml = ('<s:worksheet xmlns="' + ns + '" xmlns:s="' + ns + '" xmlns:x="urn:foreign">' +
           properties + '<sheetData/></s:worksheet>').encode()
    read_state(sheet, xml)
    assert sheet.advanced_filter == (Area(1, 1, 3, 2) if active else None)


@pytest.mark.parametrize('declaration', [
    '<!DOCTYPE worksheet [<!ENTITY x "1">]>',
    '<!DOCTYPE worksheet SYSTEM "file:///not-read.dtd">',
])
def test_filter_state_parser_refuses_declarations_without_mutation(declaration: str) -> None:
    from pyopenvba.apps.excel._advanced_filter import read_state

    app = prepared()
    before = state(app)
    raw = declaration + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
    with pytest.raises(ValueError, match=r'DTD|entities'):
        read_state(app.workbook.sheets_[0], raw.encode())
    assert state(app) == before
