"""Measured ordinary outlines, explicit observation subsets and atomic scope guards."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import re
from copy import deepcopy
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._model import Range
from pyopenvba.apps.excel._outline import Outline
from pyopenvba.exceptions import VBAError, VBARuntimeError, VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/worksheet_outlines'
RECORD: dict[str, Any] = json.loads((ROOT / 'initial.json').read_bytes())
CONTROLS = [c for c in RECORD['controls'] if c['runtime_admitted']]
SAMPLES = [s for s in RECORD['samples'] if 'inspection_source' in s]


@pytest.mark.parametrize('case', CONTROLS, ids=lambda c: c['name'])
def test_native_outline_projected_control(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['runtime_source'], name='OutlineControl')
    assert app.run('Probe') == case['expected_projection']


@pytest.mark.parametrize('case', SAMPLES, ids=lambda c: c['name'])
def test_native_outline_import_and_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['file'], with_vba=False)
    app.add_module(case['inspection_source'], name='OutlineInspect')
    assert app.run('Observation') == case['expected_projection']
    app.workbook.sheets_[0].touched()
    result = app.save(tmp_path / 'outlined.xlsm')
    assert ExcelApplication.open(result).run('Observation') == case['expected_projection']


def test_evidence_hashes_and_native_inputs_are_macro_free() -> None:
    provenance = json.loads((ROOT / 'provenance.json').read_bytes())
    for name, expected in provenance['files'].items():
        assert hashlib.sha256((ROOT / name).read_bytes()).hexdigest() == expected
        if name.endswith('.xlsx'):
            with ZipFile(ROOT / name) as z:
                assert not any('vbaproject' in part.casefold() for part in z.namelist())
    for case in RECORD['controls']:
        assert hashlib.sha256(case['source'].encode()).hexdigest() == case['raw_source_sha256']
        assert hashlib.sha256(case['raw_result'].encode()).hexdigest() == case['raw_result_sha256']


def snapshot(app: ExcelApplication) -> object:
    sheet = app.workbook.sheets_[0]
    return (deepcopy(sheet.dims.rows), deepcopy(sheet.dims.columns), deepcopy(sheet.dims.format),
            set(sheet.dims.changed_rows), sheet.dims.columns_changed, sheet.dims.format_changed,
            sheet.dirty, app.workbook.saved, list(app.workbook.sheets_),
            app.workbook.package.serialize() if app.workbook.package is not None else None)


def test_depth_limit_and_unsupported_operations_are_atomic() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    target = sheet.Rows('3:6')
    assert isinstance(target, Range)
    for _ in range(7):
        assert target.Group() is True
    before = snapshot(app)
    with pytest.raises(VBAError):
        target.Group()
    assert snapshot(app) == before
    for name, args in [('Group', (1,)), ('ShowDetail', (False,))]:
        if name == 'Group':
            with pytest.raises(VBAUnsupportedError):
                target.Group(*args)
        else:
            cell = sheet.Range('B7')
            assert isinstance(cell, Range)
            with pytest.raises(VBAUnsupportedError):
                cell.vba_set(name, args[0])
        assert snapshot(app) == before
    with pytest.raises(VBAUnsupportedError):
        Outline(sheet).vba_set('AutomaticStyles', True)
    assert snapshot(app) == before


def test_large_axis_rejected_before_materialization_and_mixed_queries_candid() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    before = snapshot(app)
    all_rows = sheet.Rows('1:1048576')
    assert isinstance(all_rows, Range)
    with pytest.raises(VBAUnsupportedError, match='50000'):
        all_rows.Group()
    assert snapshot(app) == before
    grouped = sheet.Rows('3:6')
    mixed = sheet.Rows('3:8')
    assert isinstance(grouped, Range) and isinstance(mixed, Range)
    grouped.Group()
    with pytest.raises(VBAUnsupportedError, match='heterogeneous'):
        mixed.OutlineLevel()


def test_structural_and_whole_axis_copy_guards_preserve_all_state() -> None:
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    grouped = sheet.Rows('3:6')
    destination = sheet.Rows('10:13')
    assert isinstance(grouped, Range) and isinstance(destination, Range)
    grouped.Group()
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='structural'):
        grouped.Insert()
    assert snapshot(app) == before
    with pytest.raises(VBAUnsupportedError, match='whole-axis'):
        grouped.copy_to(destination)
    assert snapshot(app) == before
    unrelated = sheet.Rows('20:20')
    assert isinstance(unrelated, Range)
    unrelated.Insert()
    assert sheet.dims.rows[3].extra['outlineLevel'] == '1'


@pytest.fixture(scope='module')
def native_outline_excel() -> Any:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='opt-in native outline output gate')
@pytest.mark.parametrize('case', CONTROLS, ids=lambda c: c['name'])
def test_generated_outline_file_reopened_by_native(case: dict[str, Any], tmp_path: Path,
                                                  native_outline_excel: Any) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['runtime_source'], name='OutlineControl')
    assert app.run('Probe') == case['expected_projection']
    phases = re.findall(r'\[(sheet=[^\]]*)\]', case['expected_projection'])
    expected = phases[-1] if phases else case['expected_projection'].split(';creator=')[0]
    if case['name'] == 'worksheet-copy':
        app.add_module(case['runtime_source'].split('Function Probe()')[0], name='OutlineControl', workbook=app.workbook)
    result = app.save(tmp_path / 'native-output.xlsm')
    native_outline_excel.open_document(result)
    assert native_outline_excel.run_raw('OutlineControl.Observation') == expected


def imported_xml(tmp_path: Path, xml: str) -> ExcelApplication:
    path = tmp_path / 'synthetic-outline-metadata.xlsx'
    with ZipFile(ROOT / 'basic-rows.xlsx') as original, ZipFile(path, 'w') as changed:
        for info in original.infolist():
            changed.writestr(info, xml.encode() if info.filename == 'xl/worksheets/sheet1.xml' else original.read(info.filename))
    return ExcelApplication.open(path, with_vba=False)


def test_metadata_only_row_edits_preserve_cell_bytes_and_unknown_nested_payload(tmp_path: Path) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    cell = "<c r='A4' t='n'><v>1.0000</v></c>"
    opaque = "<x:rowPayload x:keep='retain &amp; literal'/>"
    xml = (f'<worksheet xmlns="{ns}" xmlns:x="urn:opaque"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
           '<sheetData><row r=\'4\' ht=\'19.00\' customHeight=\'1\' x:keep=\'hidden="0" > retain\'>'
           + cell + opaque + '</row><row r="9"><c r="A9"><v>9</v></c></row></sheetData>'
           '<headerFooter><oddHeader>&amp;RUnchanged</oddHeader></headerFooter></worksheet>')
    app = imported_xml(tmp_path, xml)
    sheet = app.workbook.sheets_[0]
    grouped = sheet.Rows('3:6')
    assert isinstance(grouped, Range)
    grouped.Group()
    assert not sheet.dirty
    first = app.save(tmp_path / 'first.xlsm')
    with ZipFile(first) as z:
        saved = z.read(sheet.part_name)
        assert cell.encode() in saved and opaque.encode() in saved
        assert b"x:keep='hidden=\"0\" > retain'" in saved
        assert b'<pageSetUpPr fitToPage="1"/>' in saved
        assert b'<headerFooter><oddHeader>&amp;RUnchanged</oddHeader></headerFooter>' in saved
    second = app.save(tmp_path / 'second.xlsm')
    with ZipFile(second) as z:
        assert z.read(sheet.part_name) == saved
    reopened = ExcelApplication.open(second, with_vba=False)
    row = reopened.workbook.sheets_[0].Rows(4)
    assert isinstance(row, Range) and row.OutlineLevel() == 2


def test_column_span_split_preserves_width_local_namespace_and_opaque_attrs(tmp_path: Path) -> None:
    from pyopenvba.apps.excel._charts_file import xml_tree

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    xml = (f'<worksheet xmlns="{ns}"><cols xmlns:x="urn:column">'
           '<col min=\'2\' max=\'7\' width=\'12.500\' customWidth=\'1\' x:keep=\'hidden="0" > retain\'/>'
           '</cols><sheetData/></worksheet>')
    app = imported_xml(tmp_path, xml)
    sheet = app.workbook.sheets_[0]
    grouped = sheet.Columns('C:E')
    assert isinstance(grouped, Range)
    grouped.Group()
    assert not sheet.dirty
    result = app.save(tmp_path / 'split.xlsm')
    with ZipFile(result) as z:
        saved = z.read(sheet.part_name)
    root = xml_tree(saved)
    columns = root.find('{' + ns + '}cols')
    assert columns is not None
    for column in columns:
        assert column.get('width') == '12.500'
        assert column.get('{urn:column}keep') == 'hidden="0" > retain'
        first, last = int(column.attrib['min']), int(column.attrib['max'])
        assert column.get('outlineLevel', '0') == ('1' if 3 <= first <= last <= 5 else '0')
    assert b'xmlns:x="urn:column"' in saved


@pytest.mark.parametrize('foreign', [True, False])
def test_axis_xml_admission_failure_is_atomic(tmp_path: Path, foreign: bool) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    row = '<row xmlns="" r="3" outlineLevel="1"/>' if foreign else '<s:row r="3" outlineLevel="1"/>'
    xml = f'<worksheet xmlns="{ns}" xmlns:s="{ns}"><sheetData>' + row + '</sheetData></worksheet>'
    app = imported_xml(tmp_path, xml)
    before = snapshot(app)
    grouped = app.workbook.sheets_[0].Rows('3:6')
    assert isinstance(grouped, Range)
    with pytest.raises(VBAUnsupportedError, match='XML'):
        grouped.Group()
    assert snapshot(app) == before


def test_configuration_namespace_ownership_and_page_setup_sibling_preservation(tmp_path: Path) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    data = '<s:sheetData><s:row r="3"><s:c r="A3"><s:v>1.000</s:v></s:c></s:row></s:sheetData>'
    foreign = '<outlinePr xmlns="" summaryBelow="0" opaque="retain"/>'
    xml = (f'<s:worksheet xmlns:s="{ns}" xmlns:x="urn:unknown"><s:sheetPr x:keep=\'summaryBelow="0" > retain\'>'
           + foreign + '<s:pageSetUpPr fitToPage="1"/></s:sheetPr>' + data + '</s:worksheet>')
    app = imported_xml(tmp_path, xml)
    outline = Outline(app.workbook.sheets_[0])
    assert outline.SummaryRow() == 1
    outline.vba_set('SummaryRow', 0)
    result = app.save(tmp_path / 'flags-only.xlsm')
    with ZipFile(result) as z:
        saved = z.read(app.workbook.sheets_[0].part_name).decode()
    assert data in saved and foreign in saved
    assert '<s:outlinePr summaryBelow="0"/>' in saved or '<s:outlinePr  summaryBelow="0"/>' in saved
    assert saved.index('<s:outlinePr') < saved.index('<s:pageSetUpPr')
    assert 'x:keep=\'summaryBelow="0" > retain\'' in saved


def test_invalid_unrelated_stored_depth_does_not_partially_mutate_target() -> None:
    from pyopenvba.apps.excel._dimensions import RowRecord

    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    sheet.dims.rows[99] = RowRecord(extra={'outlineLevel': 'invalid'})
    grouped = sheet.Rows('3:6')
    assert isinstance(grouped, Range)
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='stored'):
        grouped.Group()
    assert snapshot(app) == before


def test_outline_visibility_recalculates_subtotal_without_cell_edits() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function Probe() As String
Dim n As Long,s As String
For n=2 To 8
Cells(n,1).Value=n
Next
Range("B1").Formula="=SUBTOTAL(109,A2:A8)"
s=Range("B1").Value2
Rows("3:6").Group
Rows(7).ShowDetail=False
Probe=s & "|" & Range("B1").Value2
End Function''')
    assert app.run('Probe') == '35|17'


@pytest.mark.parametrize('case', CONTROLS, ids=lambda c: c['name'])
def test_generated_outline_model_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['runtime_source'], name='OutlineControl')
    assert app.run('Probe') == case['expected_projection']
    expected = app.run('Observation')
    if case['name'] == 'worksheet-copy':
        app.add_module(case['runtime_source'].split('Function Probe()')[0], name='OutlineControl', workbook=app.workbook)
    saved = app.save(tmp_path / 'generated.xlsm')
    assert ExcelApplication.open(saved).run('Observation') == expected


def test_combined_showlevels_preflights_both_axes_before_mutation(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _outline_file

    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    app.add_module('Sub Setup()\nRows("3:6").Group\nColumns("C:F").Group\nEnd Sub')
    app.run('Setup')
    before = snapshot(app)
    original = _outline_file.patch

    def reject_columns(*args: Any, **kwargs: Any) -> str:
        assert kwargs['prospective_rows'][3].hidden
        assert kwargs['prospective_columns'][3].hidden
        raise VBAUnsupportedError('controlled combined output budget')

    monkeypatch.setattr(_outline_file, 'patch', reject_columns)
    with pytest.raises(VBAUnsupportedError, match='combined'):
        Outline(sheet).ShowLevels(1, 1)
    assert snapshot(app) == before
    monkeypatch.setattr(_outline_file, 'patch', original)
    assert Outline(sheet).ShowLevels(1, 1) is True


def test_column_attribute_amplification_budget_is_atomic(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _outline_file

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    xml = (f'<worksheet xmlns="{ns}"><cols><col min="2" max="7" width="12" customWidth="1" opaque="'
           + 'x' * 2000 + '"/></cols><sheetData/></worksheet>')
    app = imported_xml(tmp_path, xml)
    before = snapshot(app)
    monkeypatch.setattr(_outline_file, 'MAX_BYTES', len(xml.encode()) + 500)
    grouped = app.workbook.sheets_[0].Columns('C:E')
    assert isinstance(grouped, Range)
    with pytest.raises(VBAUnsupportedError):
        grouped.Group()
    assert snapshot(app) == before


def test_existing_cell_rows_count_toward_metadata_budget(monkeypatch: pytest.MonkeyPatch) -> None:
    from pyopenvba.apps.excel import _outline

    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Setup()\nRange("A1").Value=1\nRange("A2").Value=2\nRange("A3").Value=3\nEnd Sub')
    app.run('Setup')
    before = snapshot(app)
    monkeypatch.setattr(_outline, 'LIMIT', 3)
    target = app.workbook.sheets_[0].Rows(4)
    assert isinstance(target, Range)
    with pytest.raises(VBAUnsupportedError):
        target.Group()
    assert snapshot(app) == before


def test_paste_formats_whole_axis_outline_refusal_is_atomic() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Sub Setup()\nRows("3:6").Group\nRows("3:6").Copy\nEnd Sub\nSub Paste()\nRows("9:12").PasteSpecial Paste:=xlPasteFormats\nEnd Sub')
    app.run('Setup')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='outline metadata'):
        app.run('Paste')
    assert snapshot(app) == before


def test_structural_edit_cannot_drop_unparsed_qualified_outline_metadata(tmp_path: Path) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    xml = (f'<worksheet xmlns="{ns}" xmlns:s="{ns}"><sheetData>'
           '<s:row r="3" outlineLevel="1"/></sheetData></worksheet>')
    app = imported_xml(tmp_path, xml)
    target = app.workbook.sheets_[0].Rows(2)
    assert isinstance(target, Range)
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='XML'):
        target.Insert()
    assert snapshot(app) == before


def test_protection_ui_only_and_worksheet_udf_measured_noops() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('''Function UGroup() As Boolean
UGroup=Application.Caller.Parent.Rows("3:6").Group
End Function
Function ULevels() As Boolean
ULevels=Application.Caller.Parent.Outline.ShowLevels(1,0)
End Function
Function USummary() As Long
Application.Caller.Parent.Outline.SummaryRow=xlAbove
USummary=Application.Caller.Parent.Outline.SummaryRow
End Function
Sub ProtectSheet()
ActiveSheet.Protect
End Sub
Sub UIOnly()
ActiveSheet.Unprotect
ActiveSheet.Protect UserInterfaceOnly:=True
End Sub
Function UDFProbe() As String
ActiveSheet.Unprotect
Range("A1").Formula2="=UGroup()"
Range("A2").Formula2="=ULevels()"
Range("B1").Formula2="=USummary()"
UDFProbe=Range("A1").Value2 & "|" & Range("A2").Value2 & "|" & Range("B1").Value2
End Function''')
    app.run('ProtectSheet')
    before = snapshot(app)
    target = app.workbook.sheets_[0].Rows('3:6')
    assert isinstance(target, Range)
    for callback in (target.Group, lambda: Outline(target.sheet).ShowLevels(1, 0),
                     lambda: Outline(target.sheet).vba_set('SummaryRow', 0)):
        with pytest.raises(VBARuntimeError) as caught:
            callback()
        assert caught.value.number == 1004
        assert snapshot(app) == before
    app.run('UIOnly')
    assert target.Group() is True
    assert app.run('UDFProbe') == 'False|False|1'
    assert target.OutlineLevel() == 2
    assert Outline(target.sheet).SummaryRow() == 1


@pytest.mark.parametrize('rows', [True, False], ids=['qualified-rows', 'qualified-columns'])
def test_sheet_copy_cannot_drop_unparsed_qualified_outline_metadata(tmp_path: Path, rows: bool) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    axes = ('<sheetData><s:row r="3" outlineLevel="1"/></sheetData>' if rows else
            '<cols><s:col min="3" max="5" width="12" outlineLevel="1"/></cols><sheetData/>')
    app = imported_xml(tmp_path, f'<worksheet xmlns="{ns}" xmlns:s="{ns}">' + axes + '</worksheet>')
    app.add_module('Sub Probe()\nActiveSheet.Copy After:=ActiveSheet\nEnd Sub')
    before = snapshot(app)
    with pytest.raises(VBAUnsupportedError, match='XML'):
        app.run('Probe')
    assert snapshot(app) == before


def test_many_column_intervals_preserve_independent_metadata() -> None:
    from pyopenvba.apps.excel._charts_file import xml_tree
    from pyopenvba.apps.excel._dimensions import ColumnRecord
    from pyopenvba.apps.excel._outline_file import patch

    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    count = 1024
    intervals = ''.join(f'<col min="{n}" max="{n}" width="12" customWidth="1" opaque="keep-{n}"/>'
                        for n in range(1, count + 1))
    xml = f'<worksheet xmlns="{ns}"><cols>' + intervals + '</cols><sheetData/></worksheet>'
    app = ExcelApplication()
    app.add_workbook()
    sheet = app.workbook.sheets_[0]
    sheet.dims.columns = {n: ColumnRecord(width=84, custom=True, extra={'outlineLevel': str(n % 7 + 1)})
                          for n in range(1, count + 1)}
    result = patch(sheet, xml, set(), set(range(1, count + 1)))
    columns = xml_tree(result.encode()).find('{' + ns + '}cols')
    assert columns is not None and len(columns) == count
    for n, column in enumerate(columns, 1):
        assert column.attrib == {'min': str(n), 'max': str(n), 'width': '12', 'customWidth': '1',
                                 'opaque': f'keep-{n}', 'outlineLevel': str(n % 7 + 1)}


@pytest.mark.parametrize('rows', [True, False], ids=['rows', 'columns'])
@pytest.mark.parametrize('cdata', [True, False], ids=['cdata', 'comment'])
def test_lexical_phantom_outline_metadata_is_rejected_atomically(tmp_path: Path, rows: bool, cdata: bool) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    phantom = '<row r="9" outlineLevel="7"/>' if rows else '<col min="9" max="9" width="12" collapsed="1"/>'
    opaque = '<![CDATA[' + phantom + ']]>' if cdata else '<!-- ' + phantom + ' -->'
    axes = '<sheetData>' + opaque + '</sheetData>' if rows else '<cols>' + opaque + '</cols><sheetData/>'
    app = imported_xml(tmp_path, f'<worksheet xmlns="{ns}">' + axes + '</worksheet>')
    app.add_module('Sub CopySheet()\nActiveSheet.Copy After:=ActiveSheet\nEnd Sub')
    target = app.workbook.sheets_[0].Rows(9) if rows else app.workbook.sheets_[0].Columns(9)
    assert isinstance(target, Range)
    before = snapshot(app)
    for callback in (target.OutlineLevel, target.Group, lambda: app.run('CopySheet')):
        with pytest.raises(VBAUnsupportedError, match='XML'):
            callback()
        assert snapshot(app) == before


def test_admission_after_save_accepts_real_pending_outline_and_hidden_edits(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    target = app.workbook.sheets_[0].Rows('3:6')
    assert isinstance(target, Range)
    target.Group()
    app.save(tmp_path / 'first.xlsm')
    target.vba_set('Hidden', True)
    assert target.OutlineLevel() == 2
    target.Group()
    app.save(tmp_path / 'second.xlsm')
    assert target.OutlineLevel() == 3


@pytest.mark.parametrize('rows', [True, False], ids=['rows', 'columns'])
@pytest.mark.parametrize('cdata', [True, False], ids=['cdata', 'comment'])
@pytest.mark.parametrize('stored', ['outlineLevel="1"', 'collapsed="true"'], ids=['level', 'collapsed'])
def test_real_outline_overwritten_by_lexical_default_metadata_is_rejected(
        tmp_path: Path, rows: bool, cdata: bool, stored: str) -> None:
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    real = f'<row r="9" {stored}/>' if rows else f'<col min="9" max="10" width="12" {stored}/>'
    # Nondefault height materializes the fake row while leaving outline fields
    # default. A column record always materializes, including default outlines.
    phantom = '<row r="9" ht="16" customHeight="1"/>' if rows else '<col min="9" max="9" width="12"/>'
    opaque = '<![CDATA[' + phantom + ']]>' if cdata else '<!-- ' + phantom + ' -->'
    axes = '<sheetData>' + real + opaque + '</sheetData>' if rows else '<cols>' + real + opaque + '</cols><sheetData/>'
    app = imported_xml(tmp_path, f'<worksheet xmlns="{ns}">' + axes + '</worksheet>')
    app.add_module('Sub CopySheet()\nActiveSheet.Copy After:=ActiveSheet\nEnd Sub')
    target = app.workbook.sheets_[0].Rows(9) if rows else app.workbook.sheets_[0].Columns(9)
    assert isinstance(target, Range)
    before = snapshot(app)
    for callback in (target.OutlineLevel, target.Group, lambda: app.run('CopySheet')):
        with pytest.raises(VBAUnsupportedError, match='XML'):
            callback()
        assert snapshot(app) == before
