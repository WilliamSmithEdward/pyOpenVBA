"""Independent native multiple-rule transfers, identities and clipped populations."""
from __future__ import annotations

import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any
from zipfile import ZipFile

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS
from pyopenvba.exceptions import VBAUnsupportedError

ROOT = Path(__file__).parent / 'fixtures/conditional_multiple'
CASES: list[dict[str, Any]] = json.loads((ROOT / 'native.json').read_text(encoding='utf-8'))['probes']


def generated(case: dict[str, Any]) -> ExcelApplication:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeControl')
    assert app.run('Probe') == case['result']
    assert app.run('Snapshot') == case['runtime_final']
    return app


def serialized_rules(xml: str) -> list[tuple[str, str, str, tuple[str, ...]]]:
    root = xml_tree(xml)
    return [(rule.get('type', ''), rule.get('priority', ''), group.get('sqref', ''),
             tuple(child.text or '' for child in rule.findall(NS + 'formula')))
            for group in root.findall(NS + 'conditionalFormatting') for rule in group.findall(NS + 'cfRule')]


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_multiple_runtime_and_macro_enabled_roundtrip(case: dict[str, Any], tmp_path: Path) -> None:
    app = generated(case)
    path = app.save(tmp_path / 'multiple.xlsm')
    reopened = ExcelApplication.open(path)
    assert reopened.run('Snapshot') == case['reopened52']
    # Collection priority order and persisted creation order are distinct native
    # observations: compare the full independent XML sequence as well as the API.
    with ZipFile(path) as package:
        for part, native in case['xml52'].items():
            if part.startswith('xl/worksheets/sheet'):
                assert serialized_rules(package.read(part).decode('utf-8')) == serialized_rules(native)


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_multiple_macro_free_import_and_conversion(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication.open(ROOT / case['file51'], with_vba=False)
    app.add_module(case['inspection_source'], name='Inspection')
    assert app.run('Snapshot') == case['reopened51']
    for sheet in app.workbook.sheets_:
        if sheet.conditional_formats:
            sheet.conditional_changed = True
            sheet.touched()
    reopened = ExcelApplication.open(app.save(tmp_path / 'converted.xlsm'))
    assert reopened.run('Snapshot') == case['reopened51']


def test_reverse_priority_copy_keeps_creation_sequence_distinct(tmp_path: Path) -> None:
    case = next(case for case in CASES if case['name'] == 'reverse-priority-two')
    app = generated(case)
    with ZipFile(app.save(tmp_path / 'creation-order.xlsm')) as package:
        sequence = serialized_rules(package.read('xl/worksheets/sheet1.xml').decode('utf-8'))
    copied = [rule for rule in sequence if rule[2] == 'F3:G5']
    assert [rule[1] for rule in copied] == ['2', '1']
    assert [rule[3] for rule in copied] == [('F3>0',), ('F3<8',)]


def test_clipped_top_two_uses_only_the_copied_population() -> None:
    case = next(case for case in CASES if case['name'] == 'three-mixed-clipped')
    app = generated(case)
    app.add_module('''Function Population() As String
Population=CStr(Worksheets("Data").Range("G3").Value2) & ":" & CStr(Worksheets("Data").Range("G4").Value2)
Population=Population & ":" & CStr(Worksheets("Data").Range("G3").DisplayFormat.Interior.Color)
Population=Population & ":" & CStr(Worksheets("Data").Range("G4").DisplayFormat.Interior.Color)
End Function''', name='PopulationRead')
    assert app.run('Population') == '3:6:65535:65535'


def test_held_source_edits_do_not_change_any_copied_rule_style() -> None:
    case = next(case for case in CASES if case['name'] == 'three-mixed-clipped')
    app = generated(case)
    app.add_module('''Function Independent() As String
Dim f As Object,s As String
For Each f In Worksheets("Data").Cells.FormatConditions
If f.Type=2 And f.AppliesTo.Address="$A$1:$C$6" Then s=s & "sourceBold=" & CStr(f.Font.Bold) & ";"
If f.Type=2 And f.AppliesTo.Address="$F$3:$G$5" Then s=s & "copyBold=" & CStr(f.Font.Bold) & ";"
If f.Type=9 And f.AppliesTo.Address="$B$2:$C$5" Then s=s & "sourceItalic=" & CStr(f.Font.Italic) & ";"
If f.Type=9 And f.AppliesTo.Address="$F$3:$G$5" Then s=s & "copyItalic=" & CStr(f.Font.Italic) & ";"
If f.Type=5 And f.AppliesTo.Address="$G$3:$G$5" Then s=s & "copyFill=" & CStr(f.Interior.Color) & ";"
If f.Type=5 And f.AppliesTo.Address="$A$2:$A$5,$C$1:$C$4" Then s=s & "sourceFill=" & CStr(f.Interior.Color) & ";"
Next
Independent=s
End Function''', name='HeldRead')
    flags = set(str(app.run('Independent')).strip(';').split(';'))
    assert flags == {'sourceBold=False', 'copyBold=True', 'sourceItalic=False',
                     'copyItalic=True', 'sourceFill=16776960', 'copyFill=65535'}


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_promoted_multiple_samples_are_macro_free(case: dict[str, Any]) -> None:
    with ZipFile(ROOT / case['file51']) as package:
        assert not any('vbaproject' in name.lower() for name in package.namelist())


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1', reason='requires explicitly enabled native Excel')
@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_native_reopens_generated_multiple_rules(case: dict[str, Any], tmp_path: Path, native_excel: Any) -> None:
    app = generated(case)
    native_excel.open_document(app.save(tmp_path / 'native-multiple.xlsm'))
    assert native_excel.run_raw('NativeControl.Snapshot') == case['reopened52']
    assert not any(event.get('dialogs') for event in native_excel.events)

@pytest.mark.parametrize('payload', ['unsupported-family', 'duplicate-priority', 'numeric-duplicate-priority'])
def test_one_bad_rule_aborts_the_entire_multiple_transfer(payload: str) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Sub Setup()
Dim f As Object,g As Object,h As Object
Range("A1:C3").Value2=7
Range("F3:H5").Value2=11
Set f=Range("A1:C3").FormatConditions.Add(2,Formula1:="=A1>0")
f.Font.Bold=True
Set g=Range("A1:C3").FormatConditions.AddTop10
g.Interior.Color=vbYellow
Set h=Range("A1:C3").FormatConditions.AddUniqueValues
h.Font.Italic=True
End Sub
Sub CopyNow()
Range("A1:C3").Copy Destination:=Range("F3:H5")
End Sub''', name='AtomicHarness')
    app.run('Setup')
    sheet = book.sheets_[0]
    if payload == 'unsupported-family':
        sheet.conditional_formats[-1].root.set('type', 'dataBar')
    else:
        priority = '01' if payload == 'numeric-duplicate-priority' else sheet.conditional_formats[0].root.get('priority', '1')
        sheet.conditional_formats[-1].root.set('priority', priority)

    def state() -> object:
        return ({key: (cell.value, cell.formula, cell.style) for key, cell in sheet.cells_.items()},
                [(id(rule), tuple(rule.areas), xml_bytes(rule.root), rule.deleted,
                  None if rule.dxf is None else xml_bytes(rule.dxf), rule.color_slot)
                 for rule in sheet.conditional_formats],
                sheet.conditional_color_next, tuple(sheet.conditional_color_slots.items()),
                book.stylesheet.dxf_count,
                None if sheet.conditional_color_groups is None else tuple(
                    (slot, tuple(members), tuple(ranks.items()))
                    for slot, (members, ranks) in sheet.conditional_color_groups.items()))

    before = state()
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state() == before


def test_copied_sheet_ambiguous_multiple_rule_identities_are_atomic() -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Sub Setup()
Dim f As Object,g As Object,h As Object
ActiveSheet.Name="Data"
Range("A1:A3").Value2=5
Set f=Range("A1:A3").FormatConditions.Add(2,Formula1:="=A1>0")
Set g=Range("A1:A2").FormatConditions.Add(2,Formula1:="=A1<8")
ActiveSheet.Copy After:=ActiveSheet
Set h=Range("A1:A3").FormatConditions.AddTop10
h.Rank=1
Range("F1:F3").Value2=11
End Sub
Sub CopyNow()
Range("A1:A3").Copy Destination:=Range("F1:F3")
End Sub''', name='IdentityHarness')
    app.run('Setup')
    sheet = book.sheets_[-1]
    assert [rule.color_slot for rule in sheet.conditional_formats] == [1, 2, 3]
    # A copied worksheet now has its own dense registry. Construct genuinely
    # inconsistent private provenance to keep the transfer admission failure
    # meaningful, rather than relying on the previously missing initialization.
    third = sheet.conditional_formats[-1]
    third.color_slot = 1
    sheet.conditional_color_slots.pop(3)
    sheet.conditional_color_slots[1] = third
    assert [rule.color_slot for rule in sheet.conditional_formats] == [1, 2, 1]

    def state() -> object:
        return ({key: (cell.value, cell.formula, cell.style) for key, cell in sheet.cells_.items()},
                [(id(rule), tuple(rule.areas), xml_bytes(rule.root), rule.deleted,
                  None if rule.dxf is None else xml_bytes(rule.dxf), rule.color_slot)
                 for rule in sheet.conditional_formats],
                sheet.conditional_color_next, tuple(sheet.conditional_color_slots.items()),
                book.stylesheet.dxf_count,
                None if sheet.conditional_color_groups is None else tuple(
                    (slot, tuple(members), tuple(ranks.items()))
                    for slot, (members, ranks) in sheet.conditional_color_groups.items()))

    before = state()
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state() == before


PROPERTY_RECORDS: dict[str, Any] = json.loads((ROOT / 'specialized_properties.json').read_text(encoding='utf-8'))
PROPERTY_CASES = [(group, case) for group, record in PROPERTY_RECORDS.items() for case in record['probes']]


@pytest.mark.parametrize('group,case', PROPERTY_CASES, ids=[group + '-' + case['name'] for group, case in PROPERTY_CASES])
def test_native_specialized_scalar_properties_and_owner_setters(group: str, case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeProperty')
    assert app.run('Probe') == case['result']
