"""Conditional replacement planning, independent styles and atomic admission."""
from __future__ import annotations

import json
from copy import copy, deepcopy
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication, _events
from pyopenvba.apps.excel import _conditional_transfer as transfer
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule
from pyopenvba.apps.excel._conditional_transfer import plan
from pyopenvba.apps.excel._model import Range, Worksheet
from pyopenvba.apps.excel._udf import running
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

_ROOT = Path(__file__).parent / 'fixtures/conditional_transfer'
_UDF_CASES: list[dict[str, Any]] = json.loads((_ROOT / 'udf.json').read_text(encoding='utf-8'))['probes']
_UDF_RETURN_CASES: list[dict[str, Any]] = json.loads((_ROOT / 'udf_returns.json').read_text(encoding='utf-8'))['probes']


def prepared() -> tuple[ExcelApplication, Worksheet]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('''Public Held As Object
Sub Setup()
Dim g As Object
Range("A1:C4").Value=5
Range("E1:H6").Value=8
Set Held=Range("A1:C4").FormatConditions.Add(xlExpression,Formula1:="=A1>0")
Held.Font.Bold=True
Set g=Range("E1:H6").FormatConditions.Add(xlExpression,Formula1:="=E1<0")
g.Font.Italic=True
End Sub
Sub CopyNow()
Range("A1:C3").Copy Range("F3:H5")
End Sub
Function Mutation() As String
Held.Font.Bold=False
Mutation=CStr(Held.Font.Bold) & ":" & CStr(Range("F3").FormatConditions(1).Font.Bold)
End Function''', name='Harness')
    app.run('Setup')
    return app, book.sheets_[0]


def state(sheet: Worksheet) -> object:
    return ({position: copy(cell) for position, cell in sheet.cells_.items()},
            [(id(rule), list(rule.areas), xml_bytes(rule.root),
                                     None if rule.dxf is None else xml_bytes(rule.dxf), rule.deleted, rule.color_slot)
                                    for rule in sheet.conditional_formats],
            sheet.book.stylesheet.dxf_count, list(sheet.book.names_.entries), sheet.dirty,
            sheet.conditional_changed, sheet.book.saved, sheet.conditional_color_next,
            [(slot, id(rule)) for slot, rule in sheet.conditional_color_slots.items()])


def test_planning_is_pure_and_commit_preserves_existing_proxies_and_dxf_snapshots() -> None:
    app, sheet = prepared()
    original, destination = sheet.conditional_formats
    before = state(sheet)
    staged = plan(sheet, [Area(1, 1, 3, 3)], sheet, [Area(3, 6, 5, 8)])
    assert staged is not None and state(sheet) == before
    assert staged.incoming is not None and staged.incoming.root is not original.root
    assert staged.incoming.dxf is not original.dxf
    assert staged.incoming.root.findtext(NS + 'formula') == 'F3>0'
    assert staged.incoming.imported is False
    assert staged.incoming.color_slot is None
    staged.commit()
    assert original is sheet.conditional_formats[0] and destination is sheet.conditional_formats[1]
    assert destination.areas == [Area(1, 5, 2, 8), Area(6, 5, 6, 8), Area(3, 5, 5, 5)]
    assert app.run('Mutation') == 'False:True'


def test_destination_holes_are_cleared_and_one_statistical_population_is_clipped() -> None:
    _, sheet = prepared()
    source = sheet.conditional_formats[0]
    source.areas = [Area(1, 1, 4, 1), Area(1, 3, 4, 3)]
    source.root = XmlElement(NS + 'cfRule', {'type': 'top10', 'rank': '2', 'priority': '2'})
    sheet.conditional_formats[1].root.set('priority', '7')
    staged = plan(sheet, [Area(2, 1, 3, 3)], sheet, [Area(5, 6, 6, 8)])
    assert staged is not None and staged.incoming is not None
    staged.commit()
    assert staged.incoming.areas == [Area(5, 6, 6, 6), Area(5, 8, 6, 8)]
    assert not any(area.contains(5, 7) for rule in sheet.conditional_formats for area in rule.areas)
    assert [rule.root.get('priority') for rule in sheet.conditional_formats] == ['3', '8', '1']


def test_union_predicate_copy_regenerates_bounding_anchor_instead_of_first_landed_area() -> None:
    _, sheet = prepared()
    source = sheet.conditional_formats[0]
    sheet.conditional_formats[:] = [source]
    source.areas = [Area(9, 2, 10, 2), Area(3, 11, 5, 11)]
    source.root = XmlElement(NS + 'cfRule', {'type': 'containsText', 'text': 'x', 'operator': 'containsText', 'priority': '1'})
    staged = plan(sheet, [Area(3, 2, 10, 11)], sheet, [Area(20, 14, 27, 23)])
    assert staged is not None and staged.incoming is not None
    assert staged.incoming.areas == [Area(26, 14, 27, 14), Area(20, 23, 22, 23)]
    assert staged.incoming.root.findtext(NS + 'formula') == 'NOT(ISERROR(SEARCH("x",N20)))'


def test_no_source_format_still_replaces_entire_destination_without_priority_bump() -> None:
    _, sheet = prepared()
    staged = plan(sheet, [Area(1, 10, 3, 12)], sheet, [Area(2, 5, 4, 7)])
    assert staged is not None and staged.incoming is None
    staged.commit()
    assert [rule.root.get('priority') for rule in sheet.conditional_formats] == ['1', '2']
    assert not any(area.contains(3, 6) for rule in sheet.conditional_formats for area in rule.areas)


@pytest.mark.parametrize('payload', ['attribute', 'child', 'namespaced-child', 'container', 'bad-formula',
                                     'bad-cell-formula', 'multiple', 'shared-rule', 'dxf-extension', 'dxf-attribute'])
def test_unsupported_payloads_reject_before_cells_styles_names_dirty_flags_or_events(
        payload: str, monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = prepared()
    source = sheet.conditional_formats[0]
    if payload == 'attribute':
        source.root.set('unknown', 'opaque')
    elif payload == 'child':
        source.root.append(XmlElement(NS + 'extLst'))
    elif payload == 'namespaced-child':
        source.root.append(XmlElement('{urn:extension}custom'))
    elif payload == 'container':
        source.container = XmlElement(NS + 'conditionalFormatting', {'pivot': '1'})
    elif payload == 'bad-formula':
        child = source.root.find(NS + 'formula')
        assert child is not None
        child.text = '"'
    elif payload == 'bad-cell-formula':
        sheet.cells_[(1, 1)].formula = '="'
    elif payload == 'multiple':
        sheet.conditional_formats.append(Rule(list(source.areas), deepcopy(source.root)))
    elif payload == 'shared-rule':
        source.areas = [Area(1, 1, 10, 12)]
    else:
        assert source.dxf is not None
        if payload == 'dxf-attribute':
            source.dxf.set('{urn:extension}id', 'opaque')
        else:
            source.dxf.append(XmlElement('{urn:extension}custom'))
    events: list[object] = []

    def after_edit(target: object, **kwargs: object) -> None:
        events.append(target)

    monkeypatch.setattr(_events, 'after_edit', after_edit)
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state(sheet) == before and events == []


@pytest.mark.parametrize('kind,turned,operation', [(14, False, ''), (13, False, ''), (7, False, ''),
                                                 (-4122, True, ''), (-4104, False, '+')])
def test_unmeasured_paste_modes_fail_during_pure_admission(kind: int, turned: bool, operation: str) -> None:
    _, sheet = prepared()
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        plan(sheet, [Area(1, 1, 3, 3)], sheet, [Area(3, 6, 5, 8)],
             kind=kind, transpose=turned, operation=operation)
    assert state(sheet) == before


@pytest.mark.parametrize('landing', [Area(2, 2, 4, 4), Area(1, 6, 6, 11)])
def test_overlap_and_tiled_mappings_reject_atomically(landing: Area) -> None:
    _, sheet = prepared()
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        plan(sheet, [Area(1, 1, 3, 3)], sheet, [landing])
    assert state(sheet) == before


def test_cross_workbook_conditional_copy_fails_without_importing_styles_or_names() -> None:
    app, source = prepared()
    destination = app.add_workbook().sheets_[0]
    before = state(source), state(destination)
    with pytest.raises(VBAUnsupportedError):
        plan(source, [Area(1, 1, 3, 3)], destination, [Area(3, 6, 5, 8)])
    assert (state(source), state(destination)) == before


def test_simple_skip_blanks_formats_does_not_punch_a_conditional_hole() -> None:
    app, sheet = prepared()
    app.add_module('''Sub PasteBlank()
Range("B2").ClearContents
Range("A1:C3").Copy
Range("F3:H5").PasteSpecial Paste:=xlPasteFormats, SkipBlanks:=True
End Sub''', name='PasteModule')
    app.run('PasteBlank')
    incoming = next(rule for rule in sheet.conditional_formats if rule.root.get('priority') == '1')
    assert incoming.areas == [Area(3, 6, 5, 8)]
    assert incoming.root.findtext(NS + 'formula') == 'F3>0'


def test_worksheet_x14_rules_block_priority_changes_before_any_mutation(tmp_path: Path) -> None:
    app, sheet = prepared()
    app.save(tmp_path / 'before.xlsm')
    package = sheet.book.package
    assert package is not None
    root = xml_tree(package.read(sheet.part_name))
    extension = XmlElement(NS + 'extLst')
    extension.append(XmlElement('{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}conditionalFormattings'))
    root.append(extension)
    package.write(sheet.part_name, xml_bytes(root))
    before = state(sheet), package.read(sheet.part_name)
    with pytest.raises(VBAUnsupportedError, match='worksheet conditional extensions'):
        app.run('CopyNow')
    assert (state(sheet), package.read(sheet.part_name)) == before


@pytest.mark.parametrize('setup', ['Range("B1").AddComment "note"',
                                   'Range("B1:B2").FormulaArray="=ROW(B1:B2)"',
                                   'Range("B1:B2").Merge',
                                   'Range("B1").Validation.Add xlValidateWholeNumber,xlValidAlertStop,xlBetween,1,9'])
def test_affected_late_copy_layers_are_rejected_before_mutation(setup: str) -> None:
    app, sheet = prepared()
    app.add_module('Sub Extra()\n' + setup + '\nEnd Sub', name='ExtraModule')
    app.run('Extra')
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state(sheet) == before


def test_unrelated_notes_do_not_block_the_measured_copy_surface() -> None:
    app, sheet = prepared()
    app.add_module('Sub Extra()\nRange("Z20").AddComment "kept"\nEnd Sub', name='ExtraModule')
    app.run('Extra')
    note = sheet.notes[(20, 26)]
    app.run('CopyNow')
    assert sheet.notes[(20, 26)] is note
    assert sheet.conditional_formats[-1].areas == [Area(3, 6, 5, 8)]


def test_extension_query_is_cached_per_opc_entry_and_invalidated_by_part_replacement(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = prepared()
    app.save(tmp_path / 'before.xlsm')
    package = sheet.book.package
    assert package is not None
    calls: list[object] = []
    reader = transfer.xml_tree

    def counted(data: str | bytes) -> XmlElement:
        calls.append(data)
        return reader(data)

    monkeypatch.setattr(transfer, 'xml_tree', counted)
    for _ in range(20):
        assert plan(sheet, [Area(1, 10, 3, 12)], sheet, [Area(20, 20, 22, 22)]) is None
    assert len(calls) == 1 and sheet.conditional_extensions is False
    root = reader(package.read(sheet.part_name))
    root.append(XmlElement('{urn:extended}conditionalFormatting'))
    package.write(sheet.part_name, xml_bytes(root))
    with pytest.raises(VBAUnsupportedError, match='worksheet conditional extensions'):
        plan(sheet, [Area(1, 10, 3, 12)], sheet, [Area(20, 20, 22, 22)])
    assert len(calls) == 2 and sheet.conditional_extensions is True


def test_rejected_transfer_does_not_evaluate_source_udfs() -> None:
    app, sheet = prepared()
    app.add_module('''Public Calls As Long
Function Counted() As Long
Calls=Calls+1
Counted=5
End Function
Function ReadCalls() As Long
ReadCalls=Calls
End Function''', name='CountingModule')
    sheet.cells_[(1, 1)].formula = '=Counted()'
    sheet.cells_[(1, 1)].stale = True
    sheet.conditional_formats[0].root.set('unknown', 'unsupported')
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state(sheet) == before
    assert app.run('ReadCalls') == 0


@pytest.mark.parametrize('incoming', [True, False], ids=['copied-rule', 'destination-clear'])
@pytest.mark.parametrize('entry', ['copy', 'worksheet-paste', 'paste-all', 'paste-formats'])
def test_conditional_transfer_from_a_udf_refuses_without_mutation_or_events(
        incoming: bool, entry: str, monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = prepared()
    source = 'A1:C3' if incoming else 'J1:L3'
    action = (f'Range("{source}").Copy(Destination:=Range("F3:H5"))' if entry == 'copy' else
              'ActiveSheet.Paste(Destination:=Range("F3:H5"))' if entry == 'worksheet-paste' else
              'Range("F3:H5").PasteSpecial(' + ('xlPasteAll' if entry == 'paste-all' else 'xlPasteFormats') + ')')
    app.add_module(f'Sub ClipboardSetup()\nRange("{source}").Copy\nEnd Sub\n'
                   + 'Function Attempt() As Variant\nAttempt=' + action + '\nEnd Function', name='UdfTransferModule')
    app.run('ClipboardSetup')
    events: list[object] = []

    def after_edit(target: object, **kwargs: object) -> None:
        events.append(target)

    monkeypatch.setattr(_events, 'after_edit', after_edit)
    before = state(sheet)
    caller = Range(sheet, [Area(20, 13, 20, 13)])
    with running(app.application, caller):
        assert app.run('Attempt') is False
    assert state(sheet) == before and events == []
    assert app.application.udf_callers == []


def test_formula_invoked_copy_cannot_modify_conditional_destination() -> None:
    app, sheet = prepared()
    app.add_module('''Function WritesViaCopy() As Variant
WritesViaCopy=Range("A1:C3").Copy(Destination:=Range("F3:H5"))
End Function
Function EvaluateWriter() As Variant
Range("M20").Formula="=WritesViaCopy()"
EvaluateWriter=Range("M20").Value2
End Function''', name='FormulaWriterModule')
    before = state(sheet)
    assert app.run('EvaluateWriter') is False
    sheet.cells_.pop((20, 13))
    assert state(sheet) == before
    assert app.application.udf_callers == []


@pytest.mark.parametrize('conditional', [True, False], ids=['with-conditional', 'ordinary-cells'])
def test_public_python_copy_cannot_overwrite_locked_cells_when_formatting_is_allowed(conditional: bool) -> None:
    app, sheet = prepared()
    if not conditional:
        sheet.conditional_formats.clear()
    app.add_module('Sub LockSheet()\nActiveSheet.Protect AllowFormattingCells:=True\nEnd Sub', name='ProtectModule')
    app.run('LockSheet')
    before = state(sheet)
    view = app.sheet()
    with pytest.raises(VBARuntimeError) as failure:
        view.copy_range('A1:C3', view, 'F3:H5')
    assert failure.value.number == 1004
    assert state(sheet) == before


@pytest.mark.parametrize('attribute,value', [('dxfId', '-1'), ('dxfId', '999999'), ('dxfId', 'not-an-index'),
                                          ('priority', '0'), ('priority', '-1')])
@pytest.mark.parametrize('cached', [True, False])
def test_invalid_imported_style_identifiers_and_priorities_fail_before_mutation(
        attribute: str, value: str, cached: bool) -> None:
    app, sheet = prepared()
    rule = sheet.conditional_formats[0]
    rule.root.set(attribute, value)
    if not cached:
        rule.dxf = None
    before = state(sheet)
    with pytest.raises(VBAUnsupportedError):
        app.run('CopyNow')
    assert state(sheet) == before


@pytest.mark.parametrize('case', _UDF_CASES, ids=lambda case: case['name'])
def test_native_formula_invoked_transfers_leave_destination_unchanged(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeTransfer')
    assert app.run('Probe') == case['result']


@pytest.mark.parametrize('case', _UDF_RETURN_CASES, ids=lambda case: case['name'])
def test_native_transfer_returns_and_errors_in_formula_and_macro_contexts(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='NativeTransfer')
    assert app.run('Setup') == case['worksheet_result']
    assert app.run('MacroAttempt') == case['macro_result']


def test_mc_wrapped_main_namespace_rules_block_copy_before_mutation(
        tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = prepared()
    app.save(tmp_path / 'before-mc.xlsm')
    package = sheet.book.package
    assert package is not None
    root = xml_tree(package.read(sheet.part_name))
    mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
    x14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
    alternate = XmlElement('{' + mc + '}AlternateContent', {'xmlns:mc': mc, 'xmlns:x14': x14})
    choice = XmlElement('{' + mc + '}Choice', {'Requires': 'x14'})
    group = XmlElement(NS + 'conditionalFormatting', {'sqref': 'F3:H5'})
    rule = XmlElement(NS + 'cfRule', {'type': 'expression', 'priority': '1'})
    formula = XmlElement(NS + 'formula')
    formula.text = 'TRUE'
    rule.append(formula)
    group.append(rule)
    choice.append(group)
    alternate.append(choice)
    root.append(alternate)
    package.write(sheet.part_name, xml_bytes(root))
    events: list[object] = []

    def after_edit(target: object, **kwargs: object) -> None:
        events.append(target)

    monkeypatch.setattr(_events, 'after_edit', after_edit)
    before = state(sheet)
    before_parts = {part: package.read(part) for part in package.names()}
    with pytest.raises(VBAUnsupportedError, match='worksheet conditional extensions'):
        app.run('CopyNow')
    assert state(sheet) == before
    assert events == []
    assert {part: package.read(part) for part in package.names()} == before_parts
