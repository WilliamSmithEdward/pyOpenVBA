"""Supported CF transfers persist; deferred moves fail before changing state."""
from __future__ import annotations

from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import xml_bytes
from pyopenvba.apps.excel._model import Workbook
from pyopenvba.exceptions import VBAUnsupportedError


def prepare(rules: str, action: str) -> tuple[ExcelApplication, Workbook]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\nRange("A1").Value=10\nRange("D1").Value=30\n' + rules
                   + '\nEnd Sub\nSub Action()\n' + action + '\nEnd Sub', name='Harness')
    app.run('Setup')
    return app, book


def state(book: Workbook) -> object:
    return [(sheet.name, {position: (cell.value, cell.formula, cell.style) for position, cell in sheet.cells_.items()},
             [(tuple(rule.areas), xml_bytes(rule.root), rule.deleted) for rule in sheet.conditional_formats])
            for sheet in book.sheets_]


RULE = 'Range("A1:A5").FormatConditions.Add Type:=xlExpression,Formula1:="=A1>0"'


@pytest.mark.parametrize('action', [
    'Rows(3).Insert', 'Rows(1).Delete', 'Columns(1).Insert', 'Columns(1).Delete',
    'Range("A3").Insert xlShiftDown', 'Range("A3").Delete xlShiftUp',
])
def test_structural_rule_movement_is_atomic(action: str) -> None:
    app, book = prepare(RULE, action)
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before


def test_structural_reference_on_another_sheet_is_guarded() -> None:
    setup = 'ActiveSheet.Name="Data"\nWorksheets.Add.Name="Report"\nRange("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=Data!A3>0"'
    app, book = prepare(setup, 'Worksheets("Data").Rows(2).Insert')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting formulas'):
        app.run('Action')
    assert state(book) == before


@pytest.mark.parametrize('action', ['Rows(9).Insert', 'Range("D3").Delete xlShiftUp'])
def test_unrelated_structural_edits_remain_available(action: str) -> None:
    app, book = prepare(RULE, action)
    rule = book.sheets_[0].conditional_formats[0]
    before = (tuple(rule.areas), xml_bytes(rule.root))
    app.run('Action')
    assert (tuple(rule.areas), xml_bytes(rule.root)) == before


@pytest.mark.parametrize('action', [
    'Range("A1").Cut Range("D1")',
    'Range("A1").Cut\nActiveSheet.Paste Range("D1")',
])
def test_cut_cannot_drop_source_rules(action: str) -> None:
    app, book = prepare(RULE, action)
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before


@pytest.mark.parametrize(('action', 'value', 'formula'), [
    ('Range("A1").Copy Range("D1")', '10', '=5+5'),
    ('Range("A1").Copy\nRange("D1").PasteSpecial xlPasteFormats', '30', '30'),
    ('Range("A1").Copy\nRange("D1").PasteSpecial xlPasteAll', '10', '=5+5'),
])
def test_format_transfer_clones_and_replaces_destination_coverage(
    action: str, value: str, formula: str, tmp_path: Path,
) -> None:
    rules = '''Dim f As Object,g As Object
Range("A1").Formula="=5+5"
Set f=Range("A1:A5").FormatConditions.Add(xlExpression,Formula1:="=A1>0")
f.Font.Bold=True
f.Interior.Color=vbRed
f.NumberFormat="0.00"
f.StopIfTrue=False
Set g=Range("D1:D3").FormatConditions.Add(xlCellValue,xlGreater,5)
g.Font.Italic=True
g.Interior.Color=vbGreen'''
    app, _ = prepare(rules, action)
    app.add_module('''Sub ChangeSource()
Range("A1").FormatConditions(1).Font.Bold=False
End Sub
Function InspectTransfer() As String
Dim f As Object,g As Object,h As Object
Set f=Range("A1").FormatConditions(1)
Set g=Range("D1").FormatConditions(1)
Set h=Range("D2").FormatConditions(1)
InspectTransfer=CStr(Range("D1").Value2) & "|" & CStr(Range("D1").Formula) & "|" & Cells.FormatConditions.Count
InspectTransfer=InspectTransfer & "|" & f.AppliesTo.Address & "|" & f.Priority & "|" & f.Formula1
InspectTransfer=InspectTransfer & "|" & g.AppliesTo.Address & "|" & g.Priority & "|" & g.Formula1
InspectTransfer=InspectTransfer & "|" & h.AppliesTo.Address & "|" & h.Priority
InspectTransfer=InspectTransfer & "|" & f.Font.Bold & "|" & g.Font.Bold & "|" & g.Interior.Color & "|" & g.NumberFormat
InspectTransfer=InspectTransfer & "|" & Range("D1").DisplayFormat.Font.Bold & "|" & Range("D1").Text
InspectTransfer=InspectTransfer & "|" & Range("D1").Font.Bold & "|" & Range("D1").NumberFormat
End Function''', name='Inspection')
    app.run('Action')
    app.run('ChangeSource')
    expected = (value + '|' + formula + '|3|$A$1:$A$5|2|=A1>0|$D$1|1|=D1>0|$D$2:$D$3|3'
                + '|False|True|255|0.00|True|' + value + '.00|False|General')
    assert app.run('InspectTransfer') == expected
    reopened = ExcelApplication.open(app.save(tmp_path / 'supported-transfer.xlsm'))
    assert reopened.run('InspectTransfer') == expected


def test_copy_without_source_rule_clears_only_landing_coverage(tmp_path: Path) -> None:
    app, _ = prepare('Range("A1").Formula="=5+5"\n'
                     'Range("D1:D3").FormatConditions.Add xlCellValue,xlGreater,5',
                     'Range("A1").Copy Range("D1")')
    app.add_module('''Function InspectCleared() As String
Dim f As Object
Set f=Range("D2").FormatConditions(1)
InspectCleared=Range("D1").Value2 & "|" & Range("D1").Formula & "|" & Range("D1").FormatConditions.Count
InspectCleared=InspectCleared & "|" & Cells.FormatConditions.Count & "|" & f.AppliesTo.Address & "|" & f.Priority & "|" & f.Formula1
End Function''', name='Inspection')
    app.run('Action')
    expected = '10|=5+5|0|1|$D$2:$D$3|1|=5'
    assert app.run('InspectCleared') == expected
    reopened = ExcelApplication.open(app.save(tmp_path / 'clear-landing.xlsm'))
    assert reopened.run('InspectCleared') == expected


def test_value_paste_keeps_conditional_rules() -> None:
    app, book = prepare(RULE, 'Range("D1").Copy\nRange("A1").PasteSpecial xlPasteValues')
    rule = book.sheets_[0].conditional_formats[0]
    before = (tuple(rule.areas), xml_bytes(rule.root))
    app.run('Action')
    assert book.sheets_[0].cells_[(1, 1)].value == 30
    assert (tuple(rule.areas), xml_bytes(rule.root)) == before


def test_unrelated_copy_on_a_sheet_with_rules_remains_available() -> None:
    app, book = prepare(RULE, 'Range("D1").Copy Range("F1")')
    app.run('Action')
    assert book.sheets_[0].cells_[(1, 6)].value == 30


def test_cut_guards_rule_references_on_another_sheet() -> None:
    setup = 'ActiveSheet.Name="Data"\nWorksheets.Add.Name="Report"\nRange("A1").FormatConditions.Add Type:=xlExpression,Formula1:="=Data!A1>0"'
    app, book = prepare(setup, 'Worksheets("Data").Range("A1").Cut Worksheets("Data").Range("D1")')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting formulas'):
        app.run('Action')
    assert state(book) == before


def test_multi_area_copy_checks_every_area_before_pasting_first() -> None:
    app, book = prepare('Range("D3").FormatConditions.Add xlCellValue,xlGreater,5',
                        'Union(Range("D1"),Range("D3")).Copy Range("F1")')
    before = state(book)
    with pytest.raises(VBAUnsupportedError, match='conditional formatting'):
        app.run('Action')
    assert state(book) == before
