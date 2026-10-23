"""DisplayFormat composes rules while keeping the underlying cell style intact."""
from __future__ import annotations

import importlib.util
import json
import os
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._conditional_display import DisplayFormat
from pyopenvba.apps.excel._model import Range, Worksheet
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import NULL

ROOT = Path(__file__).parent / 'fixtures/conditional'
CASES: list[dict[str, Any]] = [case for name in ('display.json', 'display_edges.json')
                              for case in json.loads((ROOT / name).read_text(encoding='utf-8'))['probes']]


def sheet_with(source: str) -> tuple[ExcelApplication, Worksheet]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\nDim f As FormatCondition,g As FormatCondition\n' + source + '\nEnd Sub', name='SetupModule')
    app.run('Setup')
    return app, book.sheets_[0]


def obj(value: object) -> VBAObject:
    assert isinstance(value, VBAObject)
    return value


@pytest.mark.parametrize('case', CASES, ids=lambda case: case['name'])
def test_independent_native_display_controls(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']


def test_mixed_font_fill_and_alignment_reads_do_not_change_cell_styles() -> None:
    _, sheet = sheet_with('''Range("A1:A3").Value=20
Range("A2").Value=5
Range("A1:A3").Font.Bold=True
Range("A1:A3").HorizontalAlignment=xlCenter
Set f=Range("A1:A3").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Font.Bold=False
f.Font.Italic=True
f.Interior.Color=RGB(255,0,0)''')
    before = {key: cell.style for key, cell in sheet.cells_.items()}
    display = DisplayFormat(Range(sheet, [Area(1, 1, 3, 1)]))
    font, interior = obj(display.Font()), obj(display.Interior())
    assert font.vba_get('Bold') is NULL and font.vba_get('Italic') is NULL
    assert interior.vba_get('Color') == 0 and interior.vba_get('ColorIndex') is NULL
    assert display.HorizontalAlignment() == -4108
    assert display.Locked() is True and display.FormulaHidden() is False
    assert {key: cell.style for key, cell in sheet.cells_.items()} == before


def test_retained_display_objects_see_new_rules_values_and_priority() -> None:
    app, sheet = sheet_with('Range("A1").Value=5')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    font = obj(display.Font())
    assert font.vba_get('Bold') is False
    app.add_module('''Sub Change()
Dim f As FormatCondition
Set f=Range("A1").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Font.Bold=True
Range("A1").Value=20
End Sub''', name='ChangeModule')
    app.run('Change')
    assert font.vba_get('Bold') is True
    assert not sheet.style_at(1, 1).font.bold


def test_conditional_number_format_affects_text_but_display_reports_base_format() -> None:
    app, sheet = sheet_with('''Range("A1").Value=20
Set f=Range("A1").FormatConditions.Add(xlCellValue,xlGreater,10)
f.NumberFormat="0.00"''')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    assert display.NumberFormat() == display.NumberFormatLocal() == 'General'
    app.add_module('Function Rendered() As String\nRendered=Range("A1").Text\nEnd Function', name='RenderedModule')
    assert app.run('Rendered') == '20.00'
    assert sheet.style_at(1, 1).number_format == 'General'


def test_display_border_items_and_collection_use_conditional_sides() -> None:
    _, sheet = sheet_with('''Range("A1:A2").Value=20
Set f=Range("A1:A2").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Borders(xlBottom).LineStyle=xlContinuous
f.Borders(xlBottom).Color=RGB(255,0,0)''')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 2, 1)]))
    bottom = obj(display.Borders(9))
    assert bottom.vba_get('LineStyle') == 1 and bottom.vba_get('Color') == 255
    assert obj(display.Borders(12)).vba_get('Color') == 255
    borders = obj(display.Borders())
    assert borders.vba_get('Count') == 6
    assert borders.vba_get('LineStyle') is NULL
    assert borders.vba_get('Color') == 0
    assert not sheet.style_at(1, 1).border.bottom.style


@pytest.mark.parametrize('nested,property_name,value,code', [
    ('Interior', 'Color', 255, 1004), ('Borders', 'Color', 255, 1004),
    ('Borders', 'LineStyle', 1, 1004), ('', 'NumberFormat', '0.00', 424),
])
def test_display_writes_refuse_without_changing_base_or_rule(
        nested: str, property_name: str, value: object, code: int) -> None:
    _, sheet = sheet_with('''Range("A1").Value=20
Set f=Range("A1").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Interior.Color=255''')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    target = obj(display.vba_get(nested)) if nested else display
    before = sheet.style_at(1, 1)
    with pytest.raises(VBARuntimeError) as raised:
        target.vba_set(property_name, value)
    assert raised.value.number == code
    assert sheet.style_at(1, 1) == before
    assert obj(display.Interior()).vba_get('Color') == 255


def test_sparse_whole_column_finds_conditional_exception_to_base_font() -> None:
    _, sheet = sheet_with('''Range("A500000").Value=20
Set f=Range("A500000").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Font.Bold=True''')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1048576, 1)]))
    assert obj(display.Font()).vba_get('Bold') is NULL
    assert len(sheet.cells_) == 1


@pytest.mark.parametrize('base,conditional,assigned,expected_display', [
    (False, None, True, True), (True, None, False, False),
    (False, False, True, False), (True, True, False, True),
])
def test_display_font_writes_base_while_matching_rule_still_controls_display(
        base: bool, conditional: bool | None, assigned: bool, expected_display: bool) -> None:
    source = f'Range("A1").Value=20\nRange("A1").Font.Bold={base}'
    if conditional is not None:
        source += '\nSet f=Range("A1").FormatConditions.Add(xlCellValue,xlGreater,10)'
        source += f'\nf.Font.Bold={conditional}'
    _, sheet = sheet_with(source)
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    font = obj(display.Font())
    font.vba_set('Bold', assigned)
    assert sheet.style_at(1, 1).font.bold is assigned
    assert font.vba_get('Bold') is expected_display


def test_unsupported_advanced_rule_does_not_silently_report_base_font() -> None:
    _, sheet = sheet_with('Range("A1").Value=20')
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._conditional import NS, Rule

    # A rule type no Excel writes stands for one the model cannot evaluate; a data bar draws and formats nothing.
    sheet.conditional_formats.append(Rule([Area(1, 1, 1, 1)],
                                         XmlElement(NS + 'cfRule', {'type': 'futureRule', 'priority': '1'})))
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    with pytest.raises(VBAUnsupportedError, match='futureRule'):
        obj(display.Font()).vba_get('Bold')


GENERATED_SOURCE = '''Sub MakeDisplay()
Dim f As FormatCondition
Range("A1").Value=20
Range("A2").Value=5
Range("A1:A2").Font.Bold=True
Set f=Range("A1:A2").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Font.Bold=False
f.Font.Italic=True
f.Interior.Color=RGB(255,0,0)
f.Borders(xlBottom).LineStyle=xlContinuous
f.Borders(xlBottom).Color=RGB(255,0,0)
f.NumberFormat="0.00"
End Sub
Function DisplaySnapshot() As String
Dim c As Range
For Each c In Range("A1:A2")
DisplaySnapshot=DisplaySnapshot & c.Address & ":" & CStr(c.Value2)
DisplaySnapshot=DisplaySnapshot & "|base=" & CStr(c.Font.Bold) & "," & CStr(c.Font.Italic) & "," & CStr(c.Interior.Color) & "," & c.NumberFormat
DisplaySnapshot=DisplaySnapshot & "," & CStr(c.Borders(xlEdgeBottom).LineStyle) & "," & CStr(c.Borders(xlEdgeBottom).Color)
DisplaySnapshot=DisplaySnapshot & "|shown=" & CStr(c.DisplayFormat.Font.Bold) & "," & CStr(c.DisplayFormat.Font.Italic) & "," & CStr(c.DisplayFormat.Interior.Color)
DisplaySnapshot=DisplaySnapshot & "," & CStr(c.DisplayFormat.Borders(xlEdgeBottom).LineStyle) & "," & CStr(c.DisplayFormat.Borders(xlEdgeBottom).Color)
DisplaySnapshot=DisplaySnapshot & "," & c.DisplayFormat.NumberFormat & "," & c.Text & ";"
Next
End Function
Function RecomputeDisplay() As String
Range("A1").Value=3
Range("A2").Value=40
Application.CalculateFull
RecomputeDisplay=DisplaySnapshot()
End Function
Function RuleSnapshot() As String
Dim f As FormatCondition
Set f=Range("A1").FormatConditions(1)
RuleSnapshot=CStr(Cells.FormatConditions.Count) & "|" & CStr(f.Priority) & "|" & f.AppliesTo.Address & "|" & f.Formula1 & "|" & f.NumberFormat
RuleSnapshot=RuleSnapshot & "|" & CStr(f.Font.Bold) & "," & CStr(f.Font.Italic) & "," & CStr(f.Interior.Color)
RuleSnapshot=RuleSnapshot & "," & CStr(f.Borders(xlBottom).LineStyle) & "," & CStr(f.Borders(xlBottom).Color)
End Function
Sub SaveNative()
ThisWorkbook.Save
End Sub'''
_BASE = 'True,False,16777215,General,-4142,0'
_MATCH = 'False,True,255,1,255,General,'
_NONMATCH = 'True,False,16777215,-4142,0,General,'
INITIAL_DISPLAY = (f'$A$1:20|base={_BASE}|shown={_MATCH}20.00;'
                   f'$A$2:5|base={_BASE}|shown={_NONMATCH}5;')
UPDATED_DISPLAY = (f'$A$1:3|base={_BASE}|shown={_NONMATCH}3;'
                   f'$A$2:40|base={_BASE}|shown={_MATCH}40.00;')
RULE_SETTINGS = '1|1|$A$1:$A$2|=10|0.00|False,True,255,1,255'


def generated_display_project(path: Path) -> ExcelApplication:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(GENERATED_SOURCE, name='Harness')
    app.run('MakeDisplay')
    before = {key: cell.style for key, cell in book.sheets_[0].cells_.items()}
    assert app.run('DisplaySnapshot') == INITIAL_DISPLAY
    assert app.run('RuleSnapshot') == RULE_SETTINGS
    assert {key: cell.style for key, cell in book.sheets_[0].cells_.items()} == before
    app.save(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('DisplaySnapshot') == INITIAL_DISPLAY
    assert reopened.run('RuleSnapshot') == RULE_SETTINGS
    return reopened


def test_generated_visible_styles_and_text_recompute_after_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / 'display.xlsm'
    app = generated_display_project(path)
    assert app.run('RecomputeDisplay') == UPDATED_DISPLAY
    assert app.run('RuleSnapshot') == RULE_SETTINGS
    app.save(path)
    reopened = ExcelApplication.open(path)
    assert reopened.run('DisplaySnapshot') == UPDATED_DISPLAY
    assert reopened.run('RuleSnapshot') == RULE_SETTINGS


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                   reason='requires explicitly enabled native Excel')
def test_generated_display_and_text_reopened_and_recomputed_by_native_excel(
        tmp_path: Path, native_excel: Any) -> None:
    path = tmp_path / 'native-display.xlsm'
    generated_display_project(path)
    native_excel.open_document(path)
    assert native_excel.run_raw('Harness.DisplaySnapshot') == INITIAL_DISPLAY
    assert native_excel.run_raw('Harness.RuleSnapshot') == RULE_SETTINGS
    assert native_excel.run_raw('Harness.RecomputeDisplay') == UPDATED_DISPLAY
    assert native_excel.run_raw('Harness.RuleSnapshot') == RULE_SETTINGS
    # open_document defaults to read-only. SaveAs explicitly writes the held,
    # modified workbook to a new path instead of silently retaining the source.
    native_path = tmp_path / 'native-saved-display.xlsm'
    native_excel.save_as(native_path)
    native_excel.new_document()
    native_excel.open_document(native_path)
    assert native_excel.run_raw('Harness.DisplaySnapshot') == UPDATED_DISPLAY
    assert native_excel.run_raw('Harness.RuleSnapshot') == RULE_SETTINGS
    reopened = ExcelApplication.open(native_path)
    assert reopened.run('DisplaySnapshot') == UPDATED_DISPLAY
    assert reopened.run('RuleSnapshot') == RULE_SETTINGS
    assert not any(event.get('dialogs') for event in native_excel.events)


def test_diagonal_display_sides_are_not_the_cells_right_edges() -> None:
    _, sheet = sheet_with('''Range("A1").Borders(xlDiagonalDown).LineStyle=xlDash
Range("A1").Borders(xlDiagonalDown).Color=255
Range("A1").Borders(xlEdgeRight).LineStyle=xlContinuous
Range("A1").Borders(xlEdgeRight).Color=16711680''')
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1, 1)]))
    down, up, right = obj(display.Borders(5)), obj(display.Borders(6)), obj(display.Borders(10))
    assert down.vba_get('LineStyle') == -4115 and down.vba_get('Color') == 255
    assert up.vba_get('LineStyle') == -4142
    assert right.vba_get('LineStyle') == 1 and right.vba_get('Color') == 16711680
    assert getattr(down, 'sides')()[0].style == 'dashed'
    assert not getattr(up, 'sides')()[0].style


def test_unaffected_whole_column_borders_use_sparse_base_reader(monkeypatch: pytest.MonkeyPatch) -> None:
    _, sheet = sheet_with('''Range("Z1").Value=20
Set f=Range("Z1").FormatConditions.Add(xlCellValue,xlGreater,10)
f.Font.Bold=True''')
    from pyopenvba.apps.excel import _conditional_display

    def unexpected_computation(*args: object) -> None:
        raise AssertionError('No conditional styles intersect this column or its neighbours')

    monkeypatch.setattr(getattr(_conditional_display, '_DisplaySheet'), 'style_at', unexpected_computation)
    display = DisplayFormat(Range(sheet, [Area(1, 1, 1048576, 1)]))
    assert obj(display.Borders()).vba_get('LineStyle') == -4142


@pytest.fixture(scope='module')
def native_excel() -> Iterator[Any]:
    harness = pytest.importorskip('pyvbaharness')
    with harness.ExcelSession() as excel:
        yield excel
