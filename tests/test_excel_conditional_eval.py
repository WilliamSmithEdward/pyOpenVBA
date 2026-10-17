"""Ordinary rule evaluation, property composition, and live calculation context."""
from __future__ import annotations

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _styles as styles
from pyopenvba.apps.excel._charts_file import xml_tree
from pyopenvba.apps.excel._conditional import NS, Rule
from pyopenvba.apps.excel._conditional_eval import conditional_number_format, display_style
from pyopenvba.apps.excel._model import Worksheet
from pyopenvba.exceptions import VBAUnsupportedError


def sheet_with(source: str = '') -> tuple[ExcelApplication, Worksheet]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\n' + source + '\nEnd Sub', name='SetupModule')
    app.run('Setup')
    return app, book.sheets_[0]


def add(sheet: Worksheet, formula: str, patch: str, *, priority: int = 1,
        stop: bool = False, areas: list[Area] | None = None, kind: str = 'expression',
        operator: str = '') -> Rule:
    root = xml_tree(f'<cfRule xmlns="{NS[1:-1]}" type="{kind}" priority="{priority}" '
                    f'stopIfTrue="{int(stop)}" operator="{operator}">{formula}</cfRule>')
    rule = Rule(areas or [Area(1, 1, 10, 1)], root,
                dxf=xml_tree(f'<dxf xmlns="{NS[1:-1]}">{patch}</dxf>'))
    sheet.conditional_formats.append(rule)
    return rule


def test_relative_formula_and_caller_row_in_separate_areas() -> None:
    _, sheet = sheet_with('Range("C4").Value=10\nRange("F9").Value=20')
    add(sheet, '<formula>AND(C4&gt;5,ROW()=9,COLUMN()=6)</formula>', '<font><b/></font>',
        areas=[Area(4, 3, 5, 3), Area(9, 6, 10, 6)])
    assert not display_style(sheet, 4, 3).font.bold
    assert display_style(sheet, 9, 6).font.bold
    assert not display_style(sheet, 10, 6).font.bold


def test_priority_composes_properties_and_explicit_false() -> None:
    _, sheet = sheet_with('Range("A1").Font.Bold=True\nRange("A1").Font.Name="Courier New"')
    add(sheet, '<formula>TRUE</formula>', '<font><b val="0"/></font>', priority=1)
    add(sheet, '<formula>TRUE</formula>', '<font><b/><i/></font><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill>', priority=2)
    result = display_style(sheet, 1, 1)
    assert not result.font.bold and result.font.italic
    assert result.font.name == 'Courier New'
    assert result.fill.pattern == 'solid' and result.fill.foreground == styles.Color('rgb', 'FFFF0000')
    assert sheet.style_at(1, 1).font.bold and not sheet.style_at(1, 1).font.italic


def test_stop_if_true_without_dxf_still_stops() -> None:
    _, sheet = sheet_with()
    add(sheet, '<formula>FALSE</formula>', '<font><b/></font>', stop=True)
    add(sheet, '<formula>TRUE</formula>', '', priority=2, stop=True)
    add(sheet, '<formula>TRUE</formula>', '<font><i/></font>', priority=3)
    assert not display_style(sheet, 1, 1).font.italic


def test_cell_changes_recompute_without_style_or_rule_changes() -> None:
    app, sheet = sheet_with('Range("A1").Value=1')
    rule = add(sheet, '<formula>A1&gt;5</formula>', '<font><b/></font>')
    assert not display_style(sheet, 1, 1).font.bold
    app.add_module('Sub Change()\nRange("A1").Value=10\nEnd Sub', name='ChangeModule')
    app.run('Change')
    assert display_style(sheet, 1, 1).font.bold
    assert not sheet.style_at(1, 1).font.bold and rule.root.get('dxfId') is None


@pytest.mark.parametrize(('source', 'formula', 'operator', 'expected'), [
    ('', '0', 'equal', True),
    ('Range("A1").Value=True', '1', 'equal', False),
    ('Range("A1").Value="abc"', '"ABC"', 'equal', True),
    ('Range("A1").Value=CVErr(xlErrNA)', '0', 'notEqual', False),
    ('Range("A1").Value=5', '3', 'greaterThan', True),
])
def test_cell_comparisons_use_calculation_type_rules(source: str, formula: str,
                                                    operator: str, expected: bool) -> None:
    _, sheet = sheet_with(source)
    add(sheet, '<formula>' + formula + '</formula>', '<font><b/></font>', kind='cellIs', operator=operator)
    assert display_style(sheet, 1, 1).font.bold is expected


def test_error_conditions_do_not_stop_lower_rules() -> None:
    _, sheet = sheet_with()
    add(sheet, '<formula>1/0</formula>', '<font><b/></font>', stop=True)
    add(sheet, '<formula>TRUE</formula>', '<font><i/></font>', priority=2)
    assert display_style(sheet, 1, 1).font.italic


def test_long_formula_uses_target_context_without_evaluate_length_limit() -> None:
    _, sheet = sheet_with()
    formula = 'AND(' + ','.join(['TRUE'] * 70) + ',ROW()=6)'
    assert len(formula) > 255
    add(sheet, '<formula>' + formula + '</formula>', '<font><b/></font>')
    assert display_style(sheet, 6, 1).font.bold
    assert not display_style(sheet, 5, 1).font.bold


def test_absolute_references_remain_fixed_and_relative_references_move() -> None:
    _, sheet = sheet_with('Range("A1").Value=10\nRange("B4").Value=10\nRange("B5").Value=20')
    add(sheet, '<formula>$A$1=B4</formula>', '<font><b/></font>', areas=[Area(4, 2, 5, 2)])
    assert display_style(sheet, 4, 2).font.bold
    assert not display_style(sheet, 5, 2).font.bold


def test_partial_border_and_inline_number_format() -> None:
    _, sheet = sheet_with('Range("A1").Borders(xlEdgeTop).LineStyle=xlContinuous')
    add(sheet, '<formula>TRUE</formula>', '<border><bottom style="thin"><color rgb="FFFF0000"/></bottom></border><numFmt numFmtId="14" formatCode="0.00%"/>')
    result = display_style(sheet, 1, 1)
    assert result.border.top == sheet.style_at(1, 1).border.top
    assert result.border.bottom == styles.Side('thin', styles.Color('rgb', 'FFFF0000'))
    assert result.number_format == '0.00%'


@pytest.mark.parametrize(('operator', 'expected'), [('between', True), ('notBetween', False)])
def test_native_between_accepts_reversed_bounds(operator: str, expected: bool) -> None:
    _, sheet = sheet_with('Range("A1").Value=5')
    add(sheet, '<formula>10</formula><formula>1</formula>', '<font><b/></font>',
        kind='cellIs', operator=operator)
    assert display_style(sheet, 1, 1).font.bold is expected


def test_native_color_only_border_paints_continuous_over_lower_dash() -> None:
    _, sheet = sheet_with()
    add(sheet, '<formula>TRUE</formula>', '<border><bottom><color rgb="FFFF0000"/></bottom></border>')
    add(sheet, '<formula>TRUE</formula>', '<border><bottom style="dashed"><color rgb="FF0000FF"/></bottom></border>', priority=2)
    result = display_style(sheet, 1, 1)
    assert result.border.bottom == styles.Side('thin', styles.Color('rgb', 'FFFF0000'))


def test_unknown_rule_is_reported_only_when_reached() -> None:
    _, sheet = sheet_with()
    first = add(sheet, '<formula>TRUE</formula>', '', stop=True)
    add(sheet, '', '', priority=2, kind='dataBar')
    assert display_style(sheet, 1, 1) == sheet.style_at(1, 1)
    first.deleted = True
    with pytest.raises(VBAUnsupportedError, match='dataBar'):
        display_style(sheet, 1, 1)


def test_number_format_ignores_advanced_visual_rules_without_formats() -> None:
    _, sheet = sheet_with('Range("A1").NumberFormat="0.000"')
    add(sheet, '', '', kind='dataBar', stop=True)
    assert conditional_number_format(sheet, 1, 1) == '0.000'


def test_number_format_does_not_render_gradient_or_advanced_colors() -> None:
    _, sheet = sheet_with()
    add(sheet, '<formula>TRUE</formula>', '<fill><gradientFill/></fill>')
    add(sheet, '', '', priority=2, kind='colorScale')
    add(sheet, '<formula>TRUE</formula>', '<numFmt numFmtId="164" formatCode="0.00"/>', priority=3)
    assert conditional_number_format(sheet, 1, 1) == '0.00'


def test_number_format_honors_stop_even_when_rule_has_no_format() -> None:
    _, sheet = sheet_with()
    first = add(sheet, '<formula>TRUE</formula>', '<font><b/></font>', stop=True)
    add(sheet, '<formula>TRUE</formula>', '<numFmt numFmtId="164" formatCode="0.00"/>', priority=2)
    assert conditional_number_format(sheet, 1, 1) == 'General'
    formula = first.root.find(NS + 'formula')
    assert formula is not None
    formula.text = 'FALSE'
    assert conditional_number_format(sheet, 1, 1) == '0.00'


@pytest.mark.parametrize('patch', ['', '<numFmt numFmtId="164" formatCode="0.0"/>'])
def test_number_format_reports_relevant_unknown_rule(patch: str) -> None:
    _, sheet = sheet_with()
    add(sheet, '', patch, kind='iconSet', stop=True)
    add(sheet, '<formula>TRUE</formula>', '<numFmt numFmtId="164" formatCode="0.00"/>', priority=2)
    with pytest.raises(VBAUnsupportedError, match='iconSet'):
        conditional_number_format(sheet, 1, 1)
