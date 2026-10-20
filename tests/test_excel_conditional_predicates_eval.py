"""Stored predicate formulas, typed values, caller anchoring and clock changes."""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._conditional import NS, Rule
from pyopenvba.apps.excel._conditional_eval import conditional_number_format, display_style
from pyopenvba.apps.excel._model import Worksheet
from pyopenvba.apps.excel._xstring import encode_text
from pyopenvba.formula._calc.cells import UnsupportedFormulaError
from pyopenvba.formula._calc.functions import text as text_functions


def setup(source: str = '') -> tuple[ExcelApplication, Worksheet]:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module('Sub Setup()\n' + source + '\nEnd Sub', name='SetupModule')
    app.run('Setup')
    return app, book.sheets_[0]


def predicate(sheet: Worksheet, kind: str, formula: str | None, *, areas: list[Area] | None = None,
              priority: int = 1, stop: bool = False) -> Rule:
    root = XmlElement(NS + 'cfRule', {'type': kind, 'priority': str(priority), 'stopIfTrue': str(int(stop))})
    if formula is not None:
        child = XmlElement(NS + 'formula')
        child.text = encode_text(formula)
        root.append(child)
    patch = XmlElement(NS + 'dxf')
    patch.append(XmlElement(NS + 'font'))
    patch[0].append(XmlElement(NS + 'b'))
    patch.append(XmlElement(NS + 'numFmt', {'numFmtId': '164', 'formatCode': '0.00'}))
    rule = Rule(areas or [Area(1, 1, 10, 1)], root, dxf=patch)
    sheet.conditional_formats.append(rule)
    return rule


@pytest.mark.parametrize(('kind', 'formula', 'expected'), [
    ('containsText', 'NOT(ISERROR(SEARCH("foo",A1)))', '1100'),
    ('notContainsText', 'ISERROR(SEARCH("foo",A1))', '0011'),
    ('beginsWith', 'LEFT(A1,LEN("foo"))="foo"', '1000'),
    ('endsWith', 'RIGHT(A1,LEN("foo"))="foo"', '0100'),
])
def test_text_predicates_evaluate_native_compatibility_formulas(kind: str, formula: str, expected: str) -> None:
    _, sheet = setup('Range("A1").Value="Foobar"\nRange("A2").Value="barFOO"\nRange("A3").Value="bar"\nRange("A4").Value=CVErr(2042)')
    predicate(sheet, kind, formula)
    actual = ''.join('1' if display_style(sheet, row, 1).font.bold else '0' for row in range(1, 5))
    assert actual == expected


@pytest.mark.parametrize(('kind', 'formula', 'expected'), [
    ('containsBlanks', 'LEN(TRIM(A1))=0', '1100010'),
    ('notContainsBlanks', 'LEN(TRIM(A1))>0', '0011101'),
    ('containsErrors', 'ISERROR(A1)', '0000001'),
    ('notContainsErrors', 'NOT(ISERROR(A1))', '1111110'),
])
def test_blank_error_typed_values_use_formula_semantics(kind: str, formula: str, expected: str) -> None:
    _, sheet = setup('Range("A2").Value=" "\nRange("A3").Value=Chr(9)\nRange("A4").Value=ChrW(160)\nRange("A5").Value=0\nRange("A6").Formula="="""""\nRange("A7").Value=CVErr(2042)')
    predicate(sheet, kind, formula)
    actual = ''.join('1' if display_style(sheet, row, 1).font.bold else '0' for row in range(1, 8))
    assert actual == expected


def test_predicate_formula_is_authoritative_and_missing_formula_is_inert() -> None:
    _, sheet = setup('Range("A1").Value="foo"')
    first = predicate(sheet, 'containsText', None)
    first.root.set('text', 'foo')
    assert not display_style(sheet, 1, 1).font.bold
    second = predicate(sheet, 'containsText', 'FALSE', priority=2)
    second.root.set('text', 'foo')
    assert conditional_number_format(sheet, 1, 1) == sheet.style_at(1, 1).number_format


def test_first_area_anchor_moves_formula_to_later_union_area() -> None:
    _, sheet = setup('Range("K3").Value="none"\nRange("M9").Value="foo"')
    predicate(sheet, 'containsText', 'NOT(ISERROR(SEARCH("foo",K3)))',
              areas=[Area(3, 11, 5, 11), Area(9, 13, 10, 13)])
    assert not display_style(sheet, 3, 11).font.bold
    assert display_style(sheet, 9, 13).font.bold
    assert conditional_number_format(sheet, 9, 13) == '0.00'


def test_time_period_uses_current_injected_day_without_stale_result(monkeypatch: pytest.MonkeyPatch) -> None:
    app, sheet = setup('Range("A1").Value=DateSerial(2025,1,1)+0.75')
    moment = [dt.datetime(2025, 1, 1, 23, 59)]
    monkeypatch.setattr(app.interpreter, 'clock', lambda: moment[0])
    predicate(sheet, 'timePeriod', 'FLOOR(A1,1)=TODAY()')
    assert display_style(sheet, 1, 1).font.bold
    assert conditional_number_format(sheet, 1, 1) == '0.00'
    moment[0] = dt.datetime(2025, 1, 2, 0, 1)
    assert not display_style(sheet, 1, 1).font.bold
    assert conditional_number_format(sheet, 1, 1) == sheet.style_at(1, 1).number_format


def test_predicate_number_format_roundtrip_does_not_require_rendering_other_features(tmp_path: Path) -> None:
    app, sheet = setup('Range("A1").Value="foo"')
    predicate(sheet, 'containsText', 'NOT(ISERROR(SEARCH("foo",A1)))')
    # Public DXF persistence is performed by normal FormatCondition edits; this
    # imported-style setup explicitly records the independent stored patch.
    rule = sheet.conditional_formats[0]
    assert rule.dxf is not None
    rule.root.set('dxfId', str(sheet.book.stylesheet.ensure_dxf(rule.dxf)))
    sheet.conditional_changed = True
    reopened = ExcelApplication.open(app.save(tmp_path / 'predicate.xlsm'))
    saved = reopened.workbooks()[0].sheets_[0]
    assert display_style(saved, 1, 1).font.bold
    assert conditional_number_format(saved, 1, 1) == '0.00'


def test_native_worksheet_codepoints_and_search_units_before_and_after_save(tmp_path: Path) -> None:
    app, _ = setup('Range("A1").Value=ChrW(-10179) & ChrW(-8704) & "x"\nRange("A2").Value="a" & ChrW(-10179) & ChrW(-8704) & "x"')
    app.add_module('''Function ReadUnicode() As String
ReadUnicode=CStr(Application.Evaluate("LEN(A1)")) & "|" & CStr(Len(Application.Evaluate("LEFT(A1,1)"))) & "|" & CStr(Len(Application.Evaluate("RIGHT(A1,2)"))) & "|" & CStr(Application.Evaluate("MID(A1,2,1)")) & "|" & CStr(Application.Evaluate("SEARCH(""x"",A1)")) & "|" & CStr(Application.Evaluate("SEARCH(""??"",A1)"))
ReadUnicode=ReadUnicode & "|" & CStr(Application.Evaluate("SEARCH(""?"",A1,2)")) & "|" & CStr(Application.Evaluate("ISERROR(SEARCH(""??"",A1,2))")) & "|" & CStr(Application.Evaluate("SEARCH(""??"",A2,2)"))
End Function''', name='Readback')
    assert app.run('ReadUnicode') == '2|2|3|x|2|1|2|True|2'
    reopened = ExcelApplication.open(app.save(tmp_path / 'unicode-functions.xlsm'))
    assert reopened.run('ReadUnicode') == '2|2|3|x|2|1|2|True|2'


def test_search_adversarial_nonmatch_is_iterative_and_keeps_glob_behavior() -> None:
    app, _ = setup('Range("A1").Value=String(32700,"a")')
    pattern = '*a' * 40 + 'b'
    app.add_module('''Function SearchLong() As String
SearchLong=CStr(Application.Evaluate("ISERROR(SEARCH(""''' + pattern + '''"",A1))")) & "|" & CStr(Application.Evaluate("SEARCH(""*a*a"",A1)"))
End Function''', name='Adversarial')
    assert app.run('SearchLong') == 'True|1'


def test_search_large_wildcard_work_is_explicitly_unsupported() -> None:
    app, sheet = setup('Range("A1").Value=String(32700,"a")')
    pattern = '*a' * 200 + 'b'
    predicate(sheet, 'containsText', 'NOT(ISERROR(SEARCH("' + pattern + '",A1)))')
    with pytest.raises(UnsupportedFormulaError, match='8,000,000 evaluation states'):
        display_style(sheet, 1, 1)
    with pytest.raises(UnsupportedFormulaError, match='8,000,000 evaluation states'):
        conditional_number_format(sheet, 1, 1)
    app.add_module('Function SearchLarge() As Variant\nSearchLarge=Application.Evaluate("SEARCH(REPT(""*a"",200)&""b"",A1)")\nEnd Function',
                   name='LargeSearch')
    with pytest.raises(UnsupportedFormulaError, match='8,000,000 evaluation states'):
        app.run('SearchLarge')


def test_search_literal_fast_path_and_collapsed_stars_do_not_use_quadratic_budget() -> None:
    app, _ = setup('Range("A1").Value=String(32767,"a")')
    app.add_module('Function SearchLiteral() As Variant\nSearchLiteral=Application.Evaluate("SEARCH(A1,A1)")\nEnd Function',
                   name='LiteralSearch')
    assert app.run('SearchLiteral') == 1
    position = getattr(text_functions, '_search_position')
    assert position('*' * 32767, 'a' * 32767, 0) == 0
