"""Lossless differential formats and native partial-style semantics."""
from __future__ import annotations

from collections.abc import Callable
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.apps.excel._differential import (
    NS,
    ConditionalBorder,
    ConditionalBorders,
    ConditionalFont,
    ConditionalInterior,
    number_format,
    set_number_format,
)
from pyopenvba.apps.excel._model import Worksheet
from pyopenvba.apps.excel._styles import Latent, Stylesheet
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL


def test_dxf_append_keeps_original_xml_unknown_extensions_and_ids() -> None:
    original = '<dxf strange="kept"><font><b val="0"/></font><extLst><x:custom token="x:item"/></extLst></dxf>'
    suffix = '<extLst><x:styles/></extLst>'
    text = ('<styleSheet xmlns="' + NS[1:-1] + '" xmlns:x="urn:custom">'
            '<dxfs count="1" unknown="yes">' + original + '</dxfs>' + suffix + '</styleSheet>')
    styles = Stylesheet(text)
    assert styles.dxf_count == 1
    assert styles.written() == text
    own = styles.dxf(0)
    assert styles.ensure_dxf(own) == 0
    bold = own.find(NS + 'font/' + NS + 'b')
    assert bold is not None
    bold.set('val', '1')
    assert styles.ensure_dxf(own) == 1
    assert styles.dxf(0).find(NS + 'font/' + NS + 'b').get('val') == '0'  # type: ignore[union-attr]
    styles.prepare([], Latent())
    written = styles.written()
    assert original in written and suffix in written
    assert '<dxfs count="2" unknown="yes">' in written
    reopened = Stylesheet(written)
    assert reopened.dxf_count == 2
    assert reopened.dxf(1).find(NS + 'extLst/{urn:custom}custom') is not None
    styles.saved()
    assert styles.written() == written
    assert not styles.dirty


def test_dxf_only_save_and_repeated_append_are_stable() -> None:
    styles = Stylesheet('<styleSheet xmlns="' + NS[1:-1] + '"><extLst/></styleSheet>')
    first = xml_tree('<dxf xmlns="' + NS[1:-1] + '"><font><i/></font></dxf>')
    assert styles.ensure_dxf(first) == 0
    styles.prepare([], Latent())
    styles.saved()
    assert styles.dxf_count == 1
    assert Stylesheet(styles.written()).dxf_count == 1
    assert styles.ensure_dxf(styles.dxf(0)) == 0
    second = xml_tree('<dxf xmlns="' + NS[1:-1] + '"><font><strike/></font></dxf>')
    assert styles.ensure_dxf(second) == 1
    styles.saved()
    assert Stylesheet(styles.written()).dxf_count == 2
    assert not styles.dirty


def test_append_never_targets_nested_extension_dxf_or_quoted_greater_than() -> None:
    text = ('<styleSheet xmlns="' + NS[1:-1] + '" xmlns:x="urn:unknown">'
            '<dxfs note="count=\'99\'" count="1"><dxf unknown="a&gt;b"/><extLst><x:dxf text="kept"/></extLst></dxfs></styleSheet>')
    styles = Stylesheet(text)
    styles.ensure_dxf(xml_tree('<dxf xmlns="' + NS[1:-1] + '"><font><b/></font></dxf>'))
    written = styles.written()
    assert '<dxf unknown="a&gt;b"/>' in written
    assert '<extLst><x:dxf text="kept"/></extLst>' in written
    assert '<dxfs note="count=\'99\'" count="2">' in written
    reopened = Stylesheet(written)
    assert reopened.dxf_count == 2
    assert reopened.dxf(0).get('unknown') == 'a>b'


@pytest.mark.parametrize('index', [-1, 0, 9])
def test_dxf_invalid_index(index: int) -> None:
    with pytest.raises(IndexError):
        Stylesheet('<styleSheet/>').dxf(index)


@pytest.fixture
def patch() -> tuple[XmlElement, Callable[[], None], Worksheet, list[int]]:
    app = ExcelApplication()
    book = app.add_workbook()
    changes: list[int] = []
    return XmlElement(NS + 'dxf'), lambda: changes.append(1), book.sheets_[0], changes


def test_font_inherits_unset_flags_and_retains_explicit_false(patch: Any) -> None:
    root, changed, sheet, changes = patch
    font = ConditionalFont(root, changed, sheet)
    assert font.Bold() is NULL and font.Italic() is NULL and font.Strikethrough() is NULL
    font.vba_set('Bold', NULL)
    assert font.Bold() is False and font.Italic() is False
    assert font.Strikethrough() is NULL
    font.vba_set('Italic', True)
    font.vba_set('Bold', False)
    assert font.Bold() is False and font.Italic() is True
    assert root.find(NS + 'font/' + NS + 'b').get('val') == '0'
    assert len(changes) == 3


def test_font_color_none_restores_inheritance(patch: Any) -> None:
    root, changed, sheet, _ = patch
    font = ConditionalFont(root, changed, sheet)
    font.vba_set('Color', 255)
    assert font.Color() == 255 and font.ColorIndex() == 3
    font.vba_set('ColorIndex', -4142)
    assert font.Color() is NULL and font.ColorIndex() is NULL
    font.vba_set('ColorIndex', 0)
    assert font.Color() == 0 and font.ColorIndex() == -4105


def test_fill_changes_background_without_inventing_pattern_or_font(patch: Any) -> None:
    root, changed, sheet, _ = patch
    interior = ConditionalInterior(root, changed, sheet)
    assert interior.Color() == 0 and interior.ColorIndex() is NULL and interior.Pattern() is NULL
    interior.vba_set('Color', 255)
    assert root.find(NS + 'fill/' + NS + 'patternFill/' + NS + 'bgColor').get('rgb') == 'FFFF0000'
    assert interior.Color() == 255 and interior.Pattern() is NULL
    assert root.find(NS + 'font') is None
    interior.vba_set('PatternColor', 65280)
    pattern = root.find(NS + 'fill/' + NS + 'patternFill')
    assert [entry.tag for entry in pattern] == [NS + 'fgColor', NS + 'bgColor']
    assert interior.PatternColor() == 65280
    interior.vba_set('Pattern', -4142)
    assert interior.Color() == 16777215 and interior.ColorIndex() == -4142
    assert interior.PatternColor() is NULL


def test_inline_number_format_ignores_workbook_format_id_and_preserves_unknown(patch: Any) -> None:
    root, _, sheet, _ = patch
    entry = XmlElement(NS + 'numFmt', {'numFmtId': '14', 'formatCode': '0.00%', 'custom': 'kept'})
    root.append(entry)
    assert number_format(root, sheet.book.stylesheet) == '0.00%'
    set_number_format(root, '"_x0041_"0.00', sheet.book.stylesheet)
    assert entry.get('numFmtId') == '164' and entry.get('custom') == 'kept'
    assert number_format(xml_tree(xml_bytes(root)), sheet.book.stylesheet) == '"_x0041_"0.00'
    assert number_format(XmlElement(NS + 'dxf'), sheet.book.stylesheet) is EMPTY
    set_number_format(root, '0.00', sheet.book.stylesheet)
    assert entry.get('numFmtId') == '2'


def test_gradient_patch_stays_untouched_when_pattern_edit_is_unsupported(patch: Any) -> None:
    root, changed, sheet, changes = patch
    fill = XmlElement(NS + 'fill')
    fill.append(XmlElement(NS + 'gradientFill', {'degree': '90'}))
    root.append(fill)
    before = xml_bytes(root)
    interior = ConditionalInterior(root, changed, sheet)
    with pytest.raises(VBAUnsupportedError):
        interior.vba_set('Color', 255)
    assert xml_bytes(root) == before and changes == []


def test_custom_number_formats_share_ids_by_code_without_global_registration(patch: Any) -> None:
    root, _, sheet, _ = patch
    styles = sheet.book.stylesheet
    identifiers: list[str] = []
    for code in ('0.0000" widgets"', '0.000" items"', '0.0000" widgets"'):
        set_number_format(root, code, styles)
        entry = root.find(NS + 'numFmt')
        assert entry is not None
        identifiers.append(entry.get('numFmtId', ''))
        styles.ensure_dxf(root)
    assert identifiers == ['164', '165', '164']
    assert not any(identifier >= 164 for identifier in styles.number_formats)
    styles.number_formats[200] = '0.00000" global"'
    set_number_format(root, '0.000000" third"', styles)
    entry = root.find(NS + 'numFmt')
    assert entry is not None and entry.get('numFmtId') == '201'


@pytest.mark.parametrize('ui_only,allows', [(False, False), (True, False), (False, True)])
def test_conditional_font_protection_includes_ui_only(patch: Any, ui_only: bool, allows: bool) -> None:
    from pyopenvba.apps.excel._protection import SheetProtection

    root, changed, sheet, changes = patch
    sheet.protection = SheetProtection(user_interface_only=ui_only,
                                       allows=frozenset({'AllowFormattingCells'} if allows else ()))
    font = ConditionalFont(root, changed, sheet)
    if allows:
        font.vba_set('Bold', True)
        assert font.Bold() is True
    else:
        with pytest.raises(VBARuntimeError) as caught:
            font.vba_set('Bold', True)
        assert caught.value.number == 1004
        assert changes == [] and len(root) == 0


def test_conditional_borders_use_legacy_sides_and_write_partial_border(patch: Any) -> None:
    root, changed, sheet, _ = patch
    borders = ConditionalBorders(root, changed, sheet)
    bottom = borders.vba_lookup(-4107, [])
    assert isinstance(bottom, ConditionalBorder)
    assert bottom.LineStyle() is NULL and bottom.Weight() is NULL
    assert bottom.Color() == 0 and bottom.ColorIndex() == -4142
    bottom.vba_set('LineStyle', 1)
    bottom.vba_set('Color', 255)
    assert bottom.LineStyle() == 1 and bottom.Weight() == 2
    assert bottom.Color() == 255 and bottom.ColorIndex() == 3
    border = root.find(NS + 'border')
    assert [entry.tag for entry in border] == [NS + 'bottom']
    modern = borders.vba_lookup(9, [])
    assert isinstance(modern, ConditionalBorder)
    assert modern.Color() == 255
    with pytest.raises(VBARuntimeError) as caught:
        modern.LineStyle()
    assert caught.value.number == 1004
    before = xml_bytes(root)
    with pytest.raises(VBARuntimeError):
        bottom.vba_set('Weight', 4)
    assert xml_bytes(root) == before
