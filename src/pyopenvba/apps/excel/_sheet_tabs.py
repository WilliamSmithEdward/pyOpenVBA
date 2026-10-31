"""A worksheet's tab colour, Worksheet.Tab, as Excel sets it, reads it back and keeps it in the sheet's
sheetPr (tests/fixtures/sheet_tabs/).

A tab has no colour, an RGB colour, a ColorIndex's palette slot or a
theme colour, any of them with a tint. A save writes the colour as the
sheetPr's first child, tabColor, and leaves out a sheetPr nothing else
needs.
"""
from __future__ import annotations

import re
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba._opc_xml import opc_part
from pyopenvba.apps.excel import _styles as S
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_integer, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _long(value: int) -> VBAInt:
    return VBAInt(value, 'Long')


class Tab(ExcelObject):
    """A worksheet's tab, whose colour Color, ColorIndex, ThemeColor and TintAndShade set and read."""

    vba_type_name = 'Tab'

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet

    def _set(self, color: S.Color | None) -> None:
        self.sheet.tab_color = color
        self.sheet.tab_changed = True
        self.sheet.book.saved = False

    def _rrggbb(self, color: S.Color) -> str:
        # A tab's palette colour is Excel's own, whatever Workbook.Colors made of the slot (tests/fixtures/
        # workbook_colors.json).
        colors = self.sheet.book.stylesheet.colors
        return S.Colors(colors.theme).rrggbb(color, 'FFFFFF')

    @member
    def Parent(self) -> object:
        return self.sheet

    @member
    def Color(self) -> object:
        """The colour as VBA packs it, False for a tab with none."""
        color = self.sheet.tab_color
        return False if color is None else _long(S.bgr(self._rrggbb(color)))

    @setter('Color')
    def _set_color(self, value: object) -> None:
        """False takes the colour away; any other value is a colour, its low 24 bits, so xlNone is a pale blue
        and -1 white, as in Excel."""
        if value is False:
            self._set(None)
            return
        self._set(S.Color.from_bgr(int(to_integer(value, 'Long')) & 0xFFFFFF))

    @member
    def ColorIndex(self) -> object:
        """The colour's palette slot, as a tab finds the nearest for any other colour (see Colors.tab_index);
        xlColorIndexNone for none."""
        color = self.sheet.tab_color
        if color is None:
            return _long(S.XL_NONE)
        return _long(self.sheet.book.stylesheet.colors.tab_index(color))

    @setter('ColorIndex')
    def _set_color_index(self, value: object) -> None:
        """A palette slot from 1 to 56, xlColorIndexNone taking the colour away; anything else, xlAutomatic among
        them, error 9."""
        index = int(to_integer(value, 'Long'))
        if index == S.XL_NONE:
            self._set(None)
            return
        if not 1 <= index <= 56:
            raise error(9)
        self._set(S.Color('indexed', index + 7))

    @member
    def ThemeColor(self) -> object:
        """The theme colour's xlThemeColor, 0 for any other colour or none."""
        color = self.sheet.tab_color
        return _long(int(color.value) + 1 if color is not None and color.kind == 'theme' else 0)

    @setter('ThemeColor')
    def _set_theme_color(self, value: object) -> None:
        """xlThemeColorDark1 to xlThemeColorFollowedHyperlink, which the file numbers from 0; error 9 beyond."""
        index = int(to_integer(value, 'Long'))
        if not 1 <= index <= 12:
            raise error(9)
        self._set(S.Color('theme', index - 1))

    @member
    def TintAndShade(self) -> object:
        """The colour's tint, False for a tab with none."""
        color = self.sheet.tab_color
        return False if color is None else color.tint / S.TINT_SCALE

    @setter('TintAndShade')
    def _set_tint(self, value: object) -> None:
        """A tint of the colour, cut toward zero to n / 32767 and kept in 16 bits, so 1.5 comes back as
        -0.5000763; a tab with no colour takes none."""
        tint = S.tint_of(float(to_number(value)))
        color = self.sheet.tab_color
        if color is not None:
            self._set(replace(color, tint=(tint + 32768) % 65536 - 32768))


def read(sheet: Worksheet, xml: str) -> None:
    """The tab colour a sheet's XML keeps."""
    found = re.search(r'<(?:\w+:)?sheetPr\b[^>]*?(?:/>|>(.*?)</(?:\w+:)?sheetPr>)', xml, re.DOTALL)
    if found is None or not found.group(1):
        return
    tab = re.search(r'<(?:\w+:)?tabColor\b[^>]*>', found.group(1))
    if tab is not None:
        sheet.tab_color = S.parse_color(tab.group())


def with_tab(sheet: Worksheet, xml: str) -> str:
    """A sheet's XML with its tab colour as the model has it: the first child of sheetPr, or gone, and with it a
    sheetPr nothing else is left in."""
    part = opc_part(xml.encode(), NS, 'worksheet', container_path=('sheetPr',))
    if part is None:
        if sheet.tab_color is None:
            return xml
        # sheetPr is the worksheet's first child.
        root = opc_part(xml.encode(), NS, 'worksheet')
        prefix = root.qname.rpartition(':')[0]
        tag = f'{prefix}:sheetPr' if prefix else 'sheetPr'
        color = root.child('tabColor', S.color_attributes(sheet.tab_color))
        return root.insert(root.opening_end, f'<{tag}>'.encode() + color + f'</{tag}>'.encode()).decode('utf-8')
    data = part.data
    old = next(((start, end) for name, _a, start, end in part.children if name == NS + '}tabColor'), None)
    markup = b'' if sheet.tab_color is None else part.child('tabColor', S.color_attributes(sheet.tab_color))
    if old is not None:
        data = data[:old[0]] + markup + data[old[1]:]
    elif markup:
        data = part.insert(part.opening_end, markup)
    text = data.decode('utf-8')
    if sheet.tab_color is None:
        # A sheetPr with nothing left in it goes too, as Excel writes none.
        text = re.sub(r'<((?:\w+:)?sheetPr)\s*(?:/>|></\1>)', '', text, count=1)
    return text
