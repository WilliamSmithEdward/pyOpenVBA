"""Partial conditional formats, adapted from pyOfficeEditor's differential model.

Absent flags inherit, fills use bgColor, number formats are inline, and a
border edit touches only its selected side. Unknown XML remains on the patch.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _styles as styles
from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes
from pyopenvba.apps.excel._formats import LINES, PATTERNS, UNDERLINES, uniform, with_line_style, with_weight
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._number_format import from_file, normalized
from pyopenvba.apps.excel._xstring import decode
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, setter
from pyopenvba.interpreter._values import EMPTY, NULL, VBAInt, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
ORDER = ('font', 'numFmt', 'fill', 'alignment', 'border', 'protection', 'extLst')
#: CT_Font's children in the schema's order, which Excel writes a dxf's font in (tests/fixtures/conditional_joins/).
FONT_ORDER = ('b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name',
              'family', 'charset', 'scheme')
SIDES = {7: 'left', 8: 'top', 9: 'bottom', 10: 'right'}
LEGACY_SIDES = {1: 7, 2: 10, 3: 8, 4: 9, -4131: 7, -4152: 10, -4160: 8, -4107: 9}


def child(root: XmlElement, name: str) -> XmlElement:
    found = root.find(NS + name)
    if found is None:
        found = XmlElement(NS + name)
        order = {'dxf': ORDER, 'patternFill': ('fgColor', 'bgColor'), 'font': FONT_ORDER,
                 'border': ('left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal')}.get(
                     root.tag.rsplit('}', 1)[-1], ())
        rank = order.index(name) if name in order else len(order)
        position = next((index for index, entry in enumerate(root)
                         if entry.tag[len(NS):] in order and order.index(entry.tag[len(NS):]) > rank), len(root))
        root.insert(position, found)
    return found


def number_format(root: XmlElement) -> object:
    """A rule's number format as FormatCondition.NumberFormat reads it: the file's code, as Excel reads one."""
    entry = root.find(NS + 'numFmt')
    return EMPTY if entry is None else from_file(decode(entry.get('formatCode', '')))


def set_number_format(root: XmlElement, value: object, stylesheet: styles.Stylesheet) -> None:
    """Give a rule a number format: the code as NumberFormat keeps it, numbered and spelled as Excel writes a rule's
    (tests/fixtures/conditional_number_formats/)."""
    stylesheet.spell_rule_format(child(root, 'numFmt'), normalized(to_text(value)))


def renumber(root: XmlElement, stylesheet: styles.Stylesheet) -> None:
    """Number a rule's format as ``stylesheet``'s workbook does, for a rule a sheet copy brings from another."""
    entry = root.find(NS + 'numFmt')
    if entry is not None:
        stylesheet.spell_rule_format(entry, from_file(decode(entry.get('formatCode', ''))))


def read_color(entry: XmlElement | None) -> styles.Color | None:
    return None if entry is None else styles.parse_color(xml_bytes(entry).decode('utf-8'))


def write_color(root: XmlElement, name: str, color: styles.Color) -> None:
    entry = child(root, name)
    for attribute in ('rgb', 'indexed', 'auto', 'theme', 'tint'):
        entry.attrib.pop(attribute, None)
    entry.set(color.kind, str(color.value) if color.kind != 'auto' else '1')
    if color.tint:
        entry.set('tint', styles.excel_double(color.tint / styles.TINT_SCALE))


def palette(value: object) -> styles.Color:
    index = int(to_integer(value, 'Long'))
    if index == styles.XL_AUTOMATIC:
        return styles.Color('indexed', styles.SYSTEM_FOREGROUND)
    if not 1 <= index <= 56:
        raise error(9)
    return styles.Color('indexed', index + 7)


class _Patch(ExcelObject):
    def __init__(self, root: XmlElement, changed: Callable[[], None], sheet: Worksheet,
                 *, color_source: Callable[[], XmlElement] | None = None) -> None:
        self.root, self.changed, self.sheet = root, changed, sheet
        self.color_source = color_source

    def guard_set(self, member: str) -> None:
        protection = self.sheet.protection
        if protection is not None and protection.contents and 'AllowFormattingCells' not in protection.allows:
            raise error(1004)

    def color_answer(self, color: styles.Color, index: bool = False, automatic: str = '000000') -> object:
        colors = self.sheet.book.stylesheet.colors
        return VBAInt(colors.color_index(color, automatic), 'Long') if index else float(styles.bgr(colors.rrggbb(color, automatic)))

    @member
    def Application(self) -> object:
        return self.sheet.book.application


class ConditionalFont(_Patch):
    vba_type_name = 'Font'

    def flag(self, name: str) -> object:
        entry = self.root.find(NS + 'font/' + NS + name)
        return NULL if entry is None else entry.get('val', '1') in ('1', 'true')

    def set_flag(self, name: str, value: object) -> None:
        selected = False if value is NULL else to_bool(value)
        font = child(self.root, 'font')
        # Excel writes a flag that is on bare, <b/>, and one that is off with its value.
        flag = child(font, name)
        if selected:
            flag.attrib.pop('val', None)
        else:
            flag.set('val', '0')
        if name in ('b', 'i'):
            other = 'i' if name == 'b' else 'b'
            if font.find(NS + other) is None:
                child(font, other).set('val', '0')
        self.changed()

    @member
    def Bold(self) -> object:
        return self.flag('b')

    @setter('Bold')
    def _set_bold(self, value: object) -> None:
        self.set_flag('b', value)

    @member
    def Italic(self) -> object:
        return self.flag('i')

    @setter('Italic')
    def _set_italic(self, value: object) -> None:
        self.set_flag('i', value)

    @member
    def Strikethrough(self) -> object:
        return self.flag('strike')

    @setter('Strikethrough')
    def _set_strike(self, value: object) -> None:
        self.set_flag('strike', value)

    @member
    def Name(self) -> object:
        return NULL

    @member
    def Size(self) -> object:
        return NULL

    @member
    def Underline(self) -> object:
        entry = self.root.find(NS + 'font/' + NS + 'u')
        return NULL if entry is None else VBAInt(UNDERLINES.get(entry.get('val', 'single'), -4142), 'Long')

    @setter('Underline')
    def _set_underline(self, value: object) -> None:
        values = {-4142: 'none', 0: 'none', 2: 'single', -1: 'single', 3: 'double', -4119: 'double',
                  4: 'singleAccounting', 5: 'doubleAccounting'}
        chosen = int(to_integer(value, 'Long'))
        if chosen not in values:
            raise error(1004)
        child(child(self.root, 'font'), 'u').set('val', values[chosen])
        self.changed()

    def font_color(self, index: bool) -> object:
        color = read_color(self.root.find(NS + 'font/' + NS + 'color'))
        return NULL if color is None else self.color_answer(color, index)

    @member
    def Color(self) -> object:
        return self.font_color(False)

    @setter('Color')
    def _set_color(self, value: object) -> None:
        color = styles.Color.from_bgr(int(to_integer(value, 'Long')))
        write_color(child(self.root, 'font'), 'color', color)
        self.changed()

    @member
    def ColorIndex(self) -> object:
        return self.font_color(True)

    @setter('ColorIndex')
    def _set_color_index(self, value: object) -> None:
        selected = int(to_integer(value, 'Long'))
        if selected == styles.XL_NONE:
            font = self.root.find(NS + 'font')
            entry = None if font is None else font.find(NS + 'color')
            if entry is not None and font is not None:
                font.remove(entry)
            self.changed()
            return
        if selected == 0:
            value = styles.XL_AUTOMATIC
        color = palette(value)
        write_color(child(self.root, 'font'), 'color', color)
        self.changed()


class ConditionalInterior(_Patch):
    vba_type_name = 'Interior'

    def pattern(self, create: bool = False) -> XmlElement | None:
        fill = self.root.find(NS + 'fill')
        if fill is not None and fill.find(NS + 'gradientFill') is not None:
            raise VBAUnsupportedError('Conditional gradient fill properties are not implemented')
        if create:
            return child(child(self.root, 'fill'), 'patternFill')
        return None if fill is None else fill.find(NS + 'patternFill')

    def fill_color(self, index: bool, *, foreground: bool = False) -> object:
        pattern = self.pattern()
        color = read_color(None if pattern is None else pattern.find(NS + ('fgColor' if foreground else 'bgColor')))
        if pattern is not None and pattern.get('patternType') == 'none':
            return VBAInt(styles.XL_NONE, 'Long') if index else NULL if foreground else float(0xFFFFFF)
        if color is None:
            return NULL if index or foreground else 0.0
        result = self.color_answer(color, index, 'FFFFFF' if not foreground else '000000')
        return VBAInt(int(result), 'Long') if foreground and not index and isinstance(result, (int, float)) else result

    def set_fill_color(self, color: styles.Color, *, foreground: bool = False) -> None:
        pattern = self.pattern(True)
        assert pattern is not None
        write_color(pattern, 'fgColor' if foreground else 'bgColor', color)
        self.changed()

    @member
    def Color(self) -> object:
        return self.fill_color(False)

    @setter('Color')
    def _set_color(self, value: object) -> None:
        self.set_fill_color(styles.Color.from_bgr(int(to_integer(value, 'Long'))))

    @member
    def ColorIndex(self) -> object:
        return self.fill_color(True)

    @setter('ColorIndex')
    def _set_color_index(self, value: object) -> None:
        if int(to_integer(value, 'Long')) == styles.XL_NONE:
            self._set_pattern(styles.XL_NONE)
            return
        self.set_fill_color(palette(value))

    @member
    def Pattern(self) -> object:
        pattern = self.pattern()
        chosen = None if pattern is None else pattern.get('patternType')
        return NULL if chosen is None else VBAInt(PATTERNS.get(chosen, styles.XL_NONE), 'Long')

    @setter('Pattern')
    def _set_pattern(self, value: object) -> None:
        chosen = int(to_integer(value, 'Long'))
        names = {value: name for name, value in PATTERNS.items()}
        if chosen not in names:
            raise error(1004)
        pattern = self.pattern(True)
        assert pattern is not None
        pattern.set('patternType', names[chosen])
        self.changed()

    @member
    def PatternColor(self) -> object:
        return self.fill_color(False, foreground=True)

    @setter('PatternColor')
    def _set_pattern_color(self, value: object) -> None:
        self.set_fill_color(styles.Color.from_bgr(int(to_integer(value, 'Long'))), foreground=True)

    @member
    def PatternColorIndex(self) -> object:
        return self.fill_color(True, foreground=True)

    @setter('PatternColorIndex')
    def _set_pattern_color_index(self, value: object) -> None:
        self.set_fill_color(palette(value), foreground=True)


class ConditionalBorder(_Patch):
    vba_type_name = 'Border'

    def __init__(self, root: XmlElement, changed: Callable[[], None], sheet: Worksheet, index: int,
                 *, color_source: Callable[[], XmlElement] | None = None) -> None:
        super().__init__(root, changed, sheet, color_source=color_source)
        self.index = index

    def side(self, create: bool = False, *, root: XmlElement | None = None) -> XmlElement | None:
        index = LEGACY_SIDES.get(self.index, self.index)
        name = SIDES.get(index, {5: 'diagonal', 6: 'diagonal', 11: 'vertical', 12: 'horizontal'}.get(index, ''))
        if not name:
            raise error(1004)
        selected = self.root if root is None else root
        border = child(selected, 'border') if create else selected.find(NS + 'border')
        return None if border is None else child(border, name) if create else border.find(NS + name)

    def read(self, what: str) -> object:
        colored = what in ('Color', 'ColorIndex', 'TintAndShade', 'ThemeColor')
        side = self.side(root=self.color_source()) if colored and self.color_source is not None else self.side()
        color = read_color(None if side is None else side.find(NS + 'color'))
        if what in ('TintAndShade', 'ThemeColor'):
            if color is None:
                return NULL
            if what == 'TintAndShade':
                return color.tint / styles.TINT_SCALE
            if color.kind != 'theme':
                raise error(5)
            return VBAInt(int(color.value) + 1, 'Long')
        if what in ('Color', 'ColorIndex'):
            if color is not None:
                return self.color_answer(color, what == 'ColorIndex')
            active = side is not None and side.get('style', 'none') != 'none'
            return VBAInt(styles.XL_AUTOMATIC if active else styles.XL_NONE, 'Long') if what == 'ColorIndex' else 0.0
        if self.index not in LEGACY_SIDES:
            raise error(1004)
        style = '' if side is None else side.get('style', 'thin' if color is not None else '')
        if not style:
            return NULL
        return VBAInt(LINES.get(style, (styles.XL_NONE, 2))[int(what == 'Weight')], 'Long')

    def assign(self, what: str, value: object) -> None:
        if self.index not in LEGACY_SIDES and what not in ('Color', 'ColorIndex'):
            raise error(1004)
        if what in ('Color', 'ColorIndex'):
            color = palette(value) if what == 'ColorIndex' else styles.Color.from_bgr(int(to_integer(value, 'Long')))
            side = self.side(True)
            assert side is not None
            if side.get('style') in (None, 'none'):
                # Native color assignment creates a visible thin line. Without
                # this attribute Excel discards the saved conditional color.
                side.set('style', 'thin')
            write_color(side, 'color', color)
        else:
            current = self.side()
            old = styles.Side('' if current is None else current.get('style', ''))
            selected = int(to_integer(value, 'Long'))
            if what == 'Weight' and selected not in (1, 2):
                raise error(1004)
            if what == 'LineStyle' and selected not in (1, -4115, 4, 5, -4118, styles.XL_NONE):
                raise error(1004)
            result = with_weight(old, selected) if what == 'Weight' else with_line_style(old, selected)
            side = self.side(True)
            assert side is not None
            side.set('style', result.style or 'none')
            if result.style and side.find(NS + 'color') is None:
                # A line given no colour of its own is drawn in the automatic one, which Excel writes
                # (tests/fixtures/conditional_joins/, j_border).
                write_color(side, 'color', styles.Color('auto', 1))
        self.changed()

    @member
    def LineStyle(self) -> object:
        return self.read('LineStyle')

    @setter('LineStyle')
    def _set_line_style(self, value: object) -> None:
        self.assign('LineStyle', value)

    @member
    def Weight(self) -> object:
        return self.read('Weight')

    @setter('Weight')
    def _set_weight(self, value: object) -> None:
        self.assign('Weight', value)

    @member
    def Color(self) -> object:
        return self.read('Color')

    @setter('Color')
    def _set_color(self, value: object) -> None:
        self.assign('Color', value)

    @member
    def ColorIndex(self) -> object:
        return self.read('ColorIndex')

    @setter('ColorIndex')
    def _set_color_index(self, value: object) -> None:
        self.assign('ColorIndex', value)


    @member
    def TintAndShade(self) -> object:
        return self.read('TintAndShade')

    @member
    def ThemeColor(self) -> object:
        return self.read('ThemeColor')


class ConditionalBorders(VBACollection, _Patch):
    vba_type_name = 'Borders'

    def vba_items(self) -> list[object]:
        return [ConditionalBorder(self.root, self.changed, self.sheet, index, color_source=self.color_source)
                for index in (1, 2, 3, 4)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        selected = int(to_integer(index, 'Long'))
        if selected not in (*LEGACY_SIDES, *SIDES, 5, 6, 11, 12):
            raise error(1004)
        return ConditionalBorder(self.root, self.changed, self.sheet, selected, color_source=self.color_source)

    def read(self, what: str) -> object:
        return uniform(ConditionalBorder(self.root, self.changed, self.sheet, index,
                                        color_source=self.color_source).read(what) for index in (1, 2, 3, 4))

    def assign(self, what: str, value: object) -> None:
        for index in (1, 2, 3, 4):
            ConditionalBorder(self.root, self.changed, self.sheet, index).assign(what, value)

    @member
    def TintAndShade(self) -> object:
        return self.read('TintAndShade')

    @member
    def ThemeColor(self) -> object:
        return self.read('ThemeColor')

    @member
    def LineStyle(self) -> object:
        return self.read('LineStyle')

    @setter('LineStyle')
    def _set_line_style(self, value: object) -> None:
        self.assign('LineStyle', value)

    @member
    def Weight(self) -> object:
        return self.read('Weight')

    @setter('Weight')
    def _set_weight(self, value: object) -> None:
        self.assign('Weight', value)

    @member
    def Color(self) -> object:
        return self.read('Color')

    @setter('Color')
    def _set_color(self, value: object) -> None:
        self.assign('Color', value)

    @member
    def ColorIndex(self) -> object:
        return self.read('ColorIndex')

    @setter('ColorIndex')
    def _set_color_index(self, value: object) -> None:
        self.assign('ColorIndex', value)
