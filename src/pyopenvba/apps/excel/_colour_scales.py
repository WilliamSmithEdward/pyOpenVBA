"""Colour scales: FormatConditions.AddColorScale, the ColorScale and its criteria, and the colour a cell shows.

Measured in live Excel (scripts/measure_colour_scales.py,
tests/fixtures/colour_scales/):

- AddColorScale(2) runs from FF7128 at the lowest value to FFEF9C at the
  highest, AddColorScale(3) from F8696B through FFEB84 at the 50th
  percentile to 63BE7B. Any other count is error 9 and none error 450. The
  rule takes the next priority, as every new rule does.
- A criterion is the lowest or the highest value, a number, a percent of
  the way from the lowest to the highest, a percentile of the numbers, or
  a formula's answer. A lowest or highest one's Value reads that value of
  the range as it stands, and setting it is error 1004. Its FormatColor is
  an RGB colour, or a theme colour with a tint Excel keeps in 32767ths.
- Only numbers are coloured: text, blanks, errors and logicals keep their
  fill, and a percentile counts the numbers alone. A cell takes the colour
  of its place between the criteria, the first colour at or below the
  lowest and the last at or above the highest, and the last when the
  lowest and the highest are one value. Excel blends in doubles: the
  cell's place as a fraction of the whole span, then of its segment, each
  channel stepping from the segment's first colour by that fraction of the
  way to the next, cut toward zero. All 246 colours measured agree.
- A scale is a fill at its priority, so a higher rule's fill shows over
  it. StopIfTrue reads False and setting it is error 450.
- A file keeps a scale as a cfRule of type colorScale holding a cfvo for
  each criterion, then a color for each colour.

What was not measured reports itself: a scale whose lowest point lies
above its highest, a middle point outside them, and a formula point that
is not a number.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel import _styles as styles
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._conditional import NS, Rule, RuleObject, changed, formula_anchor, guard
from pyopenvba.apps.excel._differential import read_color
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBACollection, member, setter
from pyopenvba.interpreter._values import VBAInt, error, to_integer, to_number, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: XlConditionValueTypes as a file spells them.
_CFVO = {0: 'num', 1: 'min', 2: 'max', 3: 'percent', 4: 'formula', 5: 'percentile'}
_KINDS = {name: number for number, name in _CFVO.items()}
#: Each criterion AddColorScale makes: its cfvo type and value, and its colour.
_DEFAULTS: dict[int, list[tuple[str, str | None, str]]] = {
    2: [('min', None, 'FFFF7128'), ('max', None, 'FFFFEF9C')],
    3: [('min', None, 'FFF8696B'), ('percentile', '50', 'FFFFEB84'), ('max', None, 'FF63BE7B')],
}


def made(count: int) -> XmlElement:
    """AddColorScale's rule, Excel's two or three default colours; any other count is error 9."""
    if count not in _DEFAULTS:
        raise error(9)
    root = XmlElement(NS + 'cfRule', {'type': 'colorScale'})
    scale = XmlElement(NS + 'colorScale')
    root.append(scale)
    for kind, value, _ in _DEFAULTS[count]:
        scale.append(XmlElement(NS + 'cfvo', {'type': kind} if value is None else {'type': kind, 'val': value}))
    for _, _, rgb in _DEFAULTS[count]:
        scale.append(XmlElement(NS + 'color', {'rgb': rgb}))
    return root


def _scale(rule: Rule) -> tuple[list[XmlElement], list[XmlElement]]:
    """A scale's cfvo and color elements, two or three of each."""
    scale = rule.root.find(NS + 'colorScale')
    points = [] if scale is None else scale.findall(NS + 'cfvo')
    colours = [] if scale is None else scale.findall(NS + 'color')
    if len(points) not in (2, 3) or len(colours) != len(points):
        raise VBAUnsupportedError('A colour scale with other than two or three matched points is not implemented')
    return points, colours


# --- the colour a cell shows ---------------------------------------------------------------------------


def scale_patch(sheet: Worksheet, rule: Rule, row: int, column: int) -> XmlElement | None:
    """The fill a colour scale gives a cell, as a dxf; None for a cell it leaves, as it leaves all but numbers."""
    from pyopenvba.apps.excel._conditional_stats import numbers
    from pyopenvba.apps.excel._engine_book import scalar

    value = scalar(sheet.book.calculator.value_of(sheet.name, row, column))
    population = numbers(sheet, rule)
    if not isinstance(value, float) or not population:
        return None
    points, elements = _scale(rule)
    stops = [_point(sheet, rule, point, population) for point in points]
    colours = [_channels(sheet, element) for element in elements]
    low, high = stops[0], stops[-1]
    if high == low:
        shown = colours[-1]
    elif high < low or not low <= stops[len(stops) // 2] <= high:
        raise VBAUnsupportedError('A colour scale whose points are out of order is not implemented')
    elif value <= low:
        shown = colours[0]
    elif value >= high:
        shown = colours[-1]
    else:
        whole = (value - low) / (high - low)
        places = [(stop - low) / (high - low) for stop in stops]
        shown = colours[-1]
        for first, last, start, end in zip(places, places[1:], colours, colours[1:], strict=False):
            if whole <= last and last > first:
                fraction = (whole - first) / (last - first)
                shown = tuple(a + int((b - a) * fraction) for a, b in zip(start, end, strict=True))
                break
    dxf = XmlElement(NS + 'dxf')
    pattern = XmlElement(NS + 'patternFill')
    pattern.append(XmlElement(NS + 'bgColor', {'rgb': 'FF' + ''.join(f'{one:02X}' for one in shown)}))
    fill = XmlElement(NS + 'fill')
    fill.append(pattern)
    dxf.append(fill)
    return dxf


def _channels(sheet: Worksheet, element: XmlElement) -> tuple[int, ...]:
    shown = sheet.book.stylesheet.colors.rrggbb(read_color(element), 'FFFFFF')
    return tuple(int(shown[at:at + 2], 16) for at in (0, 2, 4))


def _point(sheet: Worksheet, rule: Rule, point: XmlElement, population: tuple[float, ...]) -> float:
    """Where a criterion stands among the range's numbers."""
    kind = point.get('type', '')
    if kind == 'min':
        return population[0]
    if kind == 'max':
        return population[-1]
    text = point.get('val', '0')
    if kind == 'formula':
        return _formula(sheet, rule, text)
    try:
        number = float(text)
    except ValueError:
        raise VBAUnsupportedError('A colour scale point whose value is a formula is not implemented') from None
    if kind == 'num':
        return number
    if kind == 'percent':
        return population[0] + (population[-1] - population[0]) * number / 100
    if kind == 'percentile':
        rank = number / 100 * (len(population) - 1)
        below = int(rank)
        if below + 1 >= len(population):
            return population[-1]
        return population[below] + (rank - below) * (population[below + 1] - population[below])
    raise VBAUnsupportedError(f'A colour scale point of type {kind!r} is not implemented')


def _formula(sheet: Worksheet, rule: Rule, text: str) -> float:
    from pyopenvba.apps.excel._udf import evaluation_origin
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.lexer import FormulaSyntaxError

    calculator = sheet.book.calculator
    anchor = formula_anchor(rule.areas)
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, anchor.top, anchor.left, today=now.date(), now=now)
    try:
        with evaluation_origin(sheet.book.application, 'conditional'):
            found = context.first(context.formula(calculator.engine_book.read('=' + text.removeprefix('='))))
    except FormulaSyntaxError:
        raise VBAUnsupportedError('A colour scale point with an unreadable formula is not implemented') from None
    if not isinstance(found, float):
        raise VBAUnsupportedError('A colour scale point whose formula is not a number is not implemented')
    return found


# --- VBA -----------------------------------------------------------------------------------------------


class ColorScale(RuleObject):
    vba_type_name = 'ColorScale'

    @member
    def Type(self) -> object:
        return VBAInt(3, 'Long')

    @member
    def ColorScaleCriteria(self) -> object:
        return ColorScaleCriteria(self)

    @member
    def StopIfTrue(self) -> object:
        return False

    @setter('StopIfTrue')
    def set_stop(self, value: object) -> None:
        raise error(450)


class ColorScaleCriteria(VBACollection, ExcelObject):
    vba_type_name = 'ColorScaleCriteria'

    def __init__(self, owner: ColorScale) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        points, _ = _scale(self.owner.rule)
        return [ColorScaleCriterion(self.owner, index) for index in range(1, len(points) + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(9)
        return items[position - 1]


class ColorScaleCriterion(ExcelObject):
    vba_type_name = 'ColorScaleCriterion'

    def __init__(self, owner: ColorScale, index: int) -> None:
        self.owner, self.index = owner, index

    def _point(self) -> XmlElement:
        return _scale(self.owner.rule)[0][self.index - 1]

    def _changed(self) -> None:
        changed(self.owner.sheet)

    @member
    def Index(self) -> object:
        return VBAInt(self.index, 'Long')

    @member
    def Type(self) -> object:
        kind = self._point().get('type', '')
        if kind not in _KINDS:
            raise VBAUnsupportedError(f'A colour scale point of type {kind!r} is not implemented')
        return VBAInt(_KINDS[kind], 'Long')

    @setter('Type')
    def set_type(self, value: object) -> None:
        guard(self.owner.sheet)
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _CFVO:
            raise error(5)
        point = self._point()
        point.set('type', _CFVO[wanted])
        if wanted in (1, 2):
            point.attrib.pop('val', None)
        elif point.get('val') is None:
            point.set('val', '0')
        self._changed()

    @member
    def Value(self) -> object:
        """A lowest or highest point's is that number of the range as it stands; a formula's is its text."""
        from pyopenvba.apps.excel._conditional_stats import numbers

        point = self._point()
        kind = point.get('type', '')
        if kind in ('min', 'max'):
            population = numbers(self.owner.sheet, self.owner.rule)
            return (population[0] if kind == 'min' else population[-1]) if population else 0.0
        text = point.get('val', '0')
        if kind == 'formula':
            return '=' + text
        try:
            return float(text)
        except ValueError:
            return '=' + text

    @setter('Value')
    def set_value(self, value: object) -> None:
        guard(self.owner.sheet)
        point = self._point()
        kind = point.get('type', '')
        if kind in ('min', 'max'):
            raise error(1004)
        if kind == 'formula':
            point.set('val', to_text(value).removeprefix('='))
        else:
            point.set('val', number_text(float(to_number(value))))
        self._changed()

    @member
    def FormatColor(self) -> object:
        return FormatColor(self.owner, _scale(self.owner.rule)[1][self.index - 1])


class FormatColor(ExcelObject):
    """A colour scale criterion's colour: RGB, or a theme colour and its tint."""

    vba_type_name = 'FormatColor'

    def __init__(self, owner: RuleObject, element: XmlElement) -> None:
        self.owner, self.element = owner, element

    def _color(self) -> styles.Color:
        found = read_color(self.element)
        if found is None:
            raise VBAUnsupportedError('A colour scale colour the file leaves out is not implemented')
        return found

    def _write(self, color: styles.Color) -> None:
        guard(self.owner.sheet)
        for attribute in ('rgb', 'indexed', 'auto', 'theme', 'tint'):
            self.element.attrib.pop(attribute, None)
        self.element.set(color.kind, str(color.value))
        if color.tint:
            self.element.set('tint', styles.excel_double(color.tint / styles.TINT_SCALE))
        changed(self.owner.sheet)

    @member
    def Color(self) -> object:
        return float(styles.bgr(self.owner.sheet.book.stylesheet.colors.rrggbb(self._color(), 'FFFFFF')))

    @setter('Color')
    def set_color(self, value: object) -> None:
        self._write(styles.Color.from_bgr(int(to_integer(value, 'Long'))))

    @member
    def ThemeColor(self) -> object:
        color = self._color()
        return VBAInt(int(color.value) + 1 if color.kind == 'theme' else 0, 'Long')

    @setter('ThemeColor')
    def set_theme_color(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if not 1 <= wanted <= len(styles.DEFAULT_THEME):
            raise error(5)
        self._write(styles.Color('theme', wanted - 1))

    @member
    def TintAndShade(self) -> object:
        return self._color().tint / styles.TINT_SCALE

    @setter('TintAndShade')
    def set_tint(self, value: object) -> None:
        tint = float(to_number(value))
        if not -1 <= tint <= 1:
            raise error(5)
        color = self._color()
        self._write(styles.Color(color.kind, color.value, styles.tint_n(repr(tint))))
