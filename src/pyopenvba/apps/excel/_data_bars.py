"""Data bars: FormatConditions.AddDatabar, the Databar and what hangs off it, and how a file keeps one.

Measured in live Excel (scripts/measure_data_bars_icons.py, tests/fixtures/data_bars_icons/):

- AddDatabar makes a gradient bar of 638EC6 with no border, from the lowest value to the highest (MinPoint and
  MaxPoint, xlConditionValueLowestValue and xlConditionValueHighestValue), 0 to 100 percent of the cell's
  width, in the sheet's direction, its value shown, no axis, negative bars in the bar's own colours.
- MinPoint and MaxPoint change through Modify alone: setting Type or Value is error 450, Modify with no type
  450 and a type past 0 to 7 error 5. A minimum cannot be the highest value or an automatic maximum, nor a
  maximum the lowest value or an automatic minimum, error 1004; a percent or percentile number outside 0 to
  100 is error 5. A number given stays one; text is a formula, as is a formula point's number, and Value reads
  it after "=", "=" alone for none given, until the file is opened again, when a number a file spells reads as
  one and a point with none reads 0. A lowest, highest or automatic point keeps a number given it as long.
- PercentMin and PercentMax run from 0 to 100, the minimum not past the maximum, else error 5; so the fill
  type, the direction, the axis position and the border's and negative bars' colour types, each error 5 past
  its values. While a negative bar's colour or border is the bar's own it reads the bar's, and setting it is
  error 1004; made its own, it starts as the bar's colour or border colour of that moment. The border's and the
  axis's colours can be set while neither is drawn, but a file keeps them only when drawn.
- StopIfTrue, Formula and ScopeType are as every point rule's (_conditional.PointRule).
- A bar draws nothing DisplayFormat shows. With ShowValue False a number it covers shows no text
  (_conditional_eval.value_hidden).
- A file keeps a bar as a cfRule of type dataBar holding its points, its colour, its lengths (none for 0 to
  100, else each but a 10 or a 90) and showValue="0" for a hidden value. A bar Excel 2007 could not draw -- a
  length of 0 or 100, a solid fill, a border, a direction, negative bars of their own colour or border, an
  axis, an automatic point -- adds an x14:cfRule (_conditional_extension) with its lengths but 10 and 90,
  border="1", gradient="0", its direction, negativeBarColorSameAsPositive="1" while the negative bars share
  the colour, negativeBarBorderColorSameAsPositive="0" for a drawn border of their own, its axis position but
  automatic, then its points and the border, negative and axis colours of what it draws. A point whose formula
  names another sheet makes that x14:cfRule all there is, with the bar's priority and its colour as fillColor.

What was not measured reports itself: an x14-only bar with its value hidden or a formula of its own.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import cast

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._colour_scales import FormatColor
from pyopenvba.apps.excel._conditional import NS, PointRule, Rule, guard
from pyopenvba.apps.excel._conditional_extension import (
    NUMBER,
    X14,
    XM,
    RuleState,
    foreign,
    given_text,
    identified,
    link,
    names_a_sheet,
    x14_color,
    x14_point,
)
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAInt, error, to_bool, to_integer, to_number, to_text

#: XlConditionValueTypes as a file spells them, the automatic ones Excel 2010's alone.
KINDS = {0: 'num', 1: 'min', 2: 'max', 3: 'percent', 4: 'formula', 5: 'percentile', 6: 'autoMin', 7: 'autoMax'}
_NUMBERS = {name: number for number, name in KINDS.items()}
#: What each point cannot be: a minimum no highest or automatic maximum, a maximum no lowest or automatic minimum.
_BARRED = ({'max', 'autoMax'}, {'min', 'autoMin'})
_DIRECTIONS = {-5002: 'context', -5003: 'leftToRight', -5004: 'rightToLeft'}
_AXES = {0: 'automatic', 1: 'middle', 2: 'none'}
_BAR_ATTRIBUTES = {'minLength', 'maxLength', 'border', 'gradient', 'direction', 'negativeBarColorSameAsPositive',
                   'negativeBarBorderColorSameAsPositive', 'axisPosition'}
_COLOURS = ('fillColor', 'borderColor', 'negativeFillColor', 'negativeBorderColor', 'axisColor')
_COLOUR_ATTRIBUTES = {'rgb', 'theme', 'tint', 'indexed', 'auto'}


def _colour(rgb: str) -> XmlElement:
    return XmlElement(NS + 'color', {'rgb': rgb})


def _keys(element: XmlElement) -> set[str]:
    return {key for key in element.attrib if key != 'xmlns' and not key.startswith('xmlns:')}


def _true(text: str | None, default: bool) -> bool:
    return default if text is None else text in ('1', 'true')


@dataclass
class Point:
    """A data bar's end: its kind, and a number, a formula's text without its "=", or None for none given."""

    kind: str
    value: float | str | None = None

    def read(self) -> object:
        """What ConditionValue.Value answers."""
        if isinstance(self.value, str):
            return '=' + self.value
        return 0.0 if self.value is None else self.value

    def written(self) -> str | None:
        """The value a file keeps, which a lowest, highest or automatic point keeps none of."""
        value = self.value
        if self.kind in ('min', 'max', 'autoMin', 'autoMax') or value is None or value == '':
            return None
        return value if isinstance(value, str) else number_text(value)


def read_value(kind: str, text: str | None) -> float | str | None:
    """A point's value as a file spells it: a formula's text, else a number where it spells one."""
    if text is None:
        return None
    return float(text) if kind != 'formula' and NUMBER.fullmatch(text) else text


@dataclass(eq=False)
class Bar(RuleState):
    """A data bar's settings; each colour an element holding its rgb, or its theme and tint."""

    points: list[Point]
    color: XmlElement
    show_value: bool = True
    min_length: int = 0
    max_length: int = 100
    gradient: bool = True
    border: bool = False
    border_color: XmlElement = field(default_factory=lambda: _colour('FF000000'))
    direction: str = 'context'
    #: The negative bars' own colour and border colour; None while they are the bar's.
    negative_color: XmlElement | None = None
    negative_border_color: XmlElement | None = None
    #: Whether the negative bars' border colour is the bar's; that alone can differ while no border is drawn.
    negative_border_same: bool = True
    axis: str = 'none'
    axis_color: XmlElement = field(default_factory=lambda: _colour('FF000000'))

    def extended(self) -> bool:
        """Whether Excel 2007 could not draw the bar, so a file adds its x14:cfRule."""
        return (self.min_length == 0 or self.max_length == 100 or not self.gradient or self.border
                or self.direction != 'context' or self.negative_color is not None or not self.negative_border_same
                or self.axis != 'none' or any(point.kind in ('autoMin', 'autoMax') for point in self.points))

    def parts(self, rule: Rule) -> tuple[XmlElement | None, XmlElement | None]:
        alone = any(isinstance(point.value, str) and names_a_sheet(point.value) for point in self.points)
        if not (alone or self.extended()):
            return self._main(rule, ''), None
        identifier = identified(rule)
        return (None if alone else self._main(rule, identifier)), self._extension(rule, identifier, alone)

    def _main(self, rule: Rule, identifier: str) -> XmlElement:
        root = XmlElement(NS + 'cfRule', {'type': 'dataBar', 'priority': rule.root.get('priority', '1')})
        for formula in rule.root.findall(NS + 'formula'):
            root.append(deepcopy(formula))
        # Excel 2007 draws 10 to 90 percent of the width, which is what a file shows for Excel 2010's 0 to 100.
        low, high = (10, 90) if (self.min_length, self.max_length) == (0, 100) else (self.min_length, self.max_length)
        attributes = {**({'minLength': str(low)} if low != 10 else {}), **({'maxLength': str(high)} if high != 90 else {}),
                      **({} if self.show_value else {'showValue': '0'})}
        bar = XmlElement(NS + 'dataBar', attributes)
        for point in self.points:
            kind = {'autoMin': 'min', 'autoMax': 'max'}.get(point.kind, point.kind)
            written = point.written()
            bar.append(XmlElement(NS + 'cfvo', {'type': kind} if written is None else {'type': kind, 'val': written}))
        bar.append(XmlElement(NS + 'color', dict(self.color.attrib)))
        root.append(bar)
        if identifier:
            link(root, identifier)
        return root

    def _extension(self, rule: Rule, identifier: str, alone: bool) -> XmlElement:
        if alone and (not self.show_value or rule.root.find(NS + 'formula') is not None):
            raise VBAUnsupportedError('Saving a data bar naming another sheet with its value hidden or a formula of '
                                      'its own is not implemented')
        entry = XmlElement(X14 + 'cfRule', {'type': 'dataBar', **({'priority': rule.root.get('priority', '1')}
                                                                   if alone else {}), 'id': identifier})
        drawn_border = self.border and not self.negative_border_same
        attributes = {**({'minLength': str(self.min_length)} if self.min_length != 10 else {}),
                      **({'maxLength': str(self.max_length)} if self.max_length != 90 else {}),
                      **({'border': '1'} if self.border else {}), **({} if self.gradient else {'gradient': '0'}),
                      **({'direction': self.direction} if self.direction != 'context' else {}),
                      **({'negativeBarColorSameAsPositive': '1'} if self.negative_color is None else {}),
                      **({'negativeBarBorderColorSameAsPositive': '0'} if drawn_border else {}),
                      **({'axisPosition': self.axis} if self.axis != 'automatic' else {})}
        bar = XmlElement(X14 + 'dataBar', attributes)
        for point in self.points:
            bar.append(x14_point(point.kind, point.written()))
        if alone:
            bar.append(x14_color(self.color, 'fillColor'))
        if self.border:
            bar.append(x14_color(self.border_color, 'borderColor'))
        if self.negative_color is not None:
            bar.append(x14_color(self.negative_color, 'negativeFillColor'))
        if drawn_border and self.negative_border_color is not None:
            bar.append(x14_color(self.negative_border_color, 'negativeBorderColor'))
        if self.axis != 'none':
            bar.append(x14_color(self.axis_color, 'axisColor'))
        entry.append(bar)
        return entry


def made() -> tuple[XmlElement, Bar]:
    """AddDatabar's rule, a stand-in cfRule and the bar's settings."""
    bar = Bar([Point('min'), Point('max')], _colour('FF638EC6'))
    bar.dirty = True
    return XmlElement(NS + 'cfRule', {'type': 'dataBar'}), bar


def extension_known(entry: XmlElement, *, paired: bool) -> bool:
    """Whether an x14:cfRule of type dataBar holds only what this module reads."""
    if not _keys(entry) <= ({'type', 'id'} | (set() if paired else {'priority'})) or len(entry) != 1:
        return False
    bar = entry[0]
    if bar.tag != X14 + 'dataBar' or not _keys(bar) <= _BAR_ATTRIBUTES:
        return False
    children = list(bar)
    points = children[:2]
    if len(points) != 2 or any(point.tag != X14 + 'cfvo' or not _keys(point) <= {'type'}
                               or point.get('type') not in _NUMBERS or len(point) > 1
                               or any(child.tag != XM + 'f' for child in point) for point in points):
        return False
    for colour in children[2:]:
        if colour.tag[len(X14):] not in _COLOURS or not colour.tag.startswith(X14) or len(colour) \
                or not _keys(colour) <= _COLOUR_ATTRIBUTES:
            return False
    try:
        lengths = [int(bar.get(name, '50')) for name in ('minLength', 'maxLength')]
    except ValueError:
        return False
    return (all(0 <= length <= 100 for length in lengths)
            and bar.get('direction', 'context') in _DIRECTIONS.values()
            and bar.get('axisPosition', 'automatic') in _AXES.values()
            and (paired or bar.find(X14 + 'fillColor') is not None))


def parse(rule: Rule) -> Bar:
    """A bar's settings as a file kept them: its cfRule's, and its x14:cfRule's where it has one."""
    main = None if rule.extension_only else rule.root.find(NS + 'dataBar')
    extra = rule.extension[0] if rule.extension is not None and len(rule.extension) else None
    if main is None and extra is None:
        raise VBAUnsupportedError('A data bar the file leaves out is not implemented')
    shown = main.find(NS + 'color') if main is not None else None
    if shown is None and extra is not None:
        shown = extra.find(X14 + 'fillColor')
    color = XmlElement(NS + 'color', dict(shown.attrib)) if shown is not None else _colour('FF638EC6')
    if extra is None:
        assert main is not None
        points = [Point(point.get('type', ''), read_value(point.get('type', ''), point.get('val')))
                  for point in main.findall(NS + 'cfvo')]
        bar = Bar(points, color, show_value=_true(main.get('showValue'), True),
                  min_length=int(main.get('minLength', '10')), max_length=int(main.get('maxLength', '90')))
    else:
        points: list[Point] = []
        for point in extra.findall(X14 + 'cfvo'):
            formula = point.find(XM + 'f')
            points.append(Point(point.get('type', ''), read_value(point.get('type', ''),
                                                                  None if formula is None else formula.text or '')))

        def found(local: str) -> XmlElement | None:
            # An element without children is false, so callers ask for None outright.
            element = extra.find(X14 + local)
            return None if element is None else XmlElement(NS + 'color', dict(element.attrib))

        border_color, negative, axis_color = found('borderColor'), found('negativeFillColor'), found('axisColor')
        same = _true(extra.get('negativeBarColorSameAsPositive'), False)
        bar = Bar(points, color, show_value=main is None or _true(main.get('showValue'), True),
                  min_length=int(extra.get('minLength', '10')), max_length=int(extra.get('maxLength', '90')),
                  gradient=_true(extra.get('gradient'), True), border=_true(extra.get('border'), False),
                  border_color=_colour('FF000000') if border_color is None else border_color,
                  direction=extra.get('direction', 'context'),
                  negative_color=None if same else deepcopy(color) if negative is None else negative,
                  negative_border_color=found('negativeBorderColor'),
                  negative_border_same=_true(extra.get('negativeBarBorderColorSameAsPositive'), True),
                  axis=extra.get('axisPosition', 'automatic'),
                  axis_color=_colour('FF000000') if axis_color is None else axis_color)
    if len(bar.points) != 2 or any(point.kind not in _NUMBERS for point in bar.points):
        raise VBAUnsupportedError('A data bar with other than two known points is not implemented')
    return bar


# --- VBA -----------------------------------------------------------------------------------------------------


class Databar(PointRule):
    vba_type_name = 'Databar'

    def bar(self) -> Bar:
        """The bar's settings, read from the file the first time they are wanted."""
        foreign(self.sheet)
        if self.rule.state is None:
            self.rule.state = parse(self.rule)
        return cast('Bar', self.rule.state)

    def touched(self) -> None:
        self.bar().dirty = True
        super().touched()

    def assign(self, name: str, value: object) -> None:
        guard(self.sheet)
        setattr(self.bar(), name, value)
        self.touched()

    @member
    def Type(self) -> object:
        return VBAInt(4, 'Long')

    @member
    def MinPoint(self) -> object:
        return ConditionValue(self, 0)

    @member
    def MaxPoint(self) -> object:
        return ConditionValue(self, 1)

    @member
    def BarColor(self) -> object:
        return FormatColor(self, self.bar().color)

    @member
    def BarFillType(self) -> object:
        return VBAInt(1 if self.bar().gradient else 0, 'Long')

    @setter('BarFillType')
    def set_fill_type(self, value: object) -> None:
        self.assign('gradient', _choice(value, (0, 1)) == 1)

    @member
    def BarBorder(self) -> object:
        return DataBarBorder(self)

    @member
    def Direction(self) -> object:
        return VBAInt(next(number for number, name in _DIRECTIONS.items() if name == self.bar().direction), 'Long')

    @setter('Direction')
    def set_direction(self, value: object) -> None:
        self.assign('direction', _DIRECTIONS[_choice(value, tuple(_DIRECTIONS))])

    @member
    def ShowValue(self) -> object:
        return self.bar().show_value

    @setter('ShowValue')
    def set_show_value(self, value: object) -> None:
        self.assign('show_value', to_bool(value))

    @member
    def PercentMin(self) -> object:
        return VBAInt(self.bar().min_length, 'Long')

    @setter('PercentMin')
    def set_percent_min(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if not 0 <= wanted <= self.bar().max_length:
            raise error(5)
        self.assign('min_length', wanted)

    @member
    def PercentMax(self) -> object:
        return VBAInt(self.bar().max_length, 'Long')

    @setter('PercentMax')
    def set_percent_max(self, value: object) -> None:
        wanted = int(to_integer(value, 'Long'))
        if not self.bar().min_length <= wanted <= 100:
            raise error(5)
        self.assign('max_length', wanted)

    @member
    def AxisPosition(self) -> object:
        return VBAInt(next(number for number, name in _AXES.items() if name == self.bar().axis), 'Long')

    @setter('AxisPosition')
    def set_axis_position(self, value: object) -> None:
        self.assign('axis', _AXES[_choice(value, tuple(_AXES))])

    @member
    def AxisColor(self) -> object:
        return FormatColor(self, self.bar().axis_color)

    @member
    def NegativeBarFormat(self) -> object:
        return NegativeBarFormat(self)


def _choice(value: object, allowed: tuple[int, ...]) -> int:
    wanted = int(to_integer(value, 'Long'))
    if wanted not in allowed:
        raise error(5)
    return wanted


class ConditionValue(ExcelObject):
    """A data bar's MinPoint or MaxPoint."""

    vba_type_name = 'ConditionValue'

    def __init__(self, owner: Databar, index: int) -> None:
        self.owner, self.index = owner, index

    def _point(self) -> Point:
        return self.owner.bar().points[self.index]

    @member
    def Type(self) -> object:
        return VBAInt(_NUMBERS[self._point().kind], 'Long')

    @setter('Type')
    def set_type(self, value: object) -> None:
        raise error(450)

    @member
    def Value(self) -> object:
        return self._point().read()

    @setter('Value')
    def set_value(self, value: object) -> None:
        raise error(450)

    @method
    def Modify(self, newtype: object = MISSING, newvalue: object = MISSING) -> object:
        sheet = self.owner.sheet
        guard(sheet)
        if newtype is MISSING:
            raise error(450)
        kind = KINDS.get(int(to_integer(newtype, 'Long')))
        if kind is None:
            raise error(5)
        if kind in _BARRED[self.index]:
            raise error(1004)
        given: float | str
        if newvalue is MISSING:
            given = ''
        elif isinstance(newvalue, str):
            given = given_text(newvalue, sheet.name)
        elif kind == 'formula':
            given = to_text(newvalue)
        else:
            given = float(to_number(newvalue))
            if kind in ('percent', 'percentile') and not 0 <= given <= 100:
                raise error(5)
        point = self._point()
        if isinstance(point.value, str) and names_a_sheet(point.value) and isinstance(given, float):
            # A point holding a formula that names another sheet takes a new type but keeps that formula for a
            # number given it, as Excel keeps it; text or no value replaces it.
            given = point.value
        point.kind, point.value = kind, given
        self.owner.touched()
        return EMPTY


class _Fixed(FormatColor):
    """A negative bar's colour while it is the bar's own: read as that, error 1004 to set."""

    vba_type_name = 'FormatColor'

    def _write(self, color: object) -> None:
        raise error(1004)


class DataBarBorder(ExcelObject):
    vba_type_name = 'DataBarBorder'

    def __init__(self, owner: Databar) -> None:
        self.owner = owner

    @member
    def Type(self) -> object:
        return VBAInt(1 if self.owner.bar().border else 0, 'Long')

    @setter('Type')
    def set_type(self, value: object) -> None:
        self.owner.assign('border', _choice(value, (0, 1)) == 1)

    @member
    def Color(self) -> object:
        return FormatColor(self.owner, self.owner.bar().border_color)


class NegativeBarFormat(ExcelObject):
    vba_type_name = 'NegativeBarFormat'

    def __init__(self, owner: Databar) -> None:
        self.owner = owner

    @member
    def ColorType(self) -> object:
        return VBAInt(1 if self.owner.bar().negative_color is None else 0, 'Long')

    @setter('ColorType')
    def set_color_type(self, value: object) -> None:
        guard(self.owner.sheet)
        bar = self.owner.bar()
        bar.negative_color = None if _choice(value, (0, 1)) == 1 else deepcopy(bar.color)
        self.owner.touched()

    @member
    def Color(self) -> object:
        bar = self.owner.bar()
        return _Fixed(self.owner, bar.color) if bar.negative_color is None else FormatColor(self.owner, bar.negative_color)

    @member
    def BorderColorType(self) -> object:
        return VBAInt(1 if self.owner.bar().negative_border_same else 0, 'Long')

    @setter('BorderColorType')
    def set_border_color_type(self, value: object) -> None:
        guard(self.owner.sheet)
        bar = self.owner.bar()
        bar.negative_border_same = _choice(value, (0, 1)) == 1
        bar.negative_border_color = None if bar.negative_border_same else deepcopy(bar.border_color)
        self.owner.touched()

    @member
    def BorderColor(self) -> object:
        bar = self.owner.bar()
        if bar.negative_border_same or bar.negative_border_color is None:
            return _Fixed(self.owner, bar.border_color)
        return FormatColor(self.owner, bar.negative_border_color)
