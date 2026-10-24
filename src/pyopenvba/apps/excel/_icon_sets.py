"""Icon sets: FormatConditions.AddIconSetCondition, the IconSetCondition and its criteria, Workbook.IconSets, and
how a file keeps a set.

Measured in live Excel (scripts/measure_data_bars_icons.py, tests/fixtures/data_bars_icons/):

- AddIconSetCondition makes three traffic lights (xl3TrafficLights1) at 0, 33 and 67 percent of the way from
  the lowest value to the highest, each criterion at or above its value, the values shown beside the icons.
- Workbook.IconSets holds Excel's twenty sets, IconSets(1) to IconSets(20), error 9 past them; an IconSet has
  its ID and Count, and Item(i) an Icon with its Index. IconSetCondition.IconSet takes one by Let (Set is error
  438, and anything but a set is let go without one), and a set of custom icons reads as ID -1, xlCustomSet.
- A new set keeps the criteria, the order and whether values show. Criteria still at their percent defaults
  take the new set's defaults; any others keep their values, the first as many as the new set has, and a
  criterion the new set adds copies the type and value of the last. Each place keeps its operator, a new one
  at or above. A new set's icons are its own: custom icons go.
- The first criterion's Type and Value are error 1004 to set, its Operator and Icon not. A criterion is a
  number, a percent, a formula or a percentile, error 1004 for the others. Made a number its value is 0; a
  percent or percentile, its place's default; a formula, its value as it was. A Value given as a number stays
  one, and text is a formula (_conditional_extension.given_text), read after "=" until the file is opened
  again; no value is out of range. Operator is xlGreaterEqual or xlGreater, error 5 for any other. Icon takes
  any of Excel's icons or xlIconNoCellIcon, error 1004 for any other, and makes the set custom, unreversed.
- ReverseOrder turns the icons round, error 5 on a custom set; ShowIconOnly hides the values, and a number
  then shows no text (_conditional_eval.value_hidden). PercentileValues reads True while every criterion is a
  percentile, and setting it makes every criterion a percentile, or a number, keeping its value.
- StopIfTrue, Formula and ScopeType are as every point rule's (_conditional.PointRule). An icon set draws
  nothing DisplayFormat shows; AutoFilter by icon finds the icon a cell shows (scripts/measure_icon_filters.py,
  tests/fixtures/icon_filters/): a number shows the icon of the highest criterion it reaches, the criteria
  standing among the range's numbers as a colour scale's points do, and a word, a blank or an error shows none.
- A file keeps a set as a cfRule of type iconSet holding a cfvo for each criterion (gte="0" for one above
  its value) in an iconSet element naming the set (but the traffic lights, its default), with showValue="0"
  and reverse="1" as set. Stars, triangles, boxes, custom icons or a formula naming another sheet put the set
  in Excel 2010's extension alone, each value an xm:f and each custom icon a cfIcon naming its place in the
  set the rule has, or where that set lacks it in the first set holding it, or NoIcons for none.

What was not measured reports itself: an icon set whose file declares percent="0", and one in the extension
alone with a formula of its own.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING, cast

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._conditional import NS, PointRule, Rule, guard
from pyopenvba.apps.excel._conditional_extension import (
    NUMBER,
    X14,
    XM,
    RuleState,
    foreign,
    given_text,
    identified,
    names_a_sheet,
    x14_point,
)
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBACollection, member, ref_setter, setter
from pyopenvba.interpreter._values import VBAInt, error, to_bool, to_integer, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet

#: XlIconSet: each set's name in a file and its icons (XlIcon), the lowest criterion's first.
SETS: dict[int, tuple[str, tuple[int, ...]]] = {
    1: ('3Arrows', (3, 2, 1)), 2: ('3ArrowsGray', (6, 5, 4)), 3: ('3Flags', (9, 8, 7)),
    4: ('3TrafficLights1', (12, 11, 10)), 5: ('3TrafficLights2', (16, 15, 14)), 6: ('3Signs', (18, 17, 10)),
    7: ('3Symbols', (21, 20, 19)), 8: ('3Symbols2', (24, 23, 22)), 9: ('4Arrows', (3, 26, 25, 1)),
    10: ('4ArrowsGray', (6, 28, 27, 4)), 11: ('4RedToBlack', (32, 31, 30, 29)), 12: ('4Rating', (38, 39, 40, 41)),
    13: ('4TrafficLights', (13, 12, 11, 10)), 14: ('5Arrows', (3, 26, 2, 25, 1)),
    15: ('5ArrowsGray', (6, 28, 5, 27, 4)), 16: ('5Rating', (37, 38, 39, 40, 41)),
    17: ('5Quarters', (36, 35, 34, 33, 32)), 18: ('3Stars', (44, 43, 42)), 19: ('3Triangles', (47, 46, 45)),
    20: ('5Boxes', (52, 51, 50, 49, 48)),
}
_BY_NAME = {name: number for number, (name, _) in SETS.items()}
#: The sets Excel 2007 did not have, which only the extension can hold.
_NEWER = (18, 19, 20)
_DEFAULT_SET = 4
_NO_ICON = -1
#: Each icon as a custom icon names it: the first set holding it, in the sets' order, and its place there.
_SOURCES: dict[int, tuple[str, int]] = {}
for _name, _icons in SETS.values():
    for _place, _icon in enumerate(_icons):
        _SOURCES.setdefault(_icon, (_name, _place))
_SOURCES[_NO_ICON] = ('NoIcons', 0)
#: IconCriterion.Type's values as a file spells them; the lowest, highest and automatic ones are refused.
_KINDS = {0: 'num', 3: 'percent', 4: 'formula', 5: 'percentile'}
_NUMBERS = {name: number for number, name in _KINDS.items()}
_GREATER, _AT_LEAST = 5, 7


def _defaults(count: int) -> list[float]:
    """A set's percents, evenly from 0: 0, 33 and 67 for three icons."""
    return [float(int(100 * place / count + 0.5)) for place in range(count)]


@dataclass
class Criterion:
    """A threshold: its kind, a number or a formula's text, and whether it is above its value rather than at or
    above it; and its icon once the set is custom."""

    kind: str
    value: float | str
    greater: bool = False
    icon: int | None = None

    def read(self) -> object:
        return '=' + self.value if isinstance(self.value, str) else self.value

    def written(self) -> str:
        return self.value if isinstance(self.value, str) else number_text(self.value)


@dataclass(eq=False)
class Icons(RuleState):
    """An icon set's settings: its set, its criteria, the lowest first, and how it shows."""

    set_id: int
    criteria: list[Criterion]
    reverse: bool = False
    show_value: bool = True
    custom: bool = False

    def icons(self) -> list[int]:
        """The icon each criterion shows."""
        if self.custom:
            return [_NO_ICON if criterion.icon is None else criterion.icon for criterion in self.criteria]
        icons = list(SETS[self.set_id][1])
        return icons[::-1] if self.reverse else icons

    def parts(self, rule: Rule) -> tuple[XmlElement | None, XmlElement | None]:
        alone = self.custom or self.set_id in _NEWER or any(
            isinstance(criterion.value, str) and names_a_sheet(criterion.value) for criterion in self.criteria)
        attributes = {**({'iconSet': SETS[self.set_id][0]} if self.set_id != _DEFAULT_SET else {}),
                      **({} if self.show_value else {'showValue': '0'}), **({'reverse': '1'} if self.reverse else {})}
        if not alone:
            root = XmlElement(NS + 'cfRule', {'type': 'iconSet', 'priority': rule.root.get('priority', '1')})
            for formula in rule.root.findall(NS + 'formula'):
                root.append(deepcopy(formula))
            shown = XmlElement(NS + 'iconSet', attributes)
            for criterion in self.criteria:
                shown.append(XmlElement(NS + 'cfvo', {'type': criterion.kind, 'val': criterion.written(),
                                                      **({'gte': '0'} if criterion.greater else {})}))
            root.append(shown)
            return root, None
        if rule.root.find(NS + 'formula') is not None:
            raise VBAUnsupportedError('Saving an icon set of Excel 2010 with a formula of its own is not implemented')
        entry = XmlElement(X14 + 'cfRule', {'type': 'iconSet', 'priority': rule.root.get('priority', '1'),
                                            'id': identified(rule)})
        shown = XmlElement(X14 + 'iconSet', {**attributes, **({'custom': '1'} if self.custom else {})})
        for criterion in self.criteria:
            shown.append(x14_point(criterion.kind, criterion.written(), greater=criterion.greater))
        if self.custom:
            name, own = SETS[self.set_id]
            for icon in self.icons():
                source, place = (name, own.index(icon)) if icon in own else _SOURCES[icon]
                shown.append(XmlElement(X14 + 'cfIcon', {'iconSet': source, 'iconId': str(place)}))
        entry.append(shown)
        return None, entry


def made() -> tuple[XmlElement, Icons]:
    """AddIconSetCondition's rule, a stand-in cfRule and the set's settings."""
    icons = Icons(_DEFAULT_SET, [Criterion('percent', value) for value in _defaults(3)])
    icons.dirty = True
    return XmlElement(NS + 'cfRule', {'type': 'iconSet'}), icons


def _read_value(kind: str, text: str) -> float | str:
    return float(text) if kind != 'formula' and NUMBER.fullmatch(text) else text


def extension_known(entry: XmlElement, *, paired: bool) -> bool:
    """Whether an x14:cfRule of type iconSet holds only what this module reads: a set the extension holds alone."""
    keys = {key for key in entry.attrib if not key.startswith('xmlns')}
    if paired or not keys <= {'type', 'priority', 'id'} or len(entry) != 1 or entry[0].tag != X14 + 'iconSet':
        return False
    shown = entry[0]
    name = shown.get('iconSet', '3TrafficLights1')
    if not {key for key in shown.attrib if not key.startswith('xmlns')} <= {'iconSet', 'showValue', 'reverse', 'custom'} \
            or name not in _BY_NAME:
        return False
    count = len(SETS[_BY_NAME[name]][1])
    children = list(shown)
    points, icons = children[:count], children[count:]
    if len(points) != count or any(point.tag != X14 + 'cfvo' or point.get('type') not in _NUMBERS
                                   or point.find(XM + 'f') is None for point in points):
        return False
    return not icons or (len(icons) == count and all(icon.tag == X14 + 'cfIcon' and _source(icon) for icon in icons))


def _source(icon: XmlElement) -> bool:
    """Whether a cfIcon names a place in one of Excel's sets, or no icon."""
    name, place = icon.get('iconSet', ''), icon.get('iconId', '')
    if not place.isdigit():
        return False
    return (name == 'NoIcons' and place == '0') or (name in _BY_NAME and int(place) < len(SETS[_BY_NAME[name]][1]))


def parse(rule: Rule) -> Icons:
    """A set's settings as a file kept them, in its cfRule or in the extension alone."""
    if rule.extension_only:
        shown = rule.extension[0] if rule.extension is not None and len(rule.extension) else None
    else:
        shown = rule.root.find(NS + 'iconSet')
    if shown is None or shown.get('percent', '1') not in ('1', 'true'):
        raise VBAUnsupportedError('An icon set the file keeps in this form is not implemented')
    name = shown.get('iconSet', '3TrafficLights1')
    if name not in _BY_NAME:
        raise VBAUnsupportedError(f'The icon set {name!r} is not implemented')
    criteria: list[Criterion] = []
    for point in shown.findall((X14 if rule.extension_only else NS) + 'cfvo'):
        kind = point.get('type', '')
        if kind not in _NUMBERS:
            raise VBAUnsupportedError(f'An icon set criterion of type {kind!r} is not implemented')
        if rule.extension_only:
            formula = point.find(XM + 'f')
            text = '' if formula is None else formula.text or ''
        else:
            text = point.get('val', '0')
        criteria.append(Criterion(kind, _read_value(kind, text), point.get('gte', '1') in ('0', 'false')))
    icons = Icons(_BY_NAME[name], criteria, reverse=shown.get('reverse', '0') in ('1', 'true'),
                  show_value=shown.get('showValue', '1') in ('1', 'true'), custom=shown.get('custom', '0') in ('1', 'true'))
    if icons.custom:
        chosen = shown.findall(X14 + 'cfIcon')
        for criterion, icon in zip(criteria, chosen, strict=True):
            source = icon.get('iconSet', '')
            criterion.icon = _NO_ICON if source == 'NoIcons' else SETS[_BY_NAME[source]][1][int(icon.get('iconId', '0'))]
    if len(criteria) != len(SETS[icons.set_id][1]):
        raise VBAUnsupportedError('An icon set with other than its own number of criteria is not implemented')
    return icons


def state_of(sheet: Worksheet, rule: Rule) -> Icons:
    """A rule's icon set settings, read from the file the first time they are wanted."""
    foreign(sheet)
    if rule.state is None:
        rule.state = parse(rule)
    return cast('Icons', rule.state)


def covering(sheet: Worksheet, row: int, column: int) -> Rule | None:
    """The icon set that shows a cell its icon: the first by priority whose range holds the cell."""
    rules = sorted((rule for rule in sheet.conditional_formats if not rule.deleted
                    and rule.root.get('type') == 'iconSet'), key=lambda rule: int(rule.root.get('priority', '0')))
    return next((rule for rule in rules if any(area.top <= row <= area.bottom and area.left <= column <= area.right
                                               for area in rule.areas)), None)


def shown_icon(sheet: Worksheet, row: int, column: int) -> tuple[int, int] | None:
    """The icon a cell shows, as the set and the place there a filter by it names; None for a cell showing none --
    a word, a blank, an error, one no set covers, or a place given no icon (tests/fixtures/icon_filters/).

    A cell shows the icon of the highest criterion it reaches, at or above its value or above it as the criterion
    says, the criteria standing among the range's numbers as a colour scale's points stand.
    """
    from pyopenvba.apps.excel._colour_scales import point_value
    from pyopenvba.apps.excel._conditional_stats import numbers
    from pyopenvba.apps.excel._engine_book import scalar

    rule = covering(sheet, row, column)
    if rule is None:
        return None
    value = scalar(sheet.book.calculator.value_of(sheet.name, row, column))
    population = numbers(sheet, rule)
    if not isinstance(value, float) or not population:
        return None
    icons = state_of(sheet, rule)
    band = 0
    for place, criterion in enumerate(icons.criteria[1:], 1):
        threshold = point_value(sheet, rule, criterion.kind, criterion.written(), population)
        if value > threshold or (not criterion.greater and value == threshold):
            band = place
    if icons.custom:
        icon = icons.icons()[band]
        if icon == _NO_ICON:
            return None
        own = SETS[icons.set_id][1]
        if icon in own:
            return icons.set_id, own.index(icon) + 1
        source, at = _SOURCES[icon]
        return _BY_NAME[source], at + 1
    return icons.set_id, (len(icons.criteria) - band if icons.reverse else band + 1)


def set_named(name: str) -> int | None:
    """The set a file names, by its XlIconSet number."""
    return _BY_NAME.get(name)


# --- VBA -----------------------------------------------------------------------------------------------------


class IconSetCondition(PointRule):
    vba_type_name = 'IconSetCondition'

    def icons(self) -> Icons:
        """The set's settings, read from the file the first time they are wanted."""
        return state_of(self.sheet, self.rule)

    def touched(self) -> None:
        self.icons().dirty = True
        super().touched()

    @member
    def Type(self) -> object:
        return VBAInt(6, 'Long')

    @member
    def IconSet(self) -> object:
        icons = self.icons()
        return IconSet(self.sheet.book, -1 if icons.custom else icons.set_id, len(icons.criteria))

    @setter('IconSet')
    def set_icon_set(self, value: object) -> None:
        guard(self.sheet)
        if not isinstance(value, IconSet) or value.identifier not in SETS:
            # Anything but one of the sets leaves the rule as it is, without an error (native probe).
            return
        icons = self.icons()
        old, count = icons.criteria, len(SETS[value.identifier][1])
        if all(one.kind == 'percent' and one.value == default for one, default in zip(old, _defaults(len(old)),
                                                                                      strict=True)):
            new = [Criterion('percent', default, old[place].greater if place < len(old) else False)
                   for place, default in enumerate(_defaults(count))]
        else:
            new = [Criterion(one.kind, one.value, one.greater) for one in old[:count]]
            new += [Criterion(old[-1].kind, old[-1].value) for _ in range(count - len(new))]
        icons.set_id, icons.criteria, icons.custom = value.identifier, new, False
        self.touched()

    @ref_setter('IconSet')
    def set_icon_set_by_reference(self, value: object) -> None:
        # The property takes its set by Let alone: Set f.IconSet = ... is error 438 (native probe).
        raise error(438)

    @member
    def IconCriteria(self) -> object:
        return IconCriteria(self)

    @member
    def ReverseOrder(self) -> object:
        return self.icons().reverse

    @setter('ReverseOrder')
    def set_reverse_order(self, value: object) -> None:
        guard(self.sheet)
        icons = self.icons()
        if icons.custom:
            raise error(5)
        icons.reverse = to_bool(value)
        self.touched()

    @member
    def ShowIconOnly(self) -> object:
        return not self.icons().show_value

    @setter('ShowIconOnly')
    def set_show_icon_only(self, value: object) -> None:
        guard(self.sheet)
        self.icons().show_value = not to_bool(value)
        self.touched()

    @member
    def PercentileValues(self) -> object:
        return all(criterion.kind == 'percentile' for criterion in self.icons().criteria)

    @setter('PercentileValues')
    def set_percentile_values(self, value: object) -> None:
        guard(self.sheet)
        kind = 'percentile' if to_bool(value) else 'num'
        for criterion in self.icons().criteria:
            criterion.kind = kind
        self.touched()


class IconCriteria(VBACollection, ExcelObject):
    vba_type_name = 'IconCriteria'

    def __init__(self, owner: IconSetCondition) -> None:
        self.owner = owner

    def vba_items(self) -> list[object]:
        return [IconCriterion(self.owner, index) for index in range(1, len(self.owner.icons().criteria) + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(9)
        return items[position - 1]


class IconCriterion(ExcelObject):
    vba_type_name = 'IconCriterion'

    def __init__(self, owner: IconSetCondition, index: int) -> None:
        self.owner, self.index = owner, index

    def _criterion(self) -> Criterion:
        return self.owner.icons().criteria[self.index - 1]

    @member
    def Index(self) -> object:
        return VBAInt(self.index, 'Long')

    @member
    def Type(self) -> object:
        return VBAInt(_NUMBERS[self._criterion().kind], 'Long')

    @setter('Type')
    def set_type(self, value: object) -> None:
        guard(self.owner.sheet)
        wanted = int(to_integer(value, 'Long'))
        if self.index == 1 or wanted not in _KINDS:
            raise error(1004)
        criterion, kind = self._criterion(), _KINDS[wanted]
        if kind != criterion.kind:
            if kind == 'num':
                criterion.value = 0.0
            elif kind in ('percent', 'percentile'):
                criterion.value = _defaults(len(self.owner.icons().criteria))[self.index - 1]
            criterion.kind = kind
            self.owner.touched()

    @member
    def Value(self) -> object:
        return self._criterion().read()

    @setter('Value')
    def set_value(self, value: object) -> None:
        guard(self.owner.sheet)
        if self.index == 1:
            raise error(1004)
        criterion = self._criterion()
        criterion.value = given_text(value, self.owner.sheet.name) if isinstance(value, str) else float(to_number(value))
        self.owner.touched()

    @member
    def Operator(self) -> object:
        return VBAInt(_GREATER if self._criterion().greater else _AT_LEAST, 'Long')

    @setter('Operator')
    def set_operator(self, value: object) -> None:
        guard(self.owner.sheet)
        wanted = int(to_integer(value, 'Long'))
        if wanted not in (_GREATER, _AT_LEAST):
            raise error(5)
        self._criterion().greater = wanted == _GREATER
        self.owner.touched()

    @member
    def Icon(self) -> object:
        return VBAInt(self.owner.icons().icons()[self.index - 1], 'Long')

    @setter('Icon')
    def set_icon(self, value: object) -> None:
        guard(self.owner.sheet)
        wanted = int(to_integer(value, 'Long'))
        if wanted not in _SOURCES:
            raise error(1004)
        icons = self.owner.icons()
        if not icons.custom:
            for criterion, icon in zip(icons.criteria, SETS[icons.set_id][1], strict=True):
                criterion.icon = icon
            icons.custom, icons.reverse = True, False
        self._criterion().icon = wanted
        self.owner.touched()


# --- Workbook.IconSets ---------------------------------------------------------------------------------------


class IconSets(VBACollection, ExcelObject):
    vba_type_name = 'IconSets'

    def __init__(self, book: Workbook) -> None:
        self.book = book

    def vba_items(self) -> list[object]:
        return [IconSet(self.book, number, len(icons)) for number, (_, icons) in SETS.items()]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(9)
        return items[position - 1]

    @member
    def Parent(self) -> object:
        return self.book


class IconSet(VBACollection, ExcelObject):
    """One of Excel's sets of icons, or a custom set's (ID -1); a property takes it by Let."""

    vba_type_name = 'IconSet'
    vba_let_as_object = True

    def __init__(self, book: Workbook, identifier: int, count: int) -> None:
        self.book, self.identifier, self.count = book, identifier, count

    def vba_items(self) -> list[object]:
        return [Icon(self, index) for index in range(1, self.count + 1)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if not 1 <= position <= len(items):
            raise error(9)
        return items[position - 1]

    @member
    def ID(self) -> object:
        return VBAInt(self.identifier, 'Long')

    @member
    def Parent(self) -> object:
        return self.book


class Icon(ExcelObject):
    vba_type_name = 'Icon'

    def __init__(self, owner: IconSet, index: int) -> None:
        self.owner, self.index = owner, index

    @member
    def Index(self) -> object:
        return VBAInt(self.index, 'Long')

    @member
    def Parent(self) -> object:
        return self.owner
