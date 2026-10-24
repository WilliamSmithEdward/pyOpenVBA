"""A pivot table's label, value and top 10 filters, as Excel keeps and applies them.

PivotField.PivotFilters.Add puts a filter on a field. A label filter tests
each item's caption as a worksheet AutoFilter's criteria test a cell, any
case alike and * and ? wild; a value filter tests each item's total of a
data field, among the items inside the same outer items; a top 10 filter
keeps the items with the highest or lowest totals, by count, ties at the
last place kept, or until the running total reaches a percent of the whole
or a sum. The table keeps each filter in its filters element as the
autoFilter that spells the test, a label filter ahead of a value filter
on its field, and marks a field with a value filter measureFilter.

Measured with scripts/measure_pivotfilters.py (tests/fixtures/pivotfilters/):
a field holds one label filter and one value filter, a second of either
error 1004, with its hidden items besides; label filters and hidden items
leave their records out first, then value filters, outer fields first. A
field on no axis keeps its filters and filters nothing; a page field and
the Values field take none, error 1004. A type Add does not know, a value
filter without a data field or with text for a value, and a between with
one value are error 5; a top count below 1 is E_INVALIDARG.
"""
from __future__ import annotations

from collections.abc import Callable
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._autofilter import Criterion, Seen, file_number, made_criterion
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivot_filters import mark_changed, orientation, rerender, shared
from pyopenvba.apps.excel._pivots import NS, PivotField, decode_text
from pyopenvba.apps.excel._xstring import encode_attribute as encoded
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import EMPTY, MISSING, NULL, VBAInt, error, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotTable

#: Label filters: their type in the file, and the AutoFilter criteria each spells, the value going in for {0}
#: and the second for {1}.
LABELS: dict[int, tuple[str, tuple[tuple[str, str], ...]]] = {
    15: ('captionEqual', (('=', '{0}'),)), 16: ('captionNotEqual', (('<>', '{0}'),)),
    17: ('captionBeginsWith', (('=', '{0}*'),)), 18: ('captionNotBeginsWith', (('<>', '{0}*'),)),
    19: ('captionEndsWith', (('=', '*{0}'),)), 20: ('captionNotEndsWith', (('<>', '*{0}'),)),
    21: ('captionContains', (('=', '*{0}*'),)), 22: ('captionNotContains', (('<>', '*{0}*'),)),
    23: ('captionGreaterThan', (('>', '{0}'),)), 24: ('captionGreaterThanOrEqual', (('>=', '{0}'),)),
    25: ('captionLessThan', (('<', '{0}'),)), 26: ('captionLessThanOrEqual', (('<=', '{0}'),)),
    27: ('captionBetween', (('>=', '{0}'), ('<=', '{1}'))), 28: ('captionNotBetween', (('<', '{0}'), ('>', '{1}'))),
}
#: Value filters, the same way.
VALUES: dict[int, tuple[str, tuple[tuple[str, str], ...]]] = {
    7: ('valueEqual', (('=', '{0}'),)), 8: ('valueNotEqual', (('<>', '{0}'),)),
    9: ('valueGreaterThan', (('>', '{0}'),)), 10: ('valueGreaterThanOrEqual', (('>=', '{0}'),)),
    11: ('valueLessThan', (('<', '{0}'),)), 12: ('valueLessThanOrEqual', (('<=', '{0}'),)),
    13: ('valueBetween', (('>=', '{0}'), ('<=', '{1}'))), 14: ('valueNotBetween', (('<', '{0}'), ('>', '{1}'))),
}
#: Top 10 filters: their type in the file and whether they keep the top.
TOPS: dict[int, tuple[str, bool]] = {1: ('count', True), 2: ('count', False), 3: ('percent', True),
                                     4: ('percent', False), 5: ('sum', True), 6: ('sum', False)}
BETWEEN = (13, 14, 27, 28)
OPERATOR_NAMES = {'<>': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual', '<': 'lessThan',
                  '<=': 'lessThanOrEqual'}
#: E_INVALIDARG, which Add raises for a top count below 1.
INVALID_ARGUMENT = -2147024809


def label_kind(kind: int) -> bool:
    return kind in LABELS


@dataclass(frozen=True)
class Rule:
    """A filter as the table keeps it: its element, its field, its type as VBA numbers it, the data field a value
    or top 10 filter totals, and its values."""

    element: XmlElement
    field: int
    kind: int
    measure: int | None
    values: tuple[str | float, ...]


def _filters(root: XmlElement) -> XmlElement | None:
    return root.find(NS + 'filters')


def _number(text: str) -> float:
    return float(text)


def read(element: XmlElement) -> Rule | None:
    """A filter element as a rule, or None for a kind of filter the model does not read."""
    file_type = element.get('type', '')
    field = int(element.get('fld', '-1'))
    column = element.find(f'{NS}autoFilter/{NS}filterColumn')
    for kind, (name, _criteria) in LABELS.items():
        if name == file_type:
            values: tuple[str | float, ...] = (decode_text(element.get('stringValue1', '')),)
            if kind in BETWEEN:
                values += (decode_text(element.get('stringValue2', '')),)
            return Rule(element, field, kind, None, values)
    measure = int(element.get('iMeasureFld', '-1'))
    if column is None or measure < 0:
        return None
    for kind, (name, _criteria) in VALUES.items():
        if name == file_type:
            given = [_number(one.get('val', '0')) for one in column.iter(NS + 'customFilter')]
            if len(given) != (2 if kind in BETWEEN else 1):
                return None
            return Rule(element, field, kind, measure, tuple(given))
    top = column.find(NS + 'top10')
    if top is not None:
        for kind, (name, keeps_top) in TOPS.items():
            if name == file_type and (top.get('top', '1') not in ('0', 'false')) == keeps_top:
                return Rule(element, field, kind, measure, (_number(top.get('val', '0')),))
    return None


def rules(root: XmlElement) -> list[Rule]:
    """Every filter of a table the model reads, refusing one it does not."""
    found: list[Rule] = []
    group = _filters(root)
    elements: list[XmlElement] = list(group) if group is not None else []
    for element in elements:
        rule = read(element)
        if rule is None:
            raise VBAUnsupportedError('A pivot table filter of this kind is not implemented')
        found.append(rule)
    return found


def readable(root: XmlElement) -> bool:
    group = _filters(root)
    return group is None or all(read(element) is not None for element in group)


# --- the file's spelling --------------------------------------------------------------------------------------

def _custom(operator: str, text: str) -> XmlElement:
    attributes = {'operator': OPERATOR_NAMES[operator]} if operator != '=' else {}
    attributes['val'] = encoded(text)
    return node('customFilter', attributes)


def _spelled(value: object) -> str:
    """A filter's value as its criteria spell it: text as given, a number as the file writes one."""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return file_number(float(value))
    return str(value)


def rule_node(field: int, kind: int, identifier: int, measure: int | None, values: tuple[str | float, ...],
              numeric: bool) -> XmlElement:
    """A filter element as Excel writes one; ``numeric`` says a label filter was given a number, which it spells as
    a criterion rather than as a value in a list."""
    attributes = {'fld': str(field)}
    criteria: XmlElement
    if kind in LABELS:
        name, spec = LABELS[kind]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier),
                           'stringValue1': encoded(str(values[0]))})
        if kind in BETWEEN:
            attributes['stringValue2'] = encoded(str(values[1]))
        texts = [pattern.format(*[str(one) for one in values]) for _operator, pattern in spec]
        if kind == 15 and not numeric and not any(char in texts[0] for char in '*?'):
            criteria = node('filters')
            criteria.append(node('filter', {'val': encoded(texts[0])}))
        else:
            criteria = node('customFilters', {'and': '1'} if kind == 27 else None)
            for (operator, _pattern), text in zip(spec, texts):
                criteria.append(_custom(operator, text))
    elif kind in VALUES:
        name, spec = VALUES[kind]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier), 'iMeasureFld': str(measure)})
        criteria = node('customFilters', {'and': '1'} if kind == 13 else None)
        for (operator, _pattern), value in zip(spec, values):
            criteria.append(_custom(operator, _spelled(value)))
    else:
        name, keeps_top = TOPS[kind]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier), 'iMeasureFld': str(measure)})
        top: dict[str, str] = {} if keeps_top else {'top': '0'}
        if name == 'percent':
            top['percent'] = '1'
        text = _spelled(values[0])
        top.update({'val': text, 'filterVal': text})
        criteria = node('top10', top)
    element = node('filter', attributes)
    automatic = node('autoFilter', {'ref': 'A1'})
    column = node('filterColumn', {'colId': '0'})
    column.append(criteria)
    automatic.append(column)
    element.append(automatic)
    return element


# --- what the filters keep ------------------------------------------------------------------------------------

def _seen(entry: XmlElement) -> Seen:
    """An item's value as a criterion looks at it: its caption, and the number or text it is."""
    from pyopenvba.apps.excel._pivot_layout import scalar

    value = scalar(entry)
    if value is EMPTY:
        return Seen('(blank)', None, '(blank)', False)
    if isinstance(value, bool):
        text = 'TRUE' if value else 'FALSE'
        return Seen(text, None, text, False)
    if isinstance(value, float):
        if entry.tag == NS + 'd':
            raise VBAUnsupportedError('A label filter over dates is not implemented')
        return Seen(number_text(value), value, None, False)
    if entry.tag == NS + 'e':
        return Seen(str(value), None, None, True)
    return Seen(str(value), None, str(value), False)


def _criteria(kind: int, values: tuple[str | float, ...]) -> list[Criterion]:
    _name, spec = LABELS[kind]
    texts = [str(one) for one in values]
    return [made_criterion(operator, pattern.format(*texts)) for operator, pattern in spec]


def label_hidden(found: list[Rule], fields: list[XmlElement], axes: set[int]) -> dict[int, set[int]]:
    """The places in the cache of the items each label filter on an axis leaves out."""
    hidden: dict[int, set[int]] = {}
    for rule in found:
        if rule.kind not in LABELS or rule.field not in axes or not 0 <= rule.field < len(fields):
            continue
        criteria = _criteria(rule.kind, rule.values)
        keeps: Callable[[Seen], bool] = ((lambda seen: any(one.matches(seen) for one in criteria)) if rule.kind == 28
                                         else (lambda seen: all(one.matches(seen) for one in criteria)))
        for place, entry in enumerate(shared(fields[rule.field])):
            if not keeps(_seen(entry)):
                hidden.setdefault(rule.field, set()).add(place)
    return hidden


def _passes(kind: int, total: float, values: tuple[str | float, ...]) -> bool:
    first = float(values[0])
    second = float(values[1]) if len(values) > 1 else first
    return {7: total == first, 8: total != first, 9: total > first, 10: total >= first, 11: total < first,
            12: total <= first, 13: first <= total <= second, 14: total < first or total > second}[kind]


def _kept(kind: int, totals: list[tuple[int, float]], amount: float) -> set[int]:
    """The items a top 10 filter keeps of those with these totals, each named by its value's place."""
    _name, keeps_top = TOPS[kind]
    ranked = sorted(totals, key=lambda pair: pair[1], reverse=keeps_top)
    if kind in (1, 2):
        count = int(amount)
        if count >= len(ranked):
            return {place for place, _ in ranked}
        last = ranked[count - 1][1]
        return {place for place, total in ranked if (total > last if keeps_top else total < last) or total == last}
    whole = sum(total for _, total in ranked)
    kept: set[int] = set()
    running = 0.0
    for place, total in ranked:
        kept.add(place)
        running += total
        if (running / whole * 100 >= amount if kind in (3, 4) and whole else running >= amount):
            break
    return kept


def value_filtered(selected: list[int], ids: list[list[int]], records: list[list[XmlElement]], found: list[Rule],
                   axes: tuple[list[int], list[int]], reports: list[XmlElement]) -> list[int]:
    """The records left once each value and top 10 filter on an axis keeps its items' totals, outer fields first,
    each field's items compared among those inside the same outer items."""
    from pyopenvba.apps.excel._pivot_aggregate import aggregate

    by_field = {rule.field: rule for rule in found if rule.kind not in LABELS}
    for axis in axes:
        fields = [index for index in axis if index >= 0]
        for depth, field in enumerate(fields):
            rule = by_field.get(field)
            if rule is None or rule.measure is None or not 0 <= rule.measure < len(reports):
                continue
            report = reports[rule.measure]
            source, function = int(report.get('fld', '0')), report.get('subtotal', 'sum')
            groups: dict[tuple[int, ...], dict[int, list[int]]] = {}
            for row in selected:
                outer = tuple(ids[row][index] for index in fields[:depth])
                groups.setdefault(outer, {}).setdefault(ids[row][field], []).append(row)
            keep: set[tuple[tuple[int, ...], int]] = set()
            for outer, items in groups.items():
                totals: list[tuple[int, float]] = []
                for place, rows in items.items():
                    total = aggregate([records[row][source] for row in rows], function)
                    if isinstance(total, (int, float)) and not isinstance(total, bool):
                        totals.append((place, float(total)))
                if rule.kind in TOPS:
                    kept = _kept(rule.kind, totals, float(rule.values[0]))
                    keep.update((outer, place) for place in kept)
                else:
                    keep.update((outer, place) for place, total in totals if _passes(rule.kind, total, rule.values))
            selected = [row for row in selected
                        if (tuple(ids[row][index] for index in fields[:depth]), ids[row][field]) in keep]
    return selected


# --- the object model -----------------------------------------------------------------------------------------

def _readback(value: object) -> object:
    """A filter's value as VBA reads it back: a whole number as a Long, any other number as a Double."""
    if isinstance(value, float) and value.is_integer() and -2 ** 31 <= value < 2 ** 31:
        return VBAInt(int(value), 'Long')
    return value


class PivotFilters(VBACollection, ExcelObject):
    vba_type_name = 'PivotFilters'

    def __init__(self, field: PivotField) -> None:
        self.field = field

    def vba_items(self) -> list[object]:
        table = self.field.table.entry
        return [PivotFilter(self.field, rule.element) for rule in rules(table.root) if rule.field == self.field.index]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, 'Long'))
        if 1 <= position <= len(items):
            return items[position - 1]
        raise error(1004)

    @member
    def Parent(self) -> object:
        return self.field

    @method
    def Add(self, Type: object, DataField: object = MISSING, Value1: object = MISSING, Value2: object = MISSING,
            Order: object = MISSING, Name: object = MISSING, Description: object = MISSING,
            MemberPropertyField: object = MISSING, WholeDayFilter: object = MISSING) -> object:
        if any(one is not MISSING for one in (Order, Name, Description, MemberPropertyField, WholeDayFilter)):
            raise VBAUnsupportedError("PivotFilters.Add's Order, Name, Description, MemberPropertyField and "
                                      'WholeDayFilter are not implemented')
        return add(self.field, Type, DataField, Value1, Value2)

    @method
    def Add2(self, Type: object, DataField: object = MISSING, Value1: object = MISSING, Value2: object = MISSING,
             Order: object = MISSING, Name: object = MISSING, Description: object = MISSING,
             MemberPropertyField: object = MISSING, WholeDayFilter: object = MISSING) -> object:
        """The type library's name for Add since Excel 2013, which takes the same arguments."""
        return self.Add(Type, DataField, Value1, Value2, Order, Name, Description, MemberPropertyField,
                        WholeDayFilter)


class PivotFilter(ExcelObject):
    """A filter of a field, found by its id each time, as a layout drawn again makes the table's filters anew."""

    vba_type_name = 'PivotFilter'

    def __init__(self, field: PivotField, element: XmlElement) -> None:
        self.field, self.identifier = field, element.get('id', '')

    def element(self) -> XmlElement | None:
        group = _filters(self.field.table.entry.root)
        if group is None:
            return None
        return next((entry for entry in group if entry.get('id') == self.identifier), None)

    def rule(self) -> Rule:
        element = self.element()
        rule = read(element) if element is not None else None
        if rule is None:
            raise error(424)
        return rule

    @member
    def FilterType(self) -> object:
        return VBAInt(self.rule().kind, 'Long')

    @member
    def Value1(self) -> object:
        return _readback(self.rule().values[0])

    @member
    def Value2(self) -> object:
        rule = self.rule()
        if len(rule.values) > 1:
            return _readback(rule.values[1])
        return EMPTY if rule.kind in LABELS else NULL

    @member
    def DataField(self) -> object:
        rule = self.rule()
        table = self.field.table.entry
        reports = table.root.findall(f'{NS}dataFields/{NS}dataField')
        if rule.measure is None or not 0 <= rule.measure < len(reports):
            raise error(1004)
        report = reports[rule.measure]
        return PivotField(table, int(report.get('fld', '0')), report, True)

    @member
    def Active(self) -> object:
        self.rule()
        return True

    @member
    def PivotField(self) -> object:
        return self.field

    @member
    def Parent(self) -> object:
        return self.field

    @method
    def Delete(self) -> object:
        self.rule()
        table = self.field.table.entry
        root = _working(table, self.field.index)
        group = _filters(root)
        assert group is not None
        group.remove(next(entry for entry in group if entry.get('id') == self.identifier))
        _tidy(root)
        _apply(table, root)
        return EMPTY


def _data_index(table: PivotTable, value: object) -> int:
    """The place among the table's data fields of the one a value filter's DataField names."""
    reports = table.root.findall(f'{NS}dataFields/{NS}dataField')
    if isinstance(value, PivotField):
        if value.data:
            return next((place for place, report in enumerate(reports) if report is value.root), -1)
        name = str(value.Name()).casefold()
        return next((place for place, report in enumerate(reports)
                     if decode_text(report.get('name', '')).casefold() == name), -1)
    return -1


def _tidy(root: XmlElement) -> None:
    """Number the table's filters, drop an empty list, and mark each field holding a value filter."""
    group = _filters(root)
    if group is not None:
        if len(group):
            group.set('count', str(len(group)))
        else:
            root.remove(group)
    measured = {rule.field for rule in rules(root) if rule.kind not in LABELS}
    from pyopenvba.apps.excel._pivot_filters import field_attributes

    for index, field in enumerate(root.findall(f'{NS}pivotFields/{NS}pivotField')):
        attributes = {key: value for key, value in field.attrib.items() if key != 'measureFilter'}
        if index in measured:
            attributes['measureFilter'] = '1'
        field.attrib.clear()
        field.attrib.update(field_attributes(attributes))


def _laid_out(table: PivotTable, index: int) -> bool:
    fields = table.root.findall(f'{NS}pivotFields/{NS}pivotField')
    return 0 <= index < len(fields) and orientation(fields[index]) != 0


def _working(table: PivotTable, index: int) -> XmlElement:
    """The definition a change of a field's filters is made in: a copy to lay out anew where the field is on an
    axis, else the table's own, which a field on no axis changes in place."""
    return deepcopy(table.root) if _laid_out(table, index) else table.root


def _apply(table: PivotTable, root: XmlElement) -> None:
    if root is table.root:
        mark_changed(table)
    else:
        rerender(table, root)


def _place_filters(root: XmlElement) -> XmlElement:
    """The table's filters element, made after its style where it has none, as the schema orders them."""
    group = _filters(root)
    if group is not None:
        return group
    group = node('filters', {'count': '0'})
    after = [NS + name for name in ('rowHierarchiesUsage', 'colHierarchiesUsage', 'extLst')]
    place = next((spot for spot, child in enumerate(root) if child.tag in after), len(root))
    root.insert(place, group)
    return group


def add(field: PivotField, kind_argument: object, data_argument: object, first: object, second: object) -> object:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's PivotFilters are not implemented")
    if orientation(field.root) == 3:
        raise error(1004)
    kind = int(to_integer(kind_argument, 'Long'))
    if 29 <= kind <= 60:
        raise VBAUnsupportedError('Filtering a pivot field by dates is not implemented')
    if kind not in LABELS and kind not in VALUES and kind not in TOPS:
        raise error(5)
    table = field.table.entry
    measure: int | None = None
    if kind not in LABELS:
        if data_argument is MISSING:
            raise error(5)
        measure = _data_index(table, data_argument)
        if measure < 0:
            raise VBAUnsupportedError('A value filter by a field that is not a data field is not implemented')
    if first is MISSING or (kind in BETWEEN and second is MISSING):
        raise error(5)
    given = (first, second) if kind in BETWEEN else (first,)
    numbers = [float(one) for one in given if isinstance(one, (int, float)) and not isinstance(one, bool)]
    numeric = False
    values: tuple[str | float, ...]
    if kind in LABELS:
        numeric = len(numbers) == len(given)
        values = tuple(number_text(number, formula=True) for number in numbers) if numeric else tuple(
            to_text(one) for one in given)
    else:
        if len(numbers) != len(given):
            raise error(5)
        values = tuple(numbers)
        if kind == 1 or kind == 2:
            if float(values[0]) < 1:
                raise error(INVALID_ARGUMENT)
        elif float(values[0]) <= 0:
            raise VBAUnsupportedError('A top 10 filter by a percent or sum of 0 or less is not implemented')
    present = [rule for rule in rules(table.root) if rule.field == field.index]
    if any((rule.kind in LABELS) == (kind in LABELS) for rule in present):
        raise error(1004)
    root = _working(table, field.index)
    group = _place_filters(root)
    identifier = max([0, *(int(entry.get('id', '0')) for entry in group)]) + 1
    element = rule_node(field.index, kind, identifier, measure, values, numeric)
    # A label filter goes ahead of a value filter on its field.
    ahead = next((spot for spot, entry in enumerate(group) if kind in LABELS
                  and int(entry.get('fld', '-1')) == field.index), None)
    if ahead is None:
        group.append(element)
    else:
        group.insert(ahead, element)
    _tidy(root)
    _apply(table, root)
    return PivotFilter(field, element)


def clear(root: XmlElement, index: int | None, *, labels: bool, values: bool) -> bool:
    """Take away the label filters, the value and top 10 filters, or both, of one field or every field; whether
    any went."""
    group = _filters(root)
    if group is None:
        return False
    gone = [entry for entry in group
            if (index is None or int(entry.get('fld', '-1')) == index)
            and ((labels and (rule := read(entry)) is not None and rule.kind in LABELS)
                 or (values and (rule := read(entry)) is not None and rule.kind not in LABELS))]
    for entry in gone:
        group.remove(entry)
    if gone:
        _tidy(root)
    return bool(gone)


def clear_kind(field: PivotField, *, labels: bool, values: bool) -> None:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's PivotFilters are not implemented")
    table = field.table.entry
    root = _working(table, field.index)
    if clear(root, field.index, labels=labels, values=values):
        _apply(table, root)


def field_filters(field: PivotField) -> PivotFilters:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's PivotFilters are not implemented")
    return PivotFilters(field)
