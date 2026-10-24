"""A pivot table's label, date, value and top 10 filters, as Excel keeps and applies them.

PivotField.PivotFilters.Add puts a filter on a field. A label filter tests
each item's caption as a worksheet AutoFilter's criteria test a cell, any
case alike and * and ? wild; a date filter tests a date field's items
against a date, two dates or a period, a period counted from today keeping
the days from its first to the one after its last as they were when it
was added, and a quarter or a month keeping any year's; a value filter
tests each item's total of a data field, among the items inside the same
outer items; a top 10 filter keeps the items with the highest or lowest
totals, by count, ties at the last place kept, or until the running total
reaches a percent of the whole or a sum. The table keeps each filter in
its filters element as the autoFilter that spells the test, a label or
date filter ahead of a value filter on its field, and marks a field with
a value filter measureFilter.

Measured with scripts/measure_pivotfilters.py (tests/fixtures/pivotfilters/)
and scripts/measure_pivot_date_filters.py (tests/fixtures/pivot_date_filters/):
a field holds one label or date filter and one value filter, a second of
either error 1004, with its hidden items besides; a date filter takes a
date field only and a label filter any other, error 1004; label and date
filters and hidden items leave their records out first, then value
filters, outer fields first. A date filter compares a moment as it is, a
time of day included, unless Add2's WholeDayFilter has it compare days. A
field on no axis keeps its filters and filters nothing; a page field and
the Values field take none, error 1004. A type Add does not know, a value
filter without a data field or with text for a value, a date filter with
no date, and a between with one value are error 5; a top count below 1 is
E_INVALIDARG.
"""
from __future__ import annotations

import datetime as dt
import math
from collections.abc import Callable
from copy import deepcopy
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._autofilter import PERIODS, Criterion, Seen, file_number, made_criterion, period_bounds
from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._model import ExcelObject
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivot_filters import mark_changed, orientation, rerender, shared, unused
from pyopenvba.apps.excel._pivots import NS, PivotField, decode_text
from pyopenvba.apps.excel._xstring import encode_attribute as encoded
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._objects import VBACollection, member, method
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBADate,
    VBAInt,
    error,
    to_bool,
    to_date,
    to_integer,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivot_calculated import Measures
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
#: Date filters by dates, the same way over the dates' serials.
DATES: dict[int, tuple[str, tuple[tuple[str, str], ...]]] = {
    29: ('dateEqual', (('=', '{0}'),)), 30: ('dateNotEqual', (('<>', '{0}'),)),
    31: ('dateOlderThan', (('<', '{0}'),)), 32: ('dateOlderThanOrEqual', (('<=', '{0}'),)),
    33: ('dateNewerThan', (('>', '{0}'),)), 34: ('dateNewerThanOrEqual', (('>=', '{0}'),)),
    35: ('dateBetween', (('>=', '{0}'), ('<=', '{1}'))), 36: ('dateNotBetween', (('<', '{0}'), ('>', '{1}'))),
}
#: Date filters by a period: the worksheet AutoFilter's dynamic filter each is, which names it in the file and
#: gives the days it keeps: periods counted from today, then each quarter and month of any year.
PERIOD_FILTERS: dict[int, int] = {37: 3, 38: 1, 39: 2, 40: 6, 41: 4, 42: 5, 43: 9, 44: 7, 45: 8, 46: 12, 47: 10,
                                  48: 11, 49: 15, 50: 13, 51: 14, 52: 16, **{53 + step: 17 + step for step in range(16)}}
#: Value filters, the same way as label filters.
VALUES: dict[int, tuple[str, tuple[tuple[str, str], ...]]] = {
    7: ('valueEqual', (('=', '{0}'),)), 8: ('valueNotEqual', (('<>', '{0}'),)),
    9: ('valueGreaterThan', (('>', '{0}'),)), 10: ('valueGreaterThanOrEqual', (('>=', '{0}'),)),
    11: ('valueLessThan', (('<', '{0}'),)), 12: ('valueLessThanOrEqual', (('<=', '{0}'),)),
    13: ('valueBetween', (('>=', '{0}'), ('<=', '{1}'))), 14: ('valueNotBetween', (('<', '{0}'), ('>', '{1}'))),
}
#: Top 10 filters: their type in the file and whether they keep the top.
TOPS: dict[int, tuple[str, bool]] = {1: ('count', True), 2: ('count', False), 3: ('percent', True),
                                     4: ('percent', False), 5: ('sum', True), 6: ('sum', False)}
BETWEEN = (13, 14, 27, 28, 35, 36)
OPERATOR_NAMES = {'<>': 'notEqual', '>': 'greaterThan', '>=': 'greaterThanOrEqual', '<': 'lessThan',
                  '<=': 'lessThanOrEqual'}
#: E_INVALIDARG, which Add raises for a top count below 1.
INVALID_ARGUMENT = -2147024809
#: The Excel 2013 extension a date filter comparing whole days keeps that in.
WHOLE_DAY = '{0605FD5F-26C8-4aeb-8148-2DB25E43C511}'
X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
EPOCH = dt.date(1899, 12, 30)


def dated(kind: int) -> bool:
    return kind in DATES or kind in PERIOD_FILTERS


def label_class(kind: int) -> bool:
    """Whether a filter takes a field's label filter's place: a label filter or a date filter."""
    return kind in LABELS or dated(kind)


@dataclass(frozen=True)
class Rule:
    """A filter as the table keeps it: its element, its field, its type as VBA numbers it, the data field a value
    or top 10 filter totals, its values (a period's first serial and the one after its last), and whether a date
    filter compares whole days."""

    element: XmlElement
    field: int
    kind: int
    measure: int | None
    values: tuple[str | float, ...]
    whole_day: bool = False


def _filters(root: XmlElement) -> XmlElement | None:
    return root.find(NS + 'filters')


def _number(text: str) -> float:
    return float(text)


def _whole_day(element: XmlElement) -> bool:
    for extension in element.findall(f'{NS}extLst/{NS}ext'):
        found = extension.find(f'{{{X15}}}pivotFilter')
        if extension.get('uri') == WHOLE_DAY and found is not None:
            return found.get('useWholeDay') in ('1', 'true')
    return False


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
    if column is not None:
        for kind, (name, _criteria) in DATES.items():
            if name == file_type:
                given = tuple(_number(one.get('val', '0')) for one in column.iter(NS + 'customFilter'))
                if len(given) != (2 if kind in BETWEEN else 1):
                    return None
                return Rule(element, field, kind, None, given, _whole_day(element))
        dynamic = column.find(NS + 'dynamicFilter')
        for kind, which in PERIOD_FILTERS.items():
            if PERIODS[which - 1] == file_type and dynamic is not None:
                bounds = tuple(_number(dynamic.get(key, '0')) for key in ('val', 'maxVal') if dynamic.get(key))
                if len(bounds) != (2 if which <= 16 else 0):
                    return None
                return Rule(element, field, kind, None, bounds, _whole_day(element))
    measure = int(element.get('iMeasureFld', '-1'))
    if column is None or measure < 0:
        return None
    for kind, (name, _criteria) in VALUES.items():
        if name == file_type:
            given = tuple(_number(one.get('val', '0')) for one in column.iter(NS + 'customFilter'))
            if len(given) != (2 if kind in BETWEEN else 1):
                return None
            return Rule(element, field, kind, measure, given)
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


def _criteria_node(spec: tuple[tuple[str, str], ...], texts: list[str], both: bool) -> XmlElement:
    criteria = node('customFilters', {'and': '1'} if both else None)
    for (operator, _pattern), text in zip(spec, texts):
        criteria.append(_custom(operator, text))
    return criteria


def rule_node(field: int, kind: int, identifier: int, measure: int | None, values: tuple[str | float, ...],
              numeric: bool, whole_day: bool = False) -> XmlElement:
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
            criteria = _criteria_node(spec, texts, kind == 27)
    elif kind in DATES:
        name, spec = DATES[kind]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier)})
        criteria = _criteria_node(spec, [_spelled(one) for one in values], kind == 35)
    elif kind in PERIOD_FILTERS:
        name = PERIODS[PERIOD_FILTERS[kind] - 1]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier)})
        dynamic = {'type': name}
        if values:
            dynamic.update({'val': _spelled(values[0]), 'maxVal': _spelled(values[1])})
        criteria = node('dynamicFilter', dynamic)
    elif kind in VALUES:
        name, spec = VALUES[kind]
        attributes.update({'type': name, 'evalOrder': '-1', 'id': str(identifier), 'iMeasureFld': str(measure)})
        criteria = _criteria_node(spec, [_spelled(one) for one in values], kind == 13)
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
    if whole_day:
        extensions = node('extLst')
        extension = node('ext', {'uri': WHOLE_DAY, 'xmlns:x15': X15})
        extension.append(XmlElement(f'{{{X15}}}pivotFilter', {'useWholeDay': '1'}))
        extensions.append(extension)
        element.append(extensions)
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


def _date_keeps(rule: Rule, entry: XmlElement) -> bool:
    """Whether a date filter keeps an item: a date compared as the moment it is, or its day where the filter
    compares whole days; a date in a period, or in a quarter or a month of any year. Anything else it leaves out."""
    from pyopenvba.apps.excel._pivot_layout import scalar

    value = scalar(entry)
    if entry.tag != NS + 'd' or not isinstance(value, float):
        return False
    if rule.kind in DATES:
        moment = math.floor(value) if rule.whole_day else value
        first, second = (float(one) for one in (rule.values * 2)[:2])
        if rule.whole_day:
            first, second = math.floor(first), math.floor(second)
        return {29: moment == first, 30: moment != first, 31: moment < first, 32: moment <= first,
                33: moment > first, 34: moment >= first, 35: first <= moment <= second,
                36: moment < first or moment > second}[rule.kind]
    which = PERIOD_FILTERS[rule.kind]
    if which <= 16:
        return float(rule.values[0]) <= value < float(rule.values[1])
    month = VBADate(value).to_datetime().month
    return (month - 1) // 3 + 1 == which - 16 if which <= 20 else month == which - 20


def label_hidden(found: list[Rule], fields: list[XmlElement], axes: set[int]) -> dict[int, set[int]]:
    """The places in the cache of the items each label or date filter on an axis leaves out."""
    hidden: dict[int, set[int]] = {}
    for rule in found:
        if not label_class(rule.kind) or rule.field not in axes or not 0 <= rule.field < len(fields):
            continue
        keeps: Callable[[XmlElement], bool]
        if dated(rule.kind):
            keeps = lambda entry, rule=rule: _date_keeps(rule, entry)  # noqa: E731
        else:
            criteria = _criteria(rule.kind, rule.values)
            keeps = ((lambda entry, criteria=criteria: any(one.matches(_seen(entry)) for one in criteria))  # noqa: E731
                     if rule.kind == 28 else
                     (lambda entry, criteria=criteria: all(one.matches(_seen(entry)) for one in criteria)))
        for place, entry in enumerate(shared(fields[rule.field])):
            if not keeps(entry):
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
                   axes: tuple[list[int], list[int]], reports: list[XmlElement], measures: Measures) -> list[int]:
    """The records left once each value and top 10 filter on an axis keeps its items' totals, outer fields first,
    each field's items compared among those inside the same outer items, a calculated field's totals worked out by
    ``measures`` (tests/fixtures/pivot_calculated_fields/ value_filter)."""
    by_field = {rule.field: rule for rule in found if not label_class(rule.kind)}
    for axis in axes:
        fields = [index for index in axis if index >= 0]
        for depth, field in enumerate(fields):
            rule = by_field.get(field)
            if rule is None or rule.measure is None or not 0 <= rule.measure < len(reports):
                continue
            report = reports[rule.measure]
            groups: dict[tuple[int, ...], dict[int, list[int]]] = {}
            for row in selected:
                outer = tuple(ids[row][index] for index in fields[:depth])
                groups.setdefault(outer, {}).setdefault(ids[row][field], []).append(row)
            keep: set[tuple[tuple[int, ...], int]] = set()
            for outer, items in groups.items():
                totals: list[tuple[int, float]] = []
                for place, rows in items.items():
                    total = measures.value(report, [records[row] for row in rows])
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


def _day_text(serial: float) -> str:
    """A period's bound as VBA reads it back: the day, spelled yyyy-mm-dd."""
    return (EPOCH + dt.timedelta(days=math.floor(serial))).isoformat()


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
            MemberPropertyField: object = MISSING) -> object:
        """Add a filter; Add takes no WholeDayFilter, error 448 (tests/fixtures/pivot_date_filters/ timed_whole)."""
        if any(one is not MISSING for one in (Order, Name, Description, MemberPropertyField)):
            raise VBAUnsupportedError("PivotFilters.Add's Order, Name, Description and MemberPropertyField are not "
                                      'implemented')
        return add(self.field, Type, DataField, Value1, Value2, False)

    @method
    def Add2(self, Type: object, DataField: object = MISSING, Value1: object = MISSING, Value2: object = MISSING,
             Order: object = MISSING, Name: object = MISSING, Description: object = MISSING,
             MemberPropertyField: object = MISSING, WholeDayFilter: object = MISSING) -> object:
        """Add, by the type library's name for it since Excel 2013, which also takes WholeDayFilter."""
        if any(one is not MISSING for one in (Order, Name, Description, MemberPropertyField)):
            raise VBAUnsupportedError("PivotFilters.Add2's Order, Name, Description and MemberPropertyField are not "
                                      'implemented')
        whole_day = to_bool(WholeDayFilter) if WholeDayFilter is not MISSING else False
        return add(self.field, Type, DataField, Value1, Value2, whole_day)


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

    def _value(self, place: int) -> object:
        """Value1 or Value2: a date filter's Date, a period's day as text, nothing for a quarter or a month; a value
        filter's number; a label filter's text, Empty where it has no second (tests/fixtures/pivot_date_filters/)."""
        rule = self.rule()
        if place < len(rule.values):
            if rule.kind in DATES:
                return VBADate(float(rule.values[place]))
            if rule.kind in PERIOD_FILTERS:
                return _day_text(float(rule.values[place]))
            return _readback(rule.values[place])
        return NULL if rule.kind in VALUES or rule.kind in TOPS or rule.kind in DATES else EMPTY

    @member
    def FilterType(self) -> object:
        return VBAInt(self.rule().kind, 'Long')

    @member
    def Value1(self) -> object:
        return self._value(0)

    @member
    def Value2(self) -> object:
        return self._value(1)

    @member
    def WholeDayFilter(self) -> object:
        rule = self.rule()
        if not dated(rule.kind):
            raise error(1004)
        return rule.whole_day

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


def _date_field(table: PivotTable, index: int) -> bool:
    """Whether a field's values, blanks aside, are all dates, which a date filter needs and a label filter refuses."""
    fields = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    entries = [entry for entry in shared(fields[index]) if entry.tag != NS + 'm' and not unused(entry)] \
        if 0 <= index < len(fields) else []
    return bool(entries) and all(entry.tag == NS + 'd' for entry in entries)


def _tidy(root: XmlElement) -> None:
    """Number the table's filters, drop an empty list, and mark each field holding a value filter."""
    group = _filters(root)
    if group is not None:
        if len(group):
            group.set('count', str(len(group)))
        else:
            root.remove(group)
    measured = {rule.field for rule in rules(root) if not label_class(rule.kind)}
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


def _serial(value: object) -> float:
    """A date filter's date as its serial: a Date, a number or text CDate reads."""
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return to_date(value).serial


def add(field: PivotField, kind_argument: object, data_argument: object, first: object, second: object,
        whole_day: bool) -> object:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's PivotFilters are not implemented")
    if orientation(field.root) == 3:
        raise error(1004)
    kind = int(to_integer(kind_argument, 'Long'))
    if kind not in LABELS and not dated(kind) and kind not in VALUES and kind not in TOPS:
        raise error(5)
    table = field.table.entry
    if (kind in LABELS) == _date_field(table, field.index) and label_class(kind):
        raise error(1004)
    measure: int | None = None
    if not label_class(kind):
        if data_argument is MISSING:
            raise error(5)
        measure = _data_index(table, data_argument)
        if measure < 0:
            raise VBAUnsupportedError('A value filter by a field that is not a data field is not implemented')
    if kind not in PERIOD_FILTERS and (first is MISSING or (kind in BETWEEN and second is MISSING)):
        raise error(5)
    given = (first, second) if kind in BETWEEN else (first,)
    numbers = [float(one) for one in given if isinstance(one, (int, float)) and not isinstance(one, bool)]
    numeric = False
    values: tuple[str | float, ...]
    if kind in LABELS:
        numeric = len(numbers) == len(given)
        values = tuple(number_text(number, formula=True) for number in numbers) if numeric else tuple(
            to_text(one) for one in given)
    elif kind in DATES:
        values = tuple(_serial(one) for one in given)
    elif kind in PERIOD_FILTERS:
        bounds = period_bounds(PERIOD_FILTERS[kind], table.sheet.book.calculator.now().date())
        values = () if bounds is None else tuple(float((day - EPOCH).days) for day in bounds)
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
    if any(label_class(rule.kind) == label_class(kind) for rule in present):
        raise error(1004)
    root = _working(table, field.index)
    group = _place_filters(root)
    identifier = max([0, *(int(entry.get('id', '0')) for entry in group)]) + 1
    element = rule_node(field.index, kind, identifier, measure, values, numeric, whole_day and dated(kind))
    # A label or date filter goes ahead of a value filter on its field.
    ahead = next((spot for spot, entry in enumerate(group) if label_class(kind)
                  and int(entry.get('fld', '-1')) == field.index), None)
    if ahead is None:
        group.append(element)
    else:
        group.insert(ahead, element)
    _tidy(root)
    _apply(table, root)
    return PivotFilter(field, element)


def clear(root: XmlElement, index: int | None, *, labels: bool, values: bool) -> bool:
    """Take away the label and date filters, the value and top 10 filters, or both, of one field or every field;
    whether any went."""
    group = _filters(root)
    if group is None:
        return False
    gone: list[XmlElement] = []
    for entry in group:
        rule = read(entry)
        if rule is None or (index is not None and rule.field != index):
            continue
        if (labels and label_class(rule.kind)) or (values and not label_class(rule.kind)):
            gone.append(entry)
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
