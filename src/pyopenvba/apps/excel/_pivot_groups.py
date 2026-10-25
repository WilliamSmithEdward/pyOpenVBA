"""Grouping a pivot table's items as Excel does: numbers into ranges, dates by periods, and chosen items into
named groups.

A cache field keeps its grouping in a fieldGroup. Numbers group in place: the field's fieldGroup names the
field itself as its base and holds a rangePr, the start, end and step, and groupItems, which take over from
the field's shared items as the items a table lists, each record's value going to the range it falls in.
Dates and chosen items group into fields of their own after the cache's others, databaseField="0", whose
fieldGroup names the field they group as their base: one field for each period of a date, its rangePr
saying which, and one of chosen items, whose discretePr gives each of the base's shared items its place in
groupItems, a group's name or an item standing alone. The field a grouping comes from names the field that
groups it as its par: a date the coarsest of its periods, a field of items its field of groups, which can
be grouped again into another over the same base.

Measured with scripts/measure_pivot_grouping.py (tests/fixtures/pivot_grouping/):

- Range.Group on one cell of a field of numbers groups it by By from Start to End, True or left out for the
  field's least and greatest values, End one more than Start where those are the same. By left out is a power
  of ten, the largest no more than a third of the span. The ranges go from Start in steps of By, as many as it
  takes to pass End, the last ending on End where a whole number of steps reaches it; a value below Start goes
  to an item of its own ahead of them, ``<Start``, a blank there too, and one past the last range to one
  after them, ``>`` the last range's end. A range shows as ``low-high``, high one less than the next range's
  low where the field holds whole numbers and Start and By are whole, and the next range's low otherwise.
- A field of dates groups into a field for each of its periods Periods chooses, months where it chooses none,
  named for the period and the field: ``Months (Date)``. Its items are every second, minute, hour, day of the
  year, month, quarter or year from Start's to End's, with ``<Start`` and ``>End`` either side, End True
  being the moment after the latest day, or the latest moment where it has a time. Days by a By of more than
  one are ranges from Start instead, the last ending on End, and a value on End is past them; the others take
  a value on End. The fields take the date's place on the axis, the coarsest outermost; another table of the
  cache showing the date keeps it, the periods outside it. Grouping the date again, from any of its periods'
  cells, takes the periods it had away.
- Range.Group on cells of two items or more of a field groups them into ``Group1``, the lowest number neither
  the field of groups' items nor the field's take, in that field, ``Region2`` beside ``Region``, which comes
  outside the field on its axis and lists an item for each group and each item left alone; a group takes the
  items it gathers from groups they were in, a group left with one item becoming that item again. The new
  group comes after the others in groupItems and the field's items sort each at its first; ungrouping one puts
  its items after the others there and in its place in each table's list, and ungrouping the last removes the
  field of groups.
- Text cannot group by ranges, nor can a field with text among its numbers, nor a page field's, a data cell or
  a total's, a By of 0, a Start past End, or Periods choosing nothing or not seven: error 1004.
- A refresh works an automatic Start or End out again, and gives a new item a place alone.
"""
from __future__ import annotations

import datetime as dt
import math
from copy import deepcopy
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAError, VBAUnsupportedError
from pyopenvba.formula._values import number_text
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, VBADate, error, to_bool, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._pivots import PivotCache, PivotField, PivotItem, PivotItems, PivotTable

#: Range.Group's Periods, in the order its array gives them, and the order the fields of a date's periods take.
PERIODS = ('seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years')
MONTHS = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()
#: The order the schema gives a rangePr's attributes.
RANGE_ORDER = 'autoStart autoEnd groupBy startNum endNum startDate endDate groupInterval'.split()
#: The days of a year that has a 29 February, which a date's days go by.
_DAYS = [dt.date(2000, 1, 1) + dt.timedelta(days=offset) for offset in range(366)]
_DAY_PLACE = {(day.month, day.day): place for place, day in enumerate(_DAYS)}


def _flag(element: XmlElement, name: str, default: bool) -> bool:
    value = element.get(name)
    return default if value is None else value in ('1', 'true')


def _entries(parent: XmlElement, tag: str) -> list[XmlElement]:
    """The children of a part of an element, none where it lacks the part."""
    found = parent.find(NS + tag)
    return list(found) if found is not None else []


def fields_of(cache: PivotCache) -> list[XmlElement]:
    return cache.root.findall(f'{NS}cacheFields/{NS}cacheField')


def is_database(cache_field: XmlElement) -> bool:
    """Whether a field's values come from the source, in each record, as all but calculated and grouped ones do."""
    return cache_field.get('databaseField') not in ('0', 'false')


def field_group(cache_field: XmlElement) -> XmlElement | None:
    return cache_field.find(NS + 'fieldGroup')


def ranged(cache_field: XmlElement) -> XmlElement | None:
    """A field's rangePr, where it groups numbers or dates by ranges or periods."""
    group = field_group(cache_field)
    return group.find(NS + 'rangePr') if group is not None else None


def discrete(cache_field: XmlElement) -> XmlElement | None:
    group = field_group(cache_field)
    return group.find(NS + 'discretePr') if group is not None else None


def base_of(cache_field: XmlElement, index: int) -> int:
    """The field a grouped field groups, the field itself where it groups in place or does not group."""
    group = field_group(cache_field)
    if group is None or group.get('base') is None:
        return index
    return int(group.get('base', str(index)))


def parent_of(cache_field: XmlElement) -> int | None:
    group = field_group(cache_field)
    return int(group.get('par', '0')) if group is not None and group.get('par') is not None else None


def is_group_field(cache_field: XmlElement) -> bool:
    """Whether a field is one a grouping made, of a date's periods or of groups of items."""
    return not is_database(cache_field) and field_group(cache_field) is not None


def items_of(cache_field: XmlElement) -> XmlElement | None:
    """The items a field lists: its groups where it is grouped, else its shared items."""
    group = field_group(cache_field)
    listed = group.find(NS + 'groupItems') if group is not None else None
    return listed if listed is not None else cache_field.find(NS + 'sharedItems')


def item_entries(cache_field: XmlElement) -> list[XmlElement]:
    found = items_of(cache_field)
    return list(found) if found is not None else []


def database_count(cache_fields: list[XmlElement]) -> int:
    """How many of a cache's fields its records hold, the source's, which come ahead of every other."""
    count = sum(1 for field in cache_fields if is_database(field))
    if any(not is_database(field) for field in cache_fields[:count]):
        raise VBAUnsupportedError('A pivot cache with a calculated or grouped field ahead of its source fields is '
                                  'not implemented')
    return count


# What a grouping's items are, and the item each value goes to.

def _number(raw: str | None, default: float = 0.0) -> float:
    try:
        return float(raw) if raw is not None else default
    except ValueError:
        return default


def _moment(raw: str | None) -> dt.datetime | None:
    try:
        return dt.datetime.fromisoformat(raw) if raw else None
    except ValueError:
        return None


def number_spelling(value: float) -> str:
    """A number as a grouping's attributes spell it, as a cache spells its numbers: a whole one without a point."""
    from pyopenvba.apps.excel._pivot_cache_file import number_text

    return number_text(value)


def steps(start: float, end: float, interval: float) -> tuple[int, bool]:
    """How many ranges of ``interval`` from ``start`` it takes to pass ``end``, and whether the last ends on it."""
    whole = (end - start) / interval
    nearest = round(whole)
    if abs(whole - nearest) <= 1e-9 * max(1.0, abs(whole)):
        return max(int(nearest), 1), True
    return max(math.ceil(whole), 1), False


def default_interval(span: float) -> float:
    """By left out: the largest power of ten no more than a third of the span (tests/fixtures/pivot_grouping/
    span_*)."""
    digits = Decimal(repr(span if span > 0 else 1.0))
    exponent = digits.adjusted()
    if digits.scaleb(-exponent) < 3:
        exponent -= 1
    return float(Decimal(1).scaleb(exponent))


def number_labels(start: float, end: float, interval: float, whole: bool, blank: bool) -> list[str]:
    """A grouping of numbers' items: ``<start``, the ranges, and the item past them; ``(blank)`` first where the
    field holds blanks (tests/fixtures/pivot_grouping/ numbers_*, value_*, mixed_gappy)."""
    count, closed = steps(start, end, interval)
    whole = whole and start.is_integer() and interval.is_integer()
    labels = ['(blank)' if blank else '<' + number_text(start)]
    for step in range(count):
        low = start + step * interval
        high = end if closed and step == count - 1 else low + interval - (1 if whole else 0)
        labels.append(f'{number_text(low)}-{number_text(high)}')
    labels.append('>' + number_text(start + count * interval))
    return labels


def _day(moment: dt.datetime | dt.date) -> str:
    return f'{moment.month}/{moment.day}/{moment.year}'


def day_ranges(start: dt.datetime, end: dt.datetime, interval: int) -> int:
    """How many ranges of ``interval`` days from ``start`` come before ``end``."""
    return max(math.ceil((end - start) / dt.timedelta(days=interval)), 1)


def date_labels(period: str, start: dt.datetime, end: dt.datetime, interval: int, blank: bool) -> list[str]:
    """A period's items, between ``<start`` and ``>end`` in m/d/yyyy whatever the field's format (tests/fixtures/
    pivot_grouping/ dates_*, times_*, mixed_iso_weeks)."""
    middle: list[str]
    if period in ('seconds', 'minutes'):
        middle = [f':{value:02d}' for value in range(60)]
    elif period == 'hours':
        middle = [f'{hour % 12 or 12} {"AM" if hour < 12 else "PM"}' for hour in range(24)]
    elif period == 'days' and interval > 1:
        middle = []
        for step in range(day_ranges(start, end, interval)):
            low = (start + dt.timedelta(days=step * interval)).date()
            high = min(low + dt.timedelta(days=interval - 1), end.date())
            middle.append(f'{_day(low)} - {_day(high)}')
    elif period == 'days':
        middle = [f'{day.day}-{MONTHS[day.month - 1]}' for day in _DAYS]
    elif period == 'months':
        middle = list(MONTHS)
    elif period == 'quarters':
        middle = [f'Qtr{quarter}' for quarter in range(1, 5)]
    elif period == 'years':
        middle = [str(year) for year in range(start.year, end.year + 1)]
    else:
        raise VBAUnsupportedError(f'Grouping dates by {period} is not implemented')
    return ['(blank)' if blank else '<' + _day(start), *middle, '>' + _day(end)]


class Grouping:
    """What a rangePr groups by, read from its attributes: numbers from a start by a step, or dates by a period."""

    def __init__(self, element: XmlElement) -> None:
        self.period = element.get('groupBy', 'range')
        self.interval = _number(element.get('groupInterval'), 1.0)
        self.auto_start = _flag(element, 'autoStart', True)
        self.auto_end = _flag(element, 'autoEnd', True)
        self.start = _number(element.get('startNum'))
        self.end = _number(element.get('endNum'))
        moment = _moment(element.get('startDate'))
        self.start_date = moment if moment is not None else dt.datetime(1900, 1, 1)
        moment = _moment(element.get('endDate'))
        self.end_date = moment if moment is not None else self.start_date
        if self.period == 'range':
            if self.interval <= 0:
                raise VBAUnsupportedError('A pivot grouping without a step is not implemented')
            self.count, self.closed = steps(self.start, self.end, self.interval)
        elif self.period == 'days' and self.interval > 1:
            self.count, self.closed = day_ranges(self.start_date, self.end_date, int(self.interval)), False
        else:
            self.count, self.closed = 0, False

    def labels(self, whole: bool, blank: bool) -> list[str]:
        if self.period == 'range':
            return number_labels(self.start, self.end, self.interval, whole, blank)
        return date_labels(self.period, self.start_date, self.end_date, int(self.interval), blank)

    def index(self, entry: XmlElement, length: int) -> int:
        """The item a value goes to, of ``length``: a blank to the first (tests/fixtures/pivot_grouping/
        mixed_gappy)."""
        kind = entry.tag[len(NS):]
        if self.period == 'range':
            if kind != 'n':
                return 0
            value = float(entry.get('v', '0'))
            if value < self.start:
                return 0
            step = math.floor((value - self.start) / self.interval)
            if step < self.count:
                return step + 1
            if self.closed and step == self.count and value <= self.end:
                return self.count
            return self.count + 1
        moment = _moment(entry.get('v')) if kind == 'd' else None
        if moment is None:
            return 0
        if moment < self.start_date:
            return 0
        if self.period == 'days' and self.interval > 1:
            step = math.floor((moment - self.start_date) / dt.timedelta(days=self.interval))
            return step + 1 if step < self.count else self.count + 1
        if moment > self.end_date:
            return length - 1
        if self.period == 'seconds':
            return moment.second + 1
        if self.period == 'minutes':
            return moment.minute + 1
        if self.period == 'hours':
            return moment.hour + 1
        if self.period == 'days':
            return _DAY_PLACE[moment.month, moment.day] + 1
        if self.period == 'months':
            return moment.month
        if self.period == 'quarters':
            return (moment.month - 1) // 3 + 1
        return moment.year - self.start_date.year + 1


def whole_numbers(entries: list[XmlElement]) -> bool:
    return all(float(entry.get('v', '0')).is_integer() for entry in entries if entry.tag == NS + 'n')


def blanks(entries: list[XmlElement]) -> bool:
    return any(entry.tag == NS + 'm' for entry in entries)


def automatic_dates(entries: list[XmlElement]) -> tuple[dt.datetime, dt.datetime] | None:
    """A field of dates' least and greatest, as an automatic Start and End take them: the greatest the moment
    after its day where it has no time (tests/fixtures/pivot_grouping/ mixed_whole, mixed_timed)."""
    moments = [moment for entry in entries if entry.tag == NS + 'd' and (moment := _moment(entry.get('v')))]
    if not moments:
        return None
    latest = max(moments)
    if latest == dt.datetime.combine(latest.date(), dt.time()):
        latest += dt.timedelta(days=1)
    return min(moments), latest


def _set_range(element: XmlElement, values: dict[str, str]) -> None:
    current = {**element.attrib, **values}
    element.attrib.clear()
    element.attrib.update({name: current[name] for name in RANGE_ORDER if name in current})


def _rewrite_items(group: XmlElement, labels: list[str]) -> None:
    listed = group.find(NS + 'groupItems')
    if listed is None:
        listed = node('groupItems')
        group.append(listed)
    listed[:] = [node('s', {'v': label}) for label in labels]
    listed.set('count', str(len(labels)))


def refresh_ranges(fields: list[XmlElement]) -> None:
    """Work out an automatic Start and End from the values a cache holds now, and the items again where they
    moved, as a refresh does (tests/fixtures/pivot_grouping/ refresh_numbers_auto, refresh_dates)."""
    for index, field in enumerate(fields):
        element = ranged(field)
        group = field_group(field)
        if element is None or group is None:
            continue
        base = base_of(field, index)
        if not 0 <= base < len(fields):
            continue
        values = _entries(fields[base], 'sharedItems')
        grouping = Grouping(element)
        changed: dict[str, str] = {}
        if grouping.period == 'range':
            numbers = [float(entry.get('v', '0')) for entry in values if entry.tag == NS + 'n']
            if not numbers:
                continue
            low, high = min(numbers), max(numbers)
            if high == low:
                high = low + 1
            if grouping.auto_start and low != grouping.start:
                changed['startNum'] = number_spelling(low)
            if grouping.auto_end and high != grouping.end:
                changed['endNum'] = number_spelling(high)
        else:
            found = automatic_dates(values)
            if found is None:
                continue
            if grouping.auto_start and found[0] != grouping.start_date:
                changed['startDate'] = found[0].isoformat()
            if grouping.auto_end and found[1] != grouping.end_date:
                changed['endDate'] = found[1].isoformat()
        if changed:
            _set_range(element, changed)
            _rewrite_items(group, Grouping(element).labels(whole_numbers(values), blanks(values)))


def layout_groups(fields: list[XmlElement], sources: int, records: list[list[XmlElement]], ids: list[list[int]],
                  positions: list[list[int]]) -> None:
    """Give each record its item in every grouped field, and each grouped field the order a table first lists
    its items in: ranges and periods as their groupItems list them, groups of items each at its first item in
    its base's order (tests/fixtures/pivot_grouping/ manual_late_numbers, manual_late_text)."""
    shared_ids = [list(row) for row in ids]
    for row in ids:
        row.extend(0 for _ in range(len(fields) - len(row)))
    while len(positions) < len(fields):
        positions.append([])
    for index, field in enumerate(fields):
        group = field_group(field)
        if group is None or group.find(NS + 'groupItems') is None:
            continue
        base = base_of(field, index)
        if not 0 <= base < sources:
            raise VBAUnsupportedError('A pivot field grouping a field the records do not hold is not implemented')
        length = len(_entries(group, 'groupItems'))
        element = group.find(NS + 'rangePr')
        mapping = group.find(NS + 'discretePr')
        if element is not None:
            grouping = Grouping(element)
            for record, row in zip(records, ids, strict=True):
                row[index] = min(grouping.index(record[base], length), max(length - 1, 0))
            positions[index] = list(range(length))
        elif mapping is not None:
            places = [int(entry.get('v', '0')) for entry in mapping]
            for shared, row in zip(shared_ids, ids, strict=True):
                spot = shared[base]
                row[index] = places[spot] if 0 <= spot < len(places) else 0
            positions[index] = _by_first(places, positions[base], length)


def _by_first(places: list[int], base_order: list[int], length: int) -> list[int]:
    """A field of groups' items, each where its first item comes in its base's order, given each of the base's
    items' place among them."""
    rank = {place: rank for rank, place in enumerate(base_order)}
    first: dict[int, int] = {}
    for spot, place in enumerate(places):
        first[place] = min(first.get(place, len(rank)), rank.get(spot, len(rank)))
    return sorted(range(length), key=lambda place: (first.get(place, len(rank)), place))


def natural_places(cache_fields: list[XmlElement], index: int) -> list[int]:
    """The places of a field's items in the order a table first lists them: a grouping's as layout_groups orders
    them, any other field's sorted, unused ones left out."""
    from pyopenvba.apps.excel._pivot_filters import sorted_places
    from pyopenvba.apps.excel._pivot_layout import order

    field = cache_fields[index]
    group = field_group(field)
    if group is None or group.find(NS + 'groupItems') is None:
        return sorted_places(item_entries(field))
    length = len(_entries(group, 'groupItems'))
    mapping = group.find(NS + 'discretePr')
    if mapping is None:
        return list(range(length))
    places = [int(entry.get('v', '0')) for entry in mapping]
    return _by_first(places, order(_entries(cache_fields[base_of(field, index)], 'sharedItems')), length)


def item_name(cache_fields: list[XmlElement], index: int, place: int) -> str | None:
    """A grouping's item's name where the file's text is not it: the first of a range's, ``(blank)`` in the
    file where the field holds blanks, is ``<Start`` (tests/fixtures/pivot_grouping/ mixed_gappy)."""
    element = ranged(cache_fields[index])
    if element is None or place != 0:
        return None
    base = base_of(cache_fields[index], index)
    values = _entries(cache_fields[base], 'sharedItems') if 0 <= base < len(cache_fields) else []
    return Grouping(element).labels(whole_numbers(values), False)[0]


# Range.Group and Range.Ungroup on a pivot table's cells.

def table_at(target: Range) -> PivotTable | None:
    """The pivot table a range's first cell is in, its page fields' cells too."""
    from pyopenvba.apps.excel._model import Range

    first = target.first
    for table in target.sheet.pivot_tables:
        extent = table.TableRange2()
        assert isinstance(extent, Range)
        area = extent.first
        if area.top <= first.top <= area.bottom and area.left <= first.left <= area.right:
            return table.entry
    return None


def _cells(target: Range) -> list[tuple[int, int]]:
    return [(row, column) for area in target.areas for row in range(area.top, area.bottom + 1)
            for column in range(area.left, area.right + 1)]


def _picked(table: PivotTable, target: Range) -> tuple[int, list[int]]:
    """The field the cells show items of, and the places in the cache of the items, in the order the cells name
    them: error 1004 where a cell shows no item or the cells show more than one field's."""
    from pyopenvba.apps.excel._pivot_ranges import cell_item

    field: int | None = None
    places: list[int] = []
    for row, column in _cells(target):
        found = cell_item(table, row, column)
        if found is None or (field is not None and found[0] != field):
            raise error(1004)
        field = found[0]
        if found[1] not in places:
            places.append(found[1])
    if field is None:
        raise error(1004)
    return field, places


def group(target: Range, start: object, end: object, by: object, periods: object) -> object:
    """Range.Group where the range is a pivot table's: one item's cell groups its field by ranges or periods, the
    cells of several items gather them into a group (tests/fixtures/pivot_grouping/)."""
    table = table_at(target)
    if table is None:
        raise VBAUnsupportedError('Group with pivot/date/number arguments is not implemented')
    from pyopenvba.apps.excel._pivot_items import has_items

    field, places = _picked(table, target)
    fields = fields_of(table.cache)
    if has_items(table.cache, field):
        # A field with calculated items groups by no ranges (tests/fixtures/pivot_calculated_items/
        # group_numbers_with_item); Excel groups its items, which the model does not.
        if len(_cells(target)) == 1 and {entry.tag for entry in _entries(fields[field], 'sharedItems')
                                         if entry.get('f') is None} <= {NS + 'n', NS + 'm'}:
            raise error(1004)
        raise VBAUnsupportedError('Grouping a pivot field with calculated items is not implemented')
    if len(_cells(target)) > 1:
        if len(places) < 2:
            raise VBAUnsupportedError('Grouping the cells of one pivot item is not implemented')
        gather(table, field, places)
        return True
    base = base_of(fields[field], field)
    if is_group_field(fields[field]) and ranged(fields[field]) is not None:
        field = base
    values = _entries(fields[field], 'sharedItems')
    kinds = {entry.tag[len(NS):] for entry in values} - {'m'}
    if not is_database(fields[field]) or not kinds or not kinds <= {'n', 'd'} or kinds == {'n', 'd'}:
        raise error(1004)
    if kinds == {'n'}:
        group_numbers(table, field, values, start, end, by)
    else:
        group_dates(table, field, values, start, end, by, periods)
    return True


def _given(value: object) -> bool:
    """Whether Start or End names a value, rather than True or nothing for the field's own."""
    if value is MISSING or value is EMPTY:
        return False
    if isinstance(value, bool):
        if not value:
            raise error(1004)
        return False
    return True


def _numeric(value: object) -> float:
    try:
        if isinstance(value, VBADate):
            return value.serial
        return float(to_number(value))
    except (VBAError, ValueError, TypeError, OverflowError):
        raise error(13) from None


def group_numbers(table: PivotTable, field: int, values: list[XmlElement], start: object, end: object,
                  by: object) -> None:
    numbers = [float(entry.get('v', '0')) for entry in values if entry.tag == NS + 'n']
    low = _numeric(start) if _given(start) else min(numbers)
    high = _numeric(end) if _given(end) else max(numbers)
    if not _given(start) and not _given(end) and high == low:
        high = low + 1
    if low > high:
        raise error(1004)
    if low == high:
        raise VBAUnsupportedError('Grouping pivot numbers from a Start to the same End is not implemented')
    interval = default_interval(high - low) if by is MISSING else _numeric(by)
    if interval <= 0:
        raise error(1004)
    attributes: dict[str, str] = {}
    if _given(start):
        attributes['autoStart'] = '0'
    if _given(end):
        attributes['autoEnd'] = '0'
    attributes.update(startNum=number_spelling(low), endNum=number_spelling(high))
    if interval != 1:
        attributes['groupInterval'] = number_spelling(interval)
    cache_field = fields_of(table.cache)[field]
    if parent_of(cache_field) is not None:
        raise VBAUnsupportedError('Grouping by ranges a pivot field whose items are grouped is not implemented')
    old = field_group(cache_field)
    if old is not None:
        cache_field.remove(old)
    made = node('fieldGroup', {'base': str(field)})
    made.append(node('rangePr', attributes))
    _rewrite_items(made, number_labels(low, high, interval, whole_numbers(values), blanks(values)))
    cache_field.append(made)
    for each in _tables(table.cache):
        _relist(each, field)
    _render_all(table.cache)


def _periods(periods: object) -> list[str]:
    if periods is MISSING:
        return ['months']
    if not isinstance(periods, VBAArray):
        raise error(1004)
    chosen = periods.elements()
    if len(chosen) != len(PERIODS):
        raise error(1004)
    try:
        picked = [period for period, value in zip(PERIODS, chosen, strict=True) if to_bool(value)]
    except (VBAError, ValueError, TypeError):
        raise error(13) from None
    if not picked:
        raise error(1004)
    return picked


def _date(value: object) -> dt.datetime:
    if isinstance(value, VBADate):
        return value.to_datetime()
    return VBADate(_numeric(value)).to_datetime()


def group_dates(table: PivotTable, field: int, values: list[XmlElement], start: object, end: object, by: object,
                periods: object) -> None:
    picked = _periods(periods)
    found = automatic_dates(values)
    assert found is not None
    low = _date(start) if _given(start) else found[0]
    high = _date(end) if _given(end) else found[1]
    if low > high:
        raise error(1004)
    interval = 1
    if picked == ['days'] and by is not MISSING:
        interval = int(_numeric(by))
        if interval <= 0:
            raise error(1004)
    cache = table.cache
    fields = fields_of(cache)
    name = decode_text(fields[field].get('name', ''))
    # Grouping a date again takes its periods away first (tests/fixtures/pivot_grouping/ dates_regroup).
    places = _remove_periods(cache, field)
    fields = fields_of(cache)
    container = cache.root.find(NS + 'cacheFields')
    assert container is not None
    made: list[int] = []
    for period in picked:
        attributes: dict[str, str] = {}
        if _given(start):
            attributes['autoStart'] = '0'
        if _given(end):
            attributes['autoEnd'] = '0'
        attributes.update(groupBy=period, startDate=low.isoformat(), endDate=high.isoformat())
        if period == 'days' and interval > 1:
            attributes['groupInterval'] = str(interval)
        cache_field = node('cacheField', {'name': f'{period.title()} ({name})', 'numFmtId': '0',
                                          'databaseField': '0'})
        grouping = node('fieldGroup', {'base': str(field)})
        grouping.append(node('rangePr', attributes))
        _rewrite_items(grouping, date_labels(period, low, high, interval, blanks(values)))
        cache_field.append(grouping)
        container.append(cache_field)
        made.append(len(container) - 1)
    container.set('count', str(len(container)))
    _set_parent(fields[field], made[-1])
    for each in _tables(cache):
        _add_fields(each, len(made))
        root = each.root
        here = places.get(id(each))
        if here is None:
            here = _axis_of(root, field)
            # The table grouping the date shows its periods in its place; another showing it keeps it, its periods
            # outside it (tests/fixtures/pivot_grouping/ shared_dates).
            if here is not None and each is table:
                _take_off(root, field)
        if here is not None:
            _place(root, here, list(reversed(made)))
        for index in made:
            _list_items(each, index)
    _render_all(cache)


def gather(table: PivotTable, field: int, places: list[int]) -> None:
    """Gather a field's items into a new group, in the field that groups them, made where it has none."""
    cache = table.cache
    fields = fields_of(cache)
    chosen = fields[field]
    if ranged(chosen) is not None or (not is_database(chosen) and discrete(chosen) is None):
        raise VBAUnsupportedError('Grouping the items of a pivot field grouped by ranges or periods is not '
                                  'implemented')
    base = base_of(chosen, field)
    parent = parent_of(chosen)
    container = cache.root.find(NS + 'cacheFields')
    assert container is not None
    if parent is None:
        # The field of groups starts with each of the field's items alone (tests/fixtures/pivot_grouping/
        # manual_group, manual_second_level).
        own = item_entries(chosen)
        name = _free_name(fields, decode_text(fields[base].get('name', '')))
        made = node('cacheField', {'name': name, 'numFmtId': '0', 'databaseField': '0'})
        grouping = node('fieldGroup', {'base': str(base)})
        through = _through(fields, field)
        grouping.append(_mapping(through))
        grouping.append(node('groupItems', {'count': str(len(own))}))
        grouping[-1].extend(_copy(entry) for entry in own)
        made.append(grouping)
        container.append(made)
        container.set('count', str(len(container)))
        parent = len(container) - 1
        _set_parent(chosen, parent)
        for each in _tables(cache):
            _add_fields(each, 1)
            spot = _axis_of(each.root, field)
            if each is table and spot is not None:
                _place(each.root, spot, [parent])
        fields = fields_of(cache)
    grouping = field_group(fields[parent])
    assert grouping is not None
    entries = _entries(grouping, 'groupItems')
    own = item_entries(fields[field])
    chosen_places = set(places)
    survivors: list[tuple[XmlElement, list[int]]] = []
    for entry, members in zip(entries, _members(fields, parent, field), strict=True):
        kept = [member for member in members if member not in chosen_places]
        if not kept:
            continue
        if len(kept) == 1 and len(members) > 1:
            # A group left with one item is that item again (tests/fixtures/pivot_grouping/ manual_overlap).
            survivors.append((_copy(own[kept[0]]), kept))
        else:
            survivors.append((entry, kept))
    _write_groups(fields, parent, field, survivors, sorted(places))
    # Each table lists the field of groups' items afresh, each at its first item (tests/fixtures/pivot_grouping/
    # manual_overlap).
    for each in _tables(cache):
        _list_items(each, parent)
    _render_all(cache)


def _write_groups(fields: list[XmlElement], parent: int, field: int, survivors: list[tuple[XmlElement, list[int]]],
                  gathered: list[int] | None) -> None:
    """Write a field of groups' items, the entries left and a new group of ``gathered`` after them, named
    ``Group`` with the lowest number neither they nor the items it groups take (tests/fixtures/pivot_grouping/
    manual_renumber, manual_swallow, manual_second_level)."""
    grouping = field_group(fields[parent])
    assert grouping is not None
    if gathered is not None:
        taken = {decode_text(entry.get('v', '')).casefold()
                 for entry in [*(entry for entry, _ in survivors), *item_entries(fields[field])]
                 if entry.tag == NS + 's'}
        number = 1
        while f'group{number}' in taken:
            number += 1
        survivors = [*survivors, (node('s', {'v': f'Group{number}'}), gathered)]
    listed = grouping.find(NS + 'groupItems')
    assert listed is not None
    listed[:] = [entry for entry, _ in survivors]
    listed.set('count', str(len(listed)))
    of_member = {member: spot for spot, (_, members) in enumerate(survivors) for member in members}
    mapping = grouping.find(NS + 'discretePr')
    assert mapping is not None
    mapping[:] = [node('x', {'v': str(of_member.get(place, 0))}) for place in _through(fields, field)]
    mapping.set('count', str(len(mapping)))


def _copy(entry: XmlElement) -> XmlElement:
    """A field's item as a field of groups lists it: its value alone, a number spelled as Excel spells it
    (tests/fixtures/pivot_grouping/ manual_numbers)."""
    made = node(entry.tag[len(NS):], {'v': entry.get('v', '')} if entry.get('v') is not None else None)
    if entry.tag == NS + 'n':
        made.set('v', number_spelling(_number(entry.get('v'))))
    return made


def _through(fields: list[XmlElement], field: int) -> list[int]:
    """For each of a field's base's shared items, the item of the field it stands at."""
    base = base_of(fields[field], field)
    shared = _entries(fields[base], 'sharedItems')
    mapping = discrete(fields[field]) if field != base else None
    if mapping is None:
        return list(range(len(shared)))
    return [int(entry.get('v', '0')) for entry in mapping]


def _mapping(places: list[int]) -> XmlElement:
    made = node('discretePr', {'count': str(len(places))})
    made.extend(node('x', {'v': str(place)}) for place in places)
    return made


def _members(fields: list[XmlElement], parent: int, field: int) -> list[list[int]]:
    """Each entry of a field of groups' items, as the places of the items of the field it groups."""
    grouping = field_group(fields[parent])
    assert grouping is not None
    count = len(_entries(grouping, 'groupItems'))
    found: list[list[int]] = [[] for _ in range(count)]
    mapping = discrete(fields[parent])
    assert mapping is not None
    upper = [int(entry.get('v', '0')) for entry in mapping]
    lower = _through(fields, field)
    for spot, place in enumerate(lower):
        if spot < len(upper) and 0 <= upper[spot] < count and place not in found[upper[spot]]:
            found[upper[spot]].append(place)
    return found


def is_group(entry: XmlElement, members: list[int], own: list[XmlElement]) -> bool:
    """Whether an entry of a field of groups is a group, rather than an item of the field it groups standing
    alone: one that gathers more than one item, or another than its own value."""
    from pyopenvba.apps.excel._pivot_layout import item_key

    if len(members) != 1 or not 0 <= members[0] < len(own):
        return len(members) > 1
    return item_key(entry) != item_key(own[members[0]])


def _free_name(fields: list[XmlElement], name: str) -> str:
    """A new field of groups' name: its base's with the lowest number from 2 no field has (tests/fixtures/
    pivot_grouping/ manual_second_level)."""
    taken = {decode_text(field.get('name', '')).casefold() for field in fields}
    number = 2
    while f'{name}{number}'.casefold() in taken:
        number += 1
    return f'{name}{number}'


def _set_parent(cache_field: XmlElement, parent: int) -> None:
    group = field_group(cache_field)
    if group is None:
        group = node('fieldGroup')
        cache_field.append(group)
    attributes = {'par': str(parent), **{key: value for key, value in group.attrib.items() if key != 'par'}}
    group.attrib.clear()
    group.attrib.update(attributes)


def _clear_parent(cache_field: XmlElement) -> None:
    group = field_group(cache_field)
    if group is None:
        return
    group.attrib.pop('par', None)
    if not group.attrib and not len(group):
        cache_field.remove(group)


def ungroup(target: Range) -> object:
    """Range.Ungroup on a pivot table's cell: a range's or a period's field back to its values, a group's items
    back to standing alone (tests/fixtures/pivot_grouping/ numbers_ungroup, dates_ungroup_years,
    manual_ungroup)."""
    table = table_at(target)
    if table is None:
        return None
    field, places = _picked(table, target)
    cache = table.cache
    fields = fields_of(cache)
    chosen = fields[field]
    if ranged(chosen) is not None and is_database(chosen):
        group_of = field_group(chosen)
        assert group_of is not None
        chosen.remove(group_of)
        for each in _tables(cache):
            _relist(each, field)
        _render_all(cache)
        return True
    if ranged(chosen) is not None:
        base = base_of(chosen, field)
        here = _remove_periods(cache, base)
        for each in _tables(cache):
            spot = here.get(id(each))
            if spot is not None and _axis_of(each.root, base) is None:
                _place(each.root, spot, [base])
        _render_all(cache)
        return True
    if discrete(chosen) is None or len(places) != 1:
        raise error(1004)
    if parent_of(chosen) is not None:
        raise VBAUnsupportedError('Ungrouping a pivot group that is grouped again is not implemented')
    child = _child_of(fields, field)
    if child is None:
        raise error(1004)
    groups = _members(fields, field, child)
    entries = item_entries(chosen)
    own = item_entries(fields[child])
    place = places[0]
    if not 0 <= place < len(entries) or not is_group(entries[place], groups[place], own):
        raise error(1004)
    if not any(is_group(entry, groups[spot], own) for spot, entry in enumerate(entries) if spot != place):
        # Ungrouping the last group takes its field away (tests/fixtures/pivot_grouping/ manual_ungroup).
        _remove_fields(cache, [field])
        _clear_parent(fields_of(cache)[child])
        _render_all(cache)
        return True
    # Its items stand alone after the others in the cache, and in its place in each table's list
    # (tests/fixtures/pivot_grouping/ manual_ungroup_one).
    survivors = [(entry, groups[spot]) for spot, entry in enumerate(entries) if spot != place]
    survivors.extend((_copy(own[member]), [member]) for member in sorted(groups[place]))
    _write_groups(fields, field, child, survivors, None)
    moved = {spot: spot - (spot > place) for spot in range(len(entries)) if spot != place}
    alone = list(range(len(entries) - 1, len(survivors)))
    for each in _tables(cache):
        _unfold(each, field, place, moved, alone)
    _render_all(cache)
    return True


def _unfold(table: PivotTable, field: int, place: int, moved: dict[int, int], alone: list[int]) -> None:
    """Put a group's items where it stood in a table's list of its field's items, the others renumbered."""
    from pyopenvba.apps.excel._pivot_filters import item_node

    pivot = table.root.findall(f'{NS}pivotFields/{NS}pivotField')[field]
    items = pivot.find(NS + 'items')
    if items is None:
        _list_items(table, field)
        return
    fresh: list[XmlElement] = []
    for item in list(items):
        if item.get('t', 'data') != 'data':
            fresh.append(item)
            continue
        spot = int(item.get('x', '-1'))
        if spot == place:
            fresh.extend(item_node(each) for each in alone)
        else:
            item.set('x', str(moved.get(spot, spot)))
            fresh.append(item)
    items[:] = fresh
    items.set('count', str(len(fresh)))


def _child_of(fields: list[XmlElement], parent: int) -> int | None:
    return next((index for index, field in enumerate(fields) if parent_of(field) == parent), None)


# Keeping each table of the cache in step with its fields.

def _tables(cache: PivotCache) -> list[PivotTable]:
    return [table for sheet in cache.book.sheets_ for table in sheet.pivot_tables if table.cache is cache]


def _render_all(cache: PivotCache) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    for table in _tables(cache):
        table.changed = True
        rerender(table, deepcopy(table.root))
    cache.changed = True
    cache.book.saved = False


def _axis_of(root: XmlElement, index: int) -> tuple[str, int] | None:
    for name in ('rowFields', 'colFields'):
        entries = root.findall(f'{NS}{name}/{NS}field')
        for position, entry in enumerate(entries):
            if int(entry.get('x', '-1')) == index:
                return name, position
    return None


def _take_off(root: XmlElement, index: int) -> None:
    """Take a field off its axis, its items kept, as a date leaves its place to its periods."""
    for name in ('rowFields', 'colFields'):
        container = root.find(NS + name)
        if container is None:
            continue
        for entry in list(container):
            if int(entry.get('x', '-1')) == index:
                container.remove(entry)
        container.set('count', str(len(container)))
    pivot = root.findall(f'{NS}pivotFields/{NS}pivotField')[index]
    pivot.attrib.pop('axis', None)


def _place(root: XmlElement, spot: tuple[str, int], indices: list[int]) -> None:
    """Put fields on an axis at a place, in the order given, the first outermost."""
    from pyopenvba.apps.excel._pivot_filters import field_attributes

    name, position = spot
    container = root.find(NS + name)
    if container is None:
        container = node(name)
        anchor = root.find(NS + 'pivotFields')
        assert anchor is not None
        root.insert(list(root).index(anchor) + 1 + (name == 'colFields' and root.find(NS + 'rowItems') is not None)
                    + (name == 'colFields' and root.find(NS + 'rowFields') is not None), container)
    for offset, index in enumerate(indices):
        container.insert(min(position + offset, len(container)), node('field', {'x': str(index)}))
    container.set('count', str(len(container)))
    pivots = root.findall(f'{NS}pivotFields/{NS}pivotField')
    for index in indices:
        attributes = dict(pivots[index].attrib)
        attributes['axis'] = 'axisRow' if name == 'rowFields' else 'axisCol'
        pivots[index].attrib.clear()
        pivots[index].attrib.update(field_attributes(attributes))


def _add_fields(table: PivotTable, count: int) -> None:
    container = table.root.find(NS + 'pivotFields')
    assert container is not None
    for _ in range(count):
        container.append(node('pivotField', {'showAll': '0'}))
    container.set('count', str(len(container)))


def _list_items(table: PivotTable, index: int) -> None:
    """List a grouped field's items in a table, as Excel lists a new grouping's in every table of its cache
    (tests/fixtures/pivot_grouping/ shared_manual)."""
    from pyopenvba.apps.excel._pivot_filters import item_node
    from pyopenvba.apps.excel._pivot_forms import functions_of

    pivot = table.root.findall(f'{NS}pivotFields/{NS}pivotField')[index]
    old = pivot.find(NS + 'items')
    if old is not None:
        pivot.remove(old)
    items = node('items')
    for place in natural_places(fields_of(table.cache), index):
        items.append(item_node(place))
    for kind in functions_of(pivot):
        items.append(node('item', {'t': kind}))
    items.set('count', str(len(items)))
    pivot.insert(0, items)


def _relist(table: PivotTable, index: int) -> None:
    """A field whose items changed what they are lists them afresh where it is on an axis, and lists none
    elsewhere."""
    pivot = table.root.findall(f'{NS}pivotFields/{NS}pivotField')[index]
    old = pivot.find(NS + 'items')
    if old is not None:
        pivot.remove(old)
    if pivot.get('axis') in ('axisRow', 'axisCol', 'axisPage'):
        _list_items(table, index)


def _remove_periods(cache: PivotCache, base: int) -> dict[int, tuple[str, int]]:
    """Take away a date's fields of periods, and say where each table showed the outermost of them."""
    fields = fields_of(cache)
    gone = [index for index, field in enumerate(fields)
            if index != base and is_group_field(field) and ranged(field) is not None and base_of(field, index) == base]
    places: dict[int, tuple[str, int]] = {}
    for table in _tables(cache):
        spots = [spot for index in gone if (spot := _axis_of(table.root, index)) is not None]
        if spots:
            places[id(table)] = min(spots, key=lambda spot: (spot[0] != spots[0][0], spot[1]))
    if gone:
        _remove_fields(cache, gone)
        _clear_parent(fields_of(cache)[base])
    return places


def _remove_fields(cache: PivotCache, gone: list[int]) -> None:
    """Remove cache fields and every table's fields for them, renumbering what names the fields after them."""
    removed = set(gone)
    fields = fields_of(cache)
    after = {index: index - sum(1 for each in removed if each < index) for index in range(len(fields))
             if index not in removed}
    container = cache.root.find(NS + 'cacheFields')
    assert container is not None
    for index in sorted(removed, reverse=True):
        container.remove(fields[index])
    container.set('count', str(len(container)))
    for field in container:
        group = field_group(field)
        if group is None:
            continue
        for name in ('par', 'base'):
            value = group.get(name)
            if value is not None:
                if int(value) in removed:
                    group.attrib.pop(name)
                else:
                    group.set(name, str(after[int(value)]))
        if not group.attrib and not len(group):
            field.remove(group)
    for table in _tables(cache):
        root = table.root
        if any(int(entry.get('fld', '-1')) in removed for entry in root.findall(f'{NS}dataFields/{NS}dataField')):
            raise VBAUnsupportedError('Ungrouping a pivot field a data field summarizes is not implemented')
        pivots = root.find(NS + 'pivotFields')
        assert pivots is not None
        for index in sorted(removed, reverse=True):
            pivots.remove(pivots[index])
        pivots.set('count', str(len(pivots)))
        for name, attribute in (('rowFields', 'x'), ('colFields', 'x'), ('pageFields', 'fld')):
            holder = root.find(NS + name)
            if holder is None:
                continue
            for entry in list(holder):
                value = int(entry.get(attribute, '-1'))
                if value in removed:
                    holder.remove(entry)
                elif value in after:
                    entry.set(attribute, str(after[value]))
            holder.set('count', str(len(holder)))
            if not len(holder):
                root.remove(holder)
        for entry in root.findall(f'{NS}dataFields/{NS}dataField'):
            entry.set('fld', str(after[int(entry.get('fld', '0'))]))
            base = int(entry.get('baseField', '0'))
            if base in after:
                entry.set('baseField', str(after[base]))
        filters = root.find(NS + 'filters')
        if filters is not None:
            for rule in list(filters):
                value = int(rule.get('fld', '-1'))
                if value in removed:
                    filters.remove(rule)
                elif value in after:
                    rule.set('fld', str(after[value]))
            filters.set('count', str(len(filters)))
            if not len(filters):
                root.remove(filters)


# What a grouping says of the fields and items it relates.

def _cache_fields(field: PivotField) -> list[XmlElement]:
    return fields_of(field.table.entry.cache)


def parent_field(field: PivotField) -> object:
    """ParentField: the field that groups this one, error 1004 where none does (tests/fixtures/pivot_grouping/
    relations_*)."""
    fields = _cache_fields(field)
    parent = parent_of(fields[field.index]) if 0 <= field.index < len(fields) and not field.data else None
    if parent is None:
        raise error(1004)
    return field.table.PivotFields(parent + 1)


def child_field(field: PivotField) -> object:
    """ChildField: the field a field of groups groups; error 1004 for any other, a date's periods' too."""
    fields = _cache_fields(field)
    if field.data or not 0 <= field.index < len(fields) or discrete(fields[field.index]) is None:
        raise error(1004)
    child = _child_of(fields, field.index)
    if child is None:
        raise error(1004)
    return field.table.PivotFields(child + 1)


def _family(fields: list[XmlElement], index: int) -> tuple[int, int]:
    """How many fields a field's base and the fields grouping it make, and how far down from the top the field is."""
    base = base_of(fields[index], index)
    size = 1 + sum(1 for spot, field in enumerate(fields)
                   if spot != base and is_group_field(field) and base_of(field, spot) == base)
    depth = 1
    seen = {index}
    parent = parent_of(fields[index])
    while parent is not None and parent not in seen and 0 <= parent < len(fields):
        seen.add(parent)
        depth += 1
        parent = parent_of(fields[parent])
    return size, depth


def group_level(field: PivotField) -> object:
    """GroupLevel: 1 for a field nothing groups, one more for each field grouping it above (tests/fixtures/
    pivot_grouping/ relations_levels, relations_date_levels)."""
    from pyopenvba.interpreter._values import VBAInt

    fields = _cache_fields(field)
    if field.data or not 0 <= field.index < len(fields):
        return VBAInt(1, 'Long')
    return VBAInt(_family(fields, field.index)[1], 'Long')


def total_levels(field: PivotField) -> object:
    """TotalLevels: the fields a grouping makes with the field it groups, together."""
    from pyopenvba.interpreter._values import VBAInt

    fields = _cache_fields(field)
    if field.data or not 0 <= field.index < len(fields):
        return VBAInt(1, 'Long')
    return VBAInt(_family(fields, field.index)[0], 'Long')


def _children(item: PivotItem) -> list[int]:
    """The places of the items a field of groups' item gathers, none where it is an item standing alone."""
    fields = _cache_fields(item.field)
    field = item.field.index
    if item.field.data or not 0 <= field < len(fields) or discrete(fields[field]) is None or item.x < 0:
        return []
    child = _child_of(fields, field)
    if child is None:
        return []
    members = _members(fields, field, child)
    entries = item_entries(fields[field])
    if not 0 <= item.x < len(members) or not is_group(entries[item.x], members[item.x],
                                                         item_entries(fields[child])):
        return []
    return members[item.x]


def _pivot_field(table: PivotTable, index: int) -> PivotField:
    from pyopenvba.apps.excel._pivots import PivotField

    found = table.PivotFields(index + 1)
    assert isinstance(found, PivotField)
    return found


def parent_items(field: PivotField) -> PivotItems:
    """ParentItems: a field of groups' groups, the items standing alone left out."""
    from pyopenvba.apps.excel._pivots import PivotItems

    found = [item for item in field.listed_items() if _children(item)]
    return PivotItems(field, list(found))


def child_items(item: PivotItem, index: object) -> object:
    """ChildItems: the items a group gathers, in its field's order; none for any other item."""
    from pyopenvba.apps.excel._pivots import PivotItems

    places = set(_children(item))
    fields = _cache_fields(item.field)
    found: list[object] = []
    if places:
        child = _child_of(fields, item.field.index)
        assert child is not None
        found = [each for each in _pivot_field(item.field.table, child).listed_items() if each.x in places]
    collection = PivotItems(item.field, found)
    return collection if index is MISSING else collection.Item(index)


def parent_item(item: PivotItem) -> object:
    """ParentItem: the group an item is in, error 1004 where it stands alone or nothing groups its field."""
    fields = _cache_fields(item.field)
    field = item.field.index
    if item.field.data or not 0 <= field < len(fields) or item.x < 0:
        raise error(1004)
    parent = parent_of(fields[field])
    if parent is None or discrete(fields[parent]) is None:
        raise error(1004)
    groups = _members(fields, parent, field)
    entries = item_entries(fields[parent])
    own = item_entries(fields[field])
    for spot, members in enumerate(groups):
        if item.x in members and is_group(entries[spot], members, own):
            for each in _pivot_field(item.field.table, parent).listed_items():
                if each.x == spot:
                    return each
    raise error(1004)


def carry_groups(before: list[XmlElement], after: list[XmlElement]) -> None:
    """Carry a cache's groupings over a refresh to the source's fields made again."""
    for old, new in zip(before, after, strict=False):
        group = field_group(old)
        if group is not None and is_database(old) and new.find(NS + 'fieldGroup') is None:
            new.append(deepcopy(group))


def remap_groups(before: list[XmlElement], after: list[XmlElement]) -> bool:
    """Find each base's groups again by its items' values once a refresh listed them afresh, a new item standing
    alone after the others (tests/fixtures/pivot_grouping/ refresh_manual); whether any field of groups changed."""
    from pyopenvba.apps.excel._pivot_layout import item_key

    changed = False
    for index, field in enumerate(after):
        mapping = discrete(field)
        group = field_group(field)
        if mapping is None or group is None or index >= len(before):
            continue
        base = base_of(field, index)
        was = _entries(before[base], 'sharedItems')
        now = _entries(after[base], 'sharedItems')
        places = [int(entry.get('v', '0')) for entry in mapping]
        known = {item_key(entry): places[spot] for spot, entry in enumerate(was) if spot < len(places)}
        listed = group.find(NS + 'groupItems')
        assert listed is not None
        fresh: list[int] = []
        for entry in now:
            key = item_key(entry)
            if key not in known:
                listed.append(_copy(entry))
                known[key] = len(listed) - 1
            fresh.append(known[key])
        listed.set('count', str(len(listed)))
        changed = changed or fresh != places
        mapping[:] = [node('x', {'v': str(place)}) for place in fresh]
        mapping.set('count', str(len(mapping)))
    return changed


def extend_groups(before: list[XmlElement], fields: list[XmlElement]) -> None:
    """Give each value a refresh keeps unused, after the source's, its group again, found by what it is."""
    from pyopenvba.apps.excel._pivot_layout import item_key

    for index, field in enumerate(fields):
        mapping = discrete(field)
        group = field_group(field)
        if mapping is None or group is None:
            continue
        base = base_of(field, index)
        now = _entries(fields[base], 'sharedItems')
        if len(mapping) >= len(now):
            continue
        was = _entries(before[base], 'sharedItems') if base < len(before) else []
        old = discrete(before[index]) if index < len(before) else None
        places = [int(entry.get('v', '0')) for entry in old] if old is not None else []
        known = {item_key(entry): places[spot] for spot, entry in enumerate(was) if spot < len(places)}
        listed = group.find(NS + 'groupItems')
        assert listed is not None
        for entry in now[len(mapping):]:
            key = item_key(entry)
            if key not in known:
                listed.append(_copy(entry))
                known[key] = len(listed) - 1
            mapping.append(node('x', {'v': str(known[key])}))
        listed.set('count', str(len(listed)))
        mapping.set('count', str(len(mapping)))
