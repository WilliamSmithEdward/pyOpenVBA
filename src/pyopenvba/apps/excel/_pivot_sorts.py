"""Sorting a pivot table's items, as Excel does.

A field sorted by its items' names keeps them in that order in its
``items`` list, ``sortType`` saying which way, and sorts a refresh's new
items in among them. One sorted by a data field's values keeps its list
and adds an ``autoSortScope`` naming the data field: the table orders the
field's items as it lays them out, by the values of the records each
shows, inside each of its outer items. Equal values keep the list's order,
and a descending sort is the ascending one turned round. xlManual and a
PivotItem.Position drop both, keeping the list as it stands.

Measured with scripts/measure_pivot_sorts.py (tests/fixtures/pivot_sorts/):
AutoSort ignores an order other than xlAscending, xlDescending and
xlManual and a field other than the field itself or a data field, by its
caption in any case; Position counts the items showing from 1, is #N/A
for a hidden item and an item of a field on no axis, and moving an item
puts it straight after the item showing before its new place, a hidden
item it moves showing again.
"""
from __future__ import annotations

from collections.abc import Callable
from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement
from pyopenvba.apps.excel._pivot_create import node
from pyopenvba.apps.excel._pivot_filters import (
    COLUMN,
    ROW,
    ensure_items,
    field_attributes,
    flag,
    listed,
    mark_changed,
    orientation,
    place_of,
    rerender,
    restate,
    shared,
    showing_places,
)
from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAErrorValue, error, string_order, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._pivots import PivotField, PivotItem

#: The Values field, -2, as the unsigned number a sort's scope names it by.
DATA_FIELD = '4294967294'
ORDERS = {1: 'ascending', 2: 'descending'}
MANUAL = -4135
#: The #N/A Position gives an item that does not show.
NOT_SHOWN = 2042
#: How a value sort orders a group of records: by the total it comes to, as value_key compares totals.
SortKey = Callable[[list[int]], tuple[int, float, str]]


def sort_kind(field: XmlElement) -> str | None:
    kind = field.get('sortType')
    return kind if kind in ('ascending', 'descending') else None


def sort_scope(field: XmlElement) -> int | None:
    """The data field whose values a field's items are sorted by, or None."""
    reference = field.find(f'{NS}autoSortScope/{NS}pivotArea/{NS}references/{NS}reference')
    if reference is None or reference.get('field') != DATA_FIELD:
        return None
    place = reference.find(NS + 'x')
    return int(place.get('v', '0')) if place is not None else None


def scope_node(data: int) -> XmlElement:
    """A sort's scope as Excel writes one: the data field, by its place among the data fields."""
    scope = node('autoSortScope')
    area = node('pivotArea', {'dataOnly': '0', 'outline': '0', 'fieldPosition': '0'})
    references = node('references', {'count': '1'})
    reference = node('reference', {'field': DATA_FIELD, 'count': '1', 'selected': '0'})
    reference.append(node('x', {'v': str(data)}))
    references.append(reference)
    area.append(references)
    scope.append(area)
    return scope


def set_sort(field: XmlElement, kind: str | None, data: int | None) -> None:
    """Give a pivotField its sortType and the data field it sorts by, or take them away."""
    attributes = {key: value for key, value in field.attrib.items() if key != 'sortType'}
    if kind is not None:
        attributes['sortType'] = kind
    field.attrib.clear()
    field.attrib.update(field_attributes(attributes))
    old = field.find(NS + 'autoSortScope')
    if old is not None:
        field.remove(old)
    if data is not None:
        items = field.find(NS + 'items')
        field.insert(list(field).index(items) + 1 if items is not None else 0, scope_node(data))


def by_name(entries: list[XmlElement], descending: bool) -> list[int]:
    """The places of a list's values sorted by their names as a field's first list is, turned round for a
    descending sort."""
    from pyopenvba.apps.excel._pivot_layout import order

    ranked = order(entries)
    return ranked[::-1] if descending else ranked


def relist(field: XmlElement, data_items: list[XmlElement]) -> None:
    """Put a pivotField's items in a new order, its subtotal entries after them as before."""
    items = field.find(NS + 'items')
    assert items is not None
    rest = [item for item in items if item.get('t', 'data') != 'data']
    items[:] = [*data_items, *rest]


def sort_by_name(table_cache_fields: list[XmlElement], field: XmlElement, index: int, descending: bool) -> None:
    entries = shared(table_cache_fields[index]) if index < len(table_cache_fields) else []
    items = listed(field)
    values = [entries[place_of(item)] for item in items if 0 <= place_of(item) < len(entries)]
    if len(values) != len(items):
        raise VBAUnsupportedError('Sorting a pivot field whose items name no cached value is not implemented')
    relist(field, [items[place] for place in by_name(values, descending)])


def value_key(value: object) -> tuple[int, float, str]:
    """How a value sort compares totals: numbers, then blanks, then errors."""
    if isinstance(value, ExcelError):
        return 2, 0.0, str(value)
    if value is EMPTY or not isinstance(value, (int, float)):
        return 1, 0.0, ''
    return 0, float(value), ''


def value_sorts(fields: list[XmlElement], axes: set[int], reports: list[XmlElement],
                records: list[list[XmlElement]]) -> dict[int, tuple[bool, SortKey]]:
    """Each field on an axis sorted by a data field's values: whether descending, and the key a group of records
    sorts by."""
    from pyopenvba.apps.excel._pivot_aggregate import aggregate
    from pyopenvba.apps.excel._pivot_calculations import stepping

    # A field a data field's values step along, against the item before or after, shows in its items' order
    # (tests/fixtures/pivot_calculations/ previous_then_sort, sort_other_previous).
    stepped = {stepping(report) for report in reports}
    found: dict[int, tuple[bool, SortKey]] = {}
    for index in axes:
        field = fields[index]
        kind, data = sort_kind(field), sort_scope(field)
        if kind is None or data is None or not 0 <= data < len(reports) or index in stepped:
            continue
        source, function = int(reports[data].get('fld', '0')), reports[data].get('subtotal', 'sum')

        def key(rows: list[int], source: int = source, function: str = function) -> tuple[int, float, str]:
            return value_key(aggregate([records[row][source] for row in rows], function))

        found[index] = (kind == 'descending', key)
    return found


def _equal(left: str, right: str) -> bool:
    return string_order(left.casefold(), right.casefold()) == 0


def auto_sort(field: PivotField, order: object, name: object, line: object, subtotal: object) -> None:
    from pyopenvba.apps.excel._pivot_calculations import unsort_stepped

    if line is not MISSING or subtotal is not MISSING:
        raise VBAUnsupportedError('Sorting a pivot field by a PivotLine or a custom subtotal is not implemented')
    if field.data or field.index < 0:
        raise VBAUnsupportedError('Sorting the Values field or a data field is not implemented')
    chosen = int(to_integer(order, 'Long'))
    wanted = to_text(name)
    table = field.table.entry
    reports = table.root.findall(f'{NS}dataFields/{NS}dataField')
    own = str(field.Name())
    data: int | None = None
    if _equal(wanted, own):
        named = own
    else:
        data = next((place for place, report in enumerate(reports)
                     if _equal(decode_text(report.get('name', '')), wanted)), None)
        if data is None:
            return
        named = decode_text(reports[data].get('name', ''))
    if chosen not in (1, 2, MANUAL):
        return
    laid_out = orientation(field.root) != 0
    root = deepcopy(table.root) if laid_out else table.root
    pivot = root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index]
    if not ensure_items(table, pivot, field.index):
        raise VBAUnsupportedError('Sorting a field its cache keeps no list of values for is not implemented')
    if chosen == MANUAL:
        set_sort(pivot, None, None)
        table.sort_names[field.index] = named
    else:
        if data is None:
            sort_by_name(table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField'), pivot, field.index, chosen == 2)
        set_sort(pivot, ORDERS[chosen], data)
        unsort_stepped(root)
        table.sort_names.pop(field.index, None)
    if laid_out:
        rerender(table, root)
    else:
        mark_changed(table)


def sort_order(field: PivotField) -> int:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's AutoSortOrder is not implemented")
    return {'ascending': 1, 'descending': 2}.get(sort_kind(field.root) or '', MANUAL)


def sort_field(field: PivotField) -> str:
    if field.index == -2:
        raise error(1004)
    if field.data:
        raise VBAUnsupportedError("A data field's AutoSortField is not implemented")
    table = field.table.entry
    data = sort_scope(field.root)
    reports = table.root.findall(f'{NS}dataFields/{NS}dataField')
    if data is not None and 0 <= data < len(reports):
        return decode_text(reports[data].get('name', ''))
    return table.sort_names.get(field.index, str(field.Name()))


def position(item: PivotItem) -> object:
    """An item's place among the items of its field that show, from 1, or #N/A where it does not show."""
    if item.root is None:
        return 1
    if item.x < 0:
        return item.position
    places = showing_places(item.field) if orientation(item.field.root) != 0 else []
    return places.index(item.x) + 1 if item.x in places else VBAErrorValue(NOT_SHOWN)


def set_position(item: PivotItem, value: object) -> None:
    wanted = int(to_integer(value, 'Long'))
    field = item.field
    if item.root is None or item.x < 0 or field.data or field.index < 0:
        raise VBAUnsupportedError("Moving the Values field's items or a page field's (All) is not implemented")
    if orientation(field.root) not in (ROW, COLUMN):
        raise VBAUnsupportedError('Moving an item of a page field or of a field on no axis is not implemented')
    table = field.table.entry
    root = deepcopy(table.root)
    pivot = root.findall(f'{NS}pivotFields/{NS}pivotField')[field.index]
    items = listed(pivot)
    target = next((entry for entry in items if place_of(entry) == item.x), None)
    if target is None:
        raise error(1004)
    if flag(target, 'h'):
        restate(target, hide=False)
    others = [entry for entry in items if entry is not target]
    showing = [entry for entry in others if not flag(entry, 'h')]
    if not 1 <= wanted <= len(showing) + 1:
        raise error(1004)
    after = others.index(showing[wanted - 2]) + 1 if wanted > 1 else 0
    relist(pivot, [*others[:after], target, *others[after:]])
    set_sort(pivot, None, None)
    table.sort_names.pop(field.index, None)
    rerender(table, root)
