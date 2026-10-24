"""A pivot table's items collapsed and expanded, ShowDetail, as Excel takes and keeps them.

An item's ShowDetail False collapses it: the file marks its entry sd="0", and on the rows or columns it shows its
total alone, with no line inside it and no subtotal line (see _pivot_layout._lines). A field's ShowDetail sets
every item's, its subtotals' entries too, and reading it is error 1004. Measured with
scripts/measure_pivot_show_detail.py (tests/fixtures/pivot_show_detail/):

- An item of a field on the rows or columns, the innermost one's too, where it changes nothing shown, or of a field
  on no axis, whose items the field then lists, takes it; one of a page field's is error 1004.
- A field on an axis takes it but the innermost one there, which is error 1004, and so do a page field and a field
  on no axis; a data field and the Values field are error 1004.
- A value that is not True or False is error 1004.
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._pivots import NS
from pyopenvba.exceptions import VBAError
from pyopenvba.interpreter._values import VBAErrorValue, error, to_bool

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._pivots import PivotField, PivotItem

VALUES = -2


def _shown(value: object) -> bool:
    try:
        return to_bool(value)
    except (VBAError, ValueError, TypeError):
        raise error(1004) from None


def _mark(entry: XmlElement, shown: bool) -> None:
    from pyopenvba.apps.excel._pivot_filters import restate

    if shown:
        entry.attrib.pop('sd', None)
    else:
        entry.set('sd', '0')
    if entry.get('t', 'data') == 'data':
        restate(entry, hide=entry.get('h') in ('1', 'true'))


def _axis(field: PivotField) -> list[int]:
    """The fields on the axis a field is on, the Values field left out."""
    name = {'axisRow': 'rowFields', 'axisCol': 'colFields'}.get(field.root.get('axis', ''))
    if name is None:
        return []
    return [int(entry.get('x', '0')) for entry in field.table.entry.root.findall(f'{NS}{name}/{NS}field')
            if entry.get('x') != str(VALUES)]


def _laid_out(field: PivotField) -> None:
    from pyopenvba.apps.excel._pivot_filters import rerender

    table = field.table.entry
    table.changed = True
    table.sheet.book.saved = False
    if field.root.get('axis') in ('axisRow', 'axisCol'):
        rerender(table, deepcopy(table.root))


def item_detail(item: PivotItem) -> object:
    """An item's ShowDetail: True unless its entry is marked collapsed; #N/A for an item of the Values field
    (tests/fixtures/pivot_show_detail/ values_rows)."""
    from pyopenvba.apps.excel._pivot_names import item_entry

    if item.field.index == VALUES:
        return VBAErrorValue(2042)
    if item.field.data or item.field.index < 0 or item.x < 0:
        return True
    entry = item_entry(item.field, item.x)
    return not (entry is not None and entry.get('sd') in ('0', 'false'))


def set_item_detail(item: PivotItem, value: object) -> None:
    from pyopenvba.apps.excel._pivot_filters import ensure_items
    from pyopenvba.apps.excel._pivot_names import item_entry

    shown = _shown(value)
    field = item.field
    if field.data or field.index < 0 or item.x < 0 or field.root.get('axis') == 'axisPage':
        raise error(1004)
    ensure_items(field.table.entry, field.root, field.index)
    entry = item_entry(field, item.x)
    if entry is None:
        raise error(1004)
    _mark(entry, shown)
    _laid_out(field)


def field_detail(field: PivotField) -> object:
    raise error(1004)


def set_field_detail(field: PivotField, value: object) -> None:
    from pyopenvba.apps.excel._pivot_filters import ensure_items

    shown = _shown(value)
    if field.data or field.index < 0:
        raise error(1004)
    axis = _axis(field)
    if axis and axis[-1] == field.index:
        raise error(1004)
    ensure_items(field.table.entry, field.root, field.index)
    for entry in field.root.findall(f'{NS}items/{NS}item'):
        _mark(entry, shown)
    _laid_out(field)
