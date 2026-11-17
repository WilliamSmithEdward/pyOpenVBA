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

A value's own ShowDetail, Range.ShowDetail on a cell of the table's values, is the detail of the records behind it
(drill-through), measured with scripts/measure_pivot_drill.py (tests/fixtures/pivot_drill.json); see drill.
"""
from __future__ import annotations

from copy import deepcopy
from typing import TYPE_CHECKING

from pyopenvba.apps.excel._pivots import NS, decode_text
from pyopenvba.exceptions import VBAError, VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, VBAErrorValue, error, to_bool

if TYPE_CHECKING:
    from pyopenvba.apps.excel._charts_file import XmlElement
    from pyopenvba.apps.excel._model import Range, Worksheet
    from pyopenvba.apps.excel._pivots import PivotField, PivotItem, PivotTable

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


# --- a value's detail (drill-through) ---------------------------------------------------------------------------

_UNABLE_GET = "Unable to get the ShowDetail property of the Range class"
_UNABLE_SET = "Unable to set the ShowDetail property of the Range class"


def value_cell(target: Range) -> PivotTable | None:
    """The pivot table whose values ``target``'s first cell is one of, a total's among them; None elsewhere."""
    from pyopenvba.apps.excel._model import Range as Cells
    from pyopenvba.apps.excel._pivot_cells import owner

    at = target.first
    found = owner(target.sheet, at.top, at.left)
    if found is None:
        return None
    table = found[0]
    try:
        body = table.vba_get('DataBodyRange')
    except VBARuntimeError:
        return None
    if isinstance(body, Cells) and any(area.contains(at.top, at.left) for area in body.areas):
        return table
    return None


def read_value_detail() -> None:
    """Reading ShowDetail on a value is error 1004 (tests/fixtures/pivot_drill.json)."""
    raise error(1004, _UNABLE_GET)


def drill(table: PivotTable, target: Range, value: object, *, select: bool) -> None:
    """ShowDetail on a value of ``table`` (tests/fixtures/pivot_drill.json), or Application.DoubleClick on one.

    Excel selects ``target`` first, as ShowDetail is set through it. False then does nothing. With EnableDrilldown
    off, True is error 1004. Otherwise a new sheet goes in before the active one, as Worksheets.Add puts one, and
    takes the records behind the value as a table at A1: the cache fields' names over the records' values, every cell
    General, in the workbook's default table style, the columns left at the width they had; the table's range is
    selected. The grand total's records, all a table counts, come grouped by the items of the row fields, then those
    of the column fields, in the order the table shows them, each group in the source's order.

    A part of the table's records, one item's or one line's, comes in an order not worked out (source order where
    the item is the first record's, else not), so its detail is not built. Nor are tables with a filter, hidden or
    collapsed items, grouped or calculated fields, or records holding dates or errors."""
    from pyopenvba.apps.excel._model import Range as Cells

    entry = table.entry
    sheet = target.sheet
    if sheet.book.active_sheet is not sheet:
        raise VBAUnsupportedError('ShowDetail on a pivot table value of a sheet that is not active is not implemented')
    if not isinstance(value, bool):
        raise VBAUnsupportedError('ShowDetail on a pivot table value with a value other than True or False is not '
                                  'implemented')
    if select:
        target.select_as_excel_does()
    if not value:
        return
    if entry.root.get('enableDrill', '1') in ('0', 'false'):
        raise error(1004, _UNABLE_SET)
    _refuse_a_part(target)
    records = _detail_records(entry)
    style = sheet.book.stylesheet.table_style_default('table')
    if style is None:
        raise VBAUnsupportedError('ShowDetail in a workbook with no default table style is not implemented')
    names = [decode_text(field.get('name', '')) for field in
             entry.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')]
    added = _detail_sheet(sheet)
    for column, name in enumerate(names, start=1):
        _put(added, 1, column, name)
    for row, values in enumerate(records, start=2):
        for column, cached in enumerate(values, start=1):
            if cached is not EMPTY:
                _put(added, row, column, cached)
    added.touched()
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._tables import YES, add

    area = Cells(added, [Area(1, 1, len(records) + 1, len(names), added.name)])
    add(added, area, YES, style)
    area.select_as_excel_does()


def _detail_sheet(sheet: Worksheet) -> Worksheet:
    """The detail's sheet: before the active one, as Worksheets.Add puts one and with its events, but named Sheet and
    the lowest number no sheet has, whatever number the workbook gave last (tests/fixtures/pivot_drill.json)."""
    from pyopenvba.apps.excel import _events

    book = sheet.book
    if book.chart_sheets:
        raise VBAUnsupportedError('ShowDetail on a pivot table value in a workbook with chart sheets is not '
                                  'implemented')
    taken = {one.name.casefold() for one in book.tabs()}
    number = 1
    while f'sheet{number}' in taken:
        number += 1
    left = book.active_sheet
    added = book.add_sheet(f'Sheet{number}', at=book.sheets_.index(left) if left is not None else None,
                           numbered=True)
    book.activate_sheet(added)
    _events.new_sheet(book, added, left)
    book.activate_sheet(added)
    return added


def _put(sheet: Worksheet, row: int, column: int, value: object) -> None:
    """A cell of the detail, its value as the cache holds it: no typing, so text stays text."""
    cell = sheet.cell(row, column, create=True)
    assert cell is not None
    cell.value = value


def _refuse_a_part(target: Range) -> None:
    """Refuse a value whose records are a part of the table's, whose order is not worked out."""
    from pyopenvba.apps.excel._pivot_cells import owner

    at = target.first
    found = owner(target.sheet, at.top, at.left)
    assert found is not None
    role = found[1]
    if (role.rows or ()) or (role.columns or ()):
        raise VBAUnsupportedError("ShowDetail on a pivot table value standing for part of its records, which Excel "
                                  "lists in an order not worked out, is not implemented")


def _detail_records(table: PivotTable) -> list[list[object]]:
    """The records a table counts, in the order its grand total's detail lists them, each with its values."""
    from pyopenvba.apps.excel._pivot_aggregate import records
    from pyopenvba.apps.excel._pivot_layout import cache_layout
    from pyopenvba.apps.excel._pivot_reports import PivotLine, axis_lines

    root = table.root
    if root.find(NS + 'filters') is not None or root.find(NS + 'cacheHierarchies') is not None:
        raise VBAUnsupportedError('ShowDetail on a value of a pivot table with filters is not implemented')
    cache_fields = table.cache.root.findall(f'{NS}cacheFields/{NS}cacheField')
    if any(field.get('formula') is not None or field.find(NS + 'fieldGroup') is not None for field in cache_fields):
        raise VBAUnsupportedError('ShowDetail on a value of a pivot table with grouped or calculated fields is not '
                                  'implemented')
    pivot_fields = root.findall(f'{NS}pivotFields/{NS}pivotField')
    for field in pivot_fields:
        for item in field.findall(f'{NS}items/{NS}item'):
            if item.get('h') in ('1', 'true') or item.get('sd') in ('0', 'false'):
                raise VBAUnsupportedError('ShowDetail on a value of a pivot table with hidden or collapsed items is '
                                          'not implemented')
    found = records(table.cache)
    _, _, ids, _ = cache_layout(table.cache, found)
    rows = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}rowFields/{NS}field')]
    columns = [int(entry.get('x', '0')) for entry in root.findall(f'{NS}colFields/{NS}field')]
    rows, columns = [one for one in rows if one >= 0], [one for one in columns if one >= 0]

    def places(index: int) -> dict[int, int]:
        """A field's items' places in the table, by the cache item each stands for."""
        items = pivot_fields[index].findall(f'{NS}items/{NS}item')
        return {int(item.get('x', '-1')): place for place, item in enumerate(items) if item.get('t') is None}

    pages: dict[int, int] = {}
    for page in root.findall(f'{NS}pageFields/{NS}pageField'):
        chosen = page.get('item')
        if chosen is not None:
            items = pivot_fields[int(page.get('fld', '0'))].findall(f'{NS}items/{NS}item')
            pages[int(page.get('fld', '0'))] = int(items[int(chosen)].get('x', '-1'))
    row_places, column_places = {one: places(one) for one in rows}, {one: places(one) for one in columns}

    def order(lines: tuple[PivotLine, ...], axis: list[int], at: dict[int, dict[int, int]], record: list[int]) -> int:
        if not axis:
            return 0
        wanted = tuple(at[field].get(record[field], -1) for field in axis)
        for number, line in enumerate(lines):
            if line.kind == 'data' and line.items[:len(axis)] == wanted:
                return number
        raise VBAUnsupportedError('ShowDetail on a value of a pivot table whose records its lines do not show is not '
                                  'implemented')

    row_lines, column_lines = axis_lines(root, 'rowItems'), axis_lines(root, 'colItems')
    kept: list[tuple[tuple[int, int, int], list[object]]] = []
    for number, (record, values) in enumerate(zip(ids, found, strict=True)):
        if any(record[field] != item for field, item in pages.items()):
            continue
        key = (order(row_lines, rows, row_places, record), order(column_lines, columns, column_places, record), number)
        kept.append((key, [_cached(value) for value in values]))
    kept.sort(key=lambda pair: pair[0])
    return [values for _, values in kept]


def _cached(entry: XmlElement) -> object:
    """A cached value as the detail sheet's cell takes it."""
    kind = entry.tag[len(NS):]
    if kind == 's':
        return decode_text(entry.get('v', ''))
    if kind == 'n':
        return float(entry.get('v', '0'))
    if kind == 'b':
        return entry.get('v') in ('1', 'true')
    if kind == 'm':
        return EMPTY
    raise VBAUnsupportedError('ShowDetail on a value of a pivot table whose records hold dates or errors is not '
                              'implemented')
