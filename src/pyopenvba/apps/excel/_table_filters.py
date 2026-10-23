"""A table's own AutoFilter, which Excel keeps apart from the sheet's.

Measured in live Excel (scripts/measure_table_autofilter.py,
tests/fixtures/table_autofilter/):

- Range.AutoFilter on a range inside a table -- the table's range, a cell
  or part of it, its data or one of its columns -- filters the table, the
  field counted from the table's first column whatever the range. A range
  reaching past the table's edge, whole columns over it among them, is
  error 1004. With no field it turns the table's filter off, showing its
  rows, or on again, and with a field it turns the filter on first. The
  criteria are a sheet filter's, read back and saved the same way, in the
  table's part; the totals row is never hidden, and no _FilterDatabase
  name is made.
- ListObject.AutoFilter is the table's AutoFilter, Nothing while
  ShowAutoFilter is False, and its Parent is the ListObject. Its Range is
  the whole table, totals row included; a table read from a file answers
  the block the file gave, which leaves the totals row out, until the
  table or its filter changes. ShowAllData clears it, with no error when
  nothing is set, and ShowAutoFilter = False takes the criteria away and
  shows the rows.
- The sheet's own members reach the filter at the active cell: a table's
  where the active cell is in a table, else the sheet's own.
  Worksheet.AutoFilter answers it, or Nothing; ShowAllData clears it, and
  is error 1004 when it has no criteria; AutoFilterMode = False turns it
  off, but only on a sheet with a filter of its own, the one
  AutoFilterMode answers for.
- FilterMode is that filter's as it was when the sheet was last
  activated, its selection moved or a filter changed. Deleting a filtered
  table leaves it True, and a sheet read from a file answers for its own
  filter until then, since the file keeps no word of a table's. Edits hold
  back from hidden rows by FilterMode, as on a sheet filtered by itself,
  while SUBTOTAL passes over every hidden row once any filter on the sheet,
  the sheet's or a table's, has criteria.
- A table made over the sheet's filtered range takes the filter over,
  criteria and all, leaving the hidden name. Resizing a filtered table
  shows the rows it leaves; the row ListRows.Add puts in shows; the
  criteria move with their columns when ListColumns.Add puts one in; and
  deleting the table shows its rows.

What was not measured reports itself: filtering a table with no header
row, which shows the header row again; rows put in or taken out inside a
filtered table; deleting a column with criteria; and resizing a filtered
table wider or longer.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._autofilter import FieldFilter, SheetFilter
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet
    from pyopenvba.apps.excel._tables import Table


def filtered_block(table: Table) -> Area:
    """What a table's filter covers: the header row and the data rows, never the totals row."""
    area = table.area
    return Area(area.top, area.left, max(area.top, area.bottom - (1 if table.totals else 0)), area.right)


def table_filter(table: Table) -> SheetFilter | None:
    """The table's AutoFilter, None where it shows none."""
    if not table.auto_filter or not table.headers:
        return None
    if table.filter is None:
        table.filter = SheetFilter(filtered_block(table), uid=table.uid)
    table.filter.area = filtered_block(table)
    return table.filter


def has_criteria(table: Table) -> bool:
    found = table_filter(table)
    return found is not None and bool(found.fields)


def filter_range(table: Table) -> Area:
    """AutoFilter.Range of a table: the whole of it, but the block its file gave while it is as it was read."""
    kept = table.filter_ref
    if kept is not None and kept[1] == table.area:
        return kept[0]
    return table.area


def changed(table: Table) -> None:
    """A change to a table's filter, which its part has to say, and which ends the block its file gave."""
    table.changed = True
    table.filter_ref = None
    table.sheet.touched()


def table_of(target: Range) -> Table | None:
    """The table a filter called on ``target`` reaches, the one the range lies in; None for a range clear of
    every table. A range reaching past a table's edge is error 1004."""
    met = [table for table in target.sheet.tables if any(_meets(table.area, area) for area in target.areas)]
    if not met:
        return None
    within = all(met[0].area.top <= area.top and area.bottom <= met[0].area.bottom and met[0].area.left <= area.left
                 and area.right <= met[0].area.right for area in target.areas)
    if len(met) > 1 or not within:
        raise error(1004, "AutoFilter method of Range class failed")
    return met[0]


def meets_table(sheet: Worksheet, area: Area) -> bool:
    return any(_meets(table.area, area) for table in sheet.tables)


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def turn_off(table: Table) -> None:
    """The table's filter turned off: its criteria gone and its rows shown."""
    from pyopenvba.apps.excel import _events

    found = table_filter(table)
    table.auto_filter = False
    table.filter = None
    if found is not None:
        found.show_all(table.sheet)
        if found.fields:
            _events.visibility_changed(table.sheet.book)
    changed(table)


def clear(table: Table) -> None:
    """The table's criteria taken away and its rows shown, as ShowAllData does; nothing where none is set."""
    from pyopenvba.apps.excel import _events

    found = table_filter(table)
    if found is None or not found.fields:
        return
    found.fields.clear()
    found.show_all(table.sheet)
    changed(table)
    refocus(table.sheet)
    _events.visibility_changed(table.sheet.book)


# --- the filter the sheet's own members reach ------------------------------------------------------------


def context_table(sheet: Worksheet) -> Table | None:
    """The table the sheet's active cell is in, whose filter the sheet's members reach rather than the sheet's."""
    from pyopenvba.apps.excel._tables import table_at

    active = sheet.active_cell_range
    row, column = (active.first.top, active.first.left) if active is not None else (1, 1)
    return table_at(sheet, row, column)


def refocus(sheet: Worksheet) -> None:
    """FilterMode taken from the filter at the active cell, as Excel takes it when the sheet is activated, its
    selection moves or a filter changes: None while that is the sheet's own, which FilterMode reads as it is."""
    table = context_table(sheet)
    sheet.table_filter_mode = None if table is None else has_criteria(table)


def filter_mode(sheet: Worksheet) -> bool:
    """Worksheet.FilterMode, which edits hold back from hidden rows by."""
    if sheet.advanced_filter is not None:
        return True
    if sheet.table_filter_mode is not None:
        return sheet.table_filter_mode
    own = sheet.auto_filter
    return own is not None and bool(own.fields)


def criteria_anywhere(sheet: Worksheet) -> bool:
    """Whether a filter on the sheet, its own or a table's, has criteria, which SUBTOTAL takes every hidden row for."""
    own = sheet.auto_filter
    return (own is not None and bool(own.fields)) or any(
        table.auto_filter and table.filter is not None and bool(table.filter.fields) for table in sheet.tables)


# --- the table changing under its filter ------------------------------------------------------------------


def take_over(sheet: Worksheet, area: Area) -> dict[int, FieldFilter] | None:
    """Before ListObjects.Add over ``area``: the criteria of the sheet's filter over exactly that range, which the
    table takes over, the sheet's filter going; None where the sheet's filter is clear of it or there is none. A
    filter over part of it reports itself."""
    own = sheet.auto_filter
    if own is None or not _meets(own.area, area):
        return None
    if (own.area.top, own.area.left, own.area.bottom, own.area.right) != (area.top, area.left, area.bottom, area.right):
        raise VBAUnsupportedError("a table over part of the sheet's filtered range is not implemented")
    return dict(own.fields)


def taken_over(table: Table, fields: dict[int, FieldFilter]) -> None:
    """ListObjects.Add over the sheet's filtered range: the table filters by its criteria, the sheet's filter
    gone; its hidden name stays."""
    sheet = table.sheet
    found = table_filter(table)
    assert found is not None
    found.fields = fields
    sheet.auto_filter = None
    sheet.filter_changed = True
    changed(table)
    refocus(sheet)


def refuse_filtered(table: Table, doing: str) -> None:
    """Report an edit of a filtered table that was not measured."""
    if table.auto_filter and table.filter is not None and table.filter.fields:
        raise VBAUnsupportedError(f"{doing} a filtered table is not implemented")


def column_added(table: Table, index: int) -> None:
    """ListColumns.Add put a column in at ``index``: the criteria move with their columns."""
    found = table.filter
    if found is not None and found.fields:
        found.fields = {(at + 1 if at >= index else at): one for at, one in found.fields.items()}


def column_deleted(table: Table, index: int) -> None:
    """Before ListColumn.Delete takes column ``index`` out: the criteria right of it move with their columns."""
    found = table.filter
    if found is None or not found.fields:
        return
    if index in found.fields:
        raise VBAUnsupportedError("deleting a column a table is filtered by is not implemented")
    found.fields = {(at - 1 if at > index else at): one for at, one in found.fields.items()}


def resized(table: Table, area: Area) -> None:
    """Before a filtered table is resized to ``area``: the rows it leaves show; any other change reports itself."""
    found = table.filter
    if not table.auto_filter or found is None or not found.fields:
        return
    old = filtered_block(table)
    if area.columns != table.area.columns or area.bottom > table.area.bottom:
        raise VBAUnsupportedError("resizing a filtered table wider or longer is not implemented")
    sheet = table.sheet
    for row in range(area.bottom + 1, old.bottom + 1):
        if sheet.dims.row_hidden(row):
            sheet.dims.hide_row(row, False)


def deleted(table: Table) -> None:
    """Before a table is deleted: the rows its filter hid show. FilterMode stays as it was, as Excel leaves it."""
    found = table.filter
    if table.auto_filter and found is not None and found.fields:
        from pyopenvba.apps.excel import _events

        found.area = filtered_block(table)
        found.show_all(table.sheet)
        _events.visibility_changed(table.sheet.book)
