"""Range.Insert and Range.Delete of cells that meet a table, and cells that move a table along, as live Excel takes
them (scripts/measure_table_shifts.py, tests/fixtures/table_shifts*.json).

Cells put in or taken away inside a table become table rows or columns:

- an Insert down whose cells start in a data row and span the table's columns puts that many table rows in, and one
  with no Shift whose cells lie in its data rows does it whatever their width;
- a Delete up whose cells lie in its data rows, in any of its columns, takes those whole table rows away;
- an Insert right whose cells span the table's rows and start inside it puts table columns in, and a Delete left
  whose cells lie in its rows and columns takes those whole table columns away.

Cells moving along take a table whole with them where it lies wholly in their band, beyond where they start; a
table only some of whose rows or columns the band crosses, cells put in that cut a table's width, and a Delete that
runs past a table's last row or its edge are error 1004, as Excel has it, "This won't work because it would move
cells in a table on your worksheet". A Delete that takes a table's header row and not the whole table is 1004 too.
In a filtered table a Delete up is 1004, "No cells were found." where every row it takes is hidden.
"""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._editing import CellShift
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.apps.excel._tables import Table

MOVES_A_TABLE = "This won't work because it would move cells in a table on your worksheet."
FILTERED = "Can't move cells in a filtered range or table."


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def _inside(inner: Area, outer: Area) -> bool:
    return outer.top <= inner.top and inner.bottom <= outer.bottom and outer.left <= inner.left \
        and inner.right <= outer.right


def _filtered(table: Table) -> bool:
    """Whether the table's filter has criteria and the sheet is in filter mode by them."""
    from pyopenvba.apps.excel._table_filters import filter_mode, has_criteria

    return has_criteria(table) and filter_mode(table.sheet)


def carried(sheet: Worksheet, area: Area, *, delete: bool, vertical: bool, auto: bool,
            refusal: str) -> Callable[[], None] | None:
    """What an Insert or a Delete of ``area`` does to the table it meets, checked now and carried out by what this
    returns; None where it meets none, or takes a whole table away, which the cell shift itself carries."""
    met = [table for table in sheet.tables if _meets(table.area, area)]

    def whole_moves(table: Table) -> bool:
        """Whether the edit takes the table whole, or moves it whole with the cells after it."""
        found = table.area
        if delete:
            return _inside(found, area)
        if vertical:
            return area.top <= found.top and area.left <= found.left and found.right <= area.right
        return area.left <= found.left and area.top <= found.top and found.bottom <= area.bottom

    if all(whole_moves(table) for table in met):
        return None
    if len(met) > 1:
        raise VBAUnsupportedError('cells put in or taken away across several tables are not implemented')
    table = met[0]
    whole, data_top = table.area, table.area.top + table.headers
    if table.totals:
        raise VBAUnsupportedError('cells put in or taken away in a table with a totals row are not implemented')
    if delete and _inside(whole, area):
        return None
    if vertical and delete:
        if table.headers and area.top <= whole.top:
            raise error(1004, refusal)
        if not (area.bottom <= whole.bottom and whole.left <= area.left and area.right <= whole.right):
            raise error(1004, MOVES_A_TABLE)
        if _filtered(table):
            hidden = all(sheet.dims.row_hidden(row) for row in range(area.top, area.bottom + 1))
            raise error(1004, 'No cells were found.' if hidden else FILTERED)
        if area.top == data_top and area.bottom == whole.bottom:
            raise VBAUnsupportedError("deleting all of a table's rows of data is not implemented")
        return lambda: _rows_gone(table, area.top, area.rows)
    if vertical:
        if area.top <= whole.top and whole.left >= area.left and whole.right <= area.right:
            return None
        spans = (area.left, area.right) == (whole.left, whole.right) or (
            auto and whole.left <= area.left and area.right <= whole.right and area.bottom <= whole.bottom)
        if data_top <= area.top <= whole.bottom and spans:
            return lambda: _rows_in(table, area.top, area.rows)
        raise error(1004, MOVES_A_TABLE)
    if auto:
        raise VBAUnsupportedError('cells put in or taken away across a table with no Shift are not implemented')
    if delete:
        if not _inside(area, whole):
            raise VBAUnsupportedError('deleting cells left past the edge of a table is not implemented')
        if (area.left, area.right) == (whole.left, whole.right):
            raise VBAUnsupportedError("deleting all of a table's columns by deleting cells is not implemented")
        return lambda: _columns_gone(table, area.left, area.columns)
    if area.top <= whole.top and whole.bottom <= area.bottom and area.left <= whole.left:
        return None
    if (area.top, area.bottom) == (whole.top, whole.bottom) and whole.left < area.left <= whole.right:
        return lambda: _columns_in(table, area.left, area.columns)
    if area.top < whole.top or whole.bottom < area.bottom:
        raise VBAUnsupportedError('putting cells in across a table and beyond its rows is not implemented')
    raise error(1004, MOVES_A_TABLE)


def _rows_gone(table: Table, top: int, count: int) -> None:
    from pyopenvba.apps.excel._table_edits import delete_row

    for row in range(top + count - 1, top - 1, -1):
        delete_row(table, row - table.area.top - table.headers + 1)


def _rows_in(table: Table, top: int, count: int) -> None:
    from pyopenvba.apps.excel._table_edits import add_row

    position = top - table.area.top - table.headers + 1
    for _ in range(count):
        add_row(table, position if position <= table.rows() else None, True)


def _columns_gone(table: Table, left: int, count: int) -> None:
    from pyopenvba.apps.excel._table_edits import delete_column

    for column in range(left + count - 1, left - 1, -1):
        delete_column(table, column - table.area.left + 1)


def _columns_in(table: Table, left: int, count: int) -> None:
    from pyopenvba.apps.excel._table_edits import add_column

    for _ in range(count):
        add_column(table, left - table.area.left + 1)


def shifted(sheet: Worksheet, shift: CellShift, region: Area, through: Table | None) -> Callable[[], None]:
    """The tables cells moving along ``region`` carry with them, as ``shift`` moves them, but ``through``, which
    its caller keeps; checked now, carried out by what this returns. One the band crosses only in part, or whose
    rows or columns the edit itself puts cells into or takes them from, is 1004."""
    vertical, band, start, count, delete = shift.vertical, shift.band, shift.start, shift.count, shift.delete
    moving: list[Table] = []
    for table in sheet.tables:
        if table is through or not _meets(table.area, region):
            continue
        area = table.area
        across = (area.left, area.right) if vertical else (area.top, area.bottom)
        along = area.top if vertical else area.left
        if band[0] <= across[0] and across[1] <= band[1] and along >= start + (count if delete else 0):
            moving.append(table)
            continue
        raise error(1004, MOVES_A_TABLE)

    def carry_out() -> None:
        step = -count if delete else count
        for table in moving:
            area = table.area
            table.area = (Area(area.top + step, area.left, area.bottom + step, area.right) if vertical
                          else Area(area.top, area.left + step, area.bottom, area.right + step))
            table.changed = True
    return carry_out
