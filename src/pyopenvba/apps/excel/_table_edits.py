"""Editing a table's rows and columns: ListRows, ListColumns, Resize, Delete, calculated columns and growth.

Measured in live Excel (scripts/measure_table_edits.py, tests/fixtures/tables/edits/;
scripts/measure_table_growth.py, tests/fixtures/tables/growth.json):

* ListRows.Add puts a row in at the end, above the totals row if there is
  one, or at a position; the table's columns move down under it, the rest
  of the sheet stays. With AlwaysInsert False and the row under the table
  empty, the table takes that row instead. ListRow.Delete takes a row out,
  the cells under it in the table's columns moving up.
* ListColumns.Add puts a column in at the end or at a position, named the
  first free ColumnN, the cells right of it in the table's rows moving
  across; ListColumn.Delete takes one out, and a structured reference to
  it becomes #REF!.
* Resize lays the table over another block with the same header row; a
  column it leaves out is #REF! wherever named, a new one takes its name
  from its header cell. Delete clears the table's cells, formats and all,
  and every reference to it becomes #REF!.
* A formula written into a table column that is otherwise empty fills the
  column, and so does one written over the whole of it: the column is
  calculated, and a row the table gains takes the formula.
* A value or a formula a macro writes into the row just under a table, or
  the column just right of it, takes the table over it -- not under a
  totals row. A value written under it also stretches each reference
  running down a column of the table from its header or first row to its
  last, so SUM(C2:C4) reads SUM(C2:C5).
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_number
from pyopenvba.apps.excel._tables import (Table, TableColumn, cell_text, column_names, guid, name_columns,
                                          rewrite_formulas, table_at)
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import FormulaError, Token, read_structured, split_sheet, tokenize
from pyopenvba.formula._structured import rewritten
from pyopenvba.interpreter._values import EMPTY, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

_CORNER = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$")


# --- rows ------------------------------------------------------------------------------------------


def add_row(table: Table, position: int | None, always_insert: bool) -> int:
    """ListRows.Add: the index of the row put in."""
    from pyopenvba.apps.excel._editing import shift_cells
    from pyopenvba.apps.excel._model import Range

    sheet, area, count = table.sheet, table.area, table.rows()
    data_top = area.top + table.headers
    if position is None or position == count + 1:
        index, row = count + 1, data_top + count
    elif 1 <= position <= count:
        index, row = position, data_top + position - 1
    else:
        raise error(9, "Subscript out of range")
    at_end = index == count + 1
    below = Area(row, area.left, row, area.right, sheet.name)
    if at_end and not table.totals and not always_insert and _empty(sheet, below):
        pass
    else:
        shift_cells(Range(sheet, [below]), below, delete=False, vertical=True, through=table)
    table.area = Area(area.top, area.left, area.bottom + 1, area.right)
    if at_end:
        _stretch(table, row - 1)
    _calculate_row(table, row)
    _changed(table)
    return index


def delete_row(table: Table, index: int) -> None:
    """ListRow.Delete: the row's cells in the table's columns deleted, those under them moving up."""
    from pyopenvba.apps.excel._editing import shift_cells
    from pyopenvba.apps.excel._model import Range

    sheet, area = table.sheet, table.area
    if table.rows() <= 1:
        raise VBAUnsupportedError("deleting a table's last row of data is not implemented")
    row = area.top + table.headers + index - 1
    gone = Area(row, area.left, row, area.right, sheet.name)
    table.area = Area(area.top, area.left, area.bottom - 1, area.right)
    shift_cells(Range(sheet, [gone]), gone, delete=True, vertical=True, through=table)
    _changed(table)


def _empty(sheet: Worksheet, area: Area) -> bool:
    return not any(_filled(sheet, row, column)
                   for row in range(area.top, area.bottom + 1) for column in range(area.left, area.right + 1))


def _filled(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether a cell holds something: a formula or a value, an empty string not counting."""
    cell = sheet.cell(row, column)
    return cell is not None and (bool(cell.formula) or cell.value not in (EMPTY, ""))


def _calculate_row(table: Table, row: int) -> None:
    """A row the table gained takes each calculated column's formula."""
    for offset, column in enumerate(table.columns):
        if column.calculated:
            cell = table.sheet.cell(row, table.area.left + offset, create=True)
            assert cell is not None
            cell.value, cell.formula, cell.stale, cell.shared = EMPTY, column.calculated, True, None
            table.sheet.cell_changed(row, table.area.left + offset)


def _stretch(table: Table, last: int) -> None:
    """A table grown down past ``last``, its old last row: each reference running down columns of the table from
    its header or first row to ``last`` runs on to the new last row."""
    sheet, area = table.sheet, table.area
    data_top, bottom = area.top + table.headers, area.bottom - table.totals

    def change(text: str, owner: str) -> str:
        try:
            tokens = tokenize(text.removeprefix("="))
        except FormulaError:
            return text
        pieces: list[tuple[int, int, str]] = []
        for token in tokens:
            if token.kind != "ref":
                continue
            named, reference = split_sheet(token.text)
            corners = [_CORNER.fullmatch(part) for part in reference.split(":")]
            if (named or owner).lower() != sheet.name.lower() or len(corners) != 2 or None in corners:
                continue
            (first, second) = [corner.groups() for corner in corners if corner is not None]
            left, right = sorted((column_number(first[1]), column_number(second[1])))
            top, low = sorted((int(first[3]), int(second[3])))
            if area.left <= left and right <= area.right and top in (area.top, data_top) and low == last:
                lower = second if int(second[3]) == low else first
                moved = lower[0] + lower[1] + lower[2] + str(bottom)
                parts = [moved if group is lower else "".join(group) for group in (first, second)]
                pieces.append((token.at, token.at + len(token.text), token.text[:-len(reference)] + ":".join(parts)))
        body = text.removeprefix("=")
        for start, stop, replacement in reversed(pieces):
            body = body[:start] + replacement + body[stop:]
        return ("=" if text.startswith("=") else "") + body

    for owner in sheet.book.sheets_:
        for cell in owner.cells_.values():
            if cell.formula:
                updated = change(cell.formula, owner.name)
                if updated != cell.formula:
                    cell.formula = updated
                    owner.touched()
    for entry in sheet.book.names_.entries:
        scope, _ = split_sheet(entry.name)
        updated = change(entry.refers_to, scope or sheet.name)
        if updated != entry.refers_to:
            entry.refers_to = updated
            sheet.book.names_.changed = True
    sheet.shape_changed()


# --- columns ---------------------------------------------------------------------------------------


def add_column(table: Table, position: int | None) -> int:
    """ListColumns.Add: the index of the column put in, named the first free ColumnN."""
    from pyopenvba.apps.excel._editing import shift_cells
    from pyopenvba.apps.excel._model import Range

    sheet, area, count = table.sheet, table.area, len(table.columns)
    index = count + 1 if position is None else position
    if not 1 <= index <= count + 1:
        raise error(9, "Subscript out of range")
    column = area.left + index - 1
    band = Area(area.top, column, area.bottom, column, sheet.name)
    shift_cells(Range(sheet, [band]), band, delete=False, vertical=False, through=table)
    table.area = Area(area.top, area.left, area.bottom, area.right + 1)
    names = [one.name for one in table.columns]
    name = column_names([*names, ""])[-1]
    table.columns.insert(index - 1, TableColumn(name=name, id=max(one.id for one in table.columns) + 1, uid=guid()))
    name_columns(table, [one.name for one in table.columns])
    _changed(table)
    return index


def delete_column(table: Table, index: int) -> None:
    """ListColumn.Delete: the column's cells deleted, those right of them in the table's rows moving across, and
    every structured reference to it #REF!."""
    from pyopenvba.apps.excel._editing import shift_cells
    from pyopenvba.apps.excel._model import Range

    sheet, area = table.sheet, table.area
    if len(table.columns) <= 1:
        raise VBAUnsupportedError("deleting a table's last column is not implemented")
    gone = table.columns[index - 1]
    column = area.left + index - 1
    band = Area(area.top, column, area.bottom, column, sheet.name)
    table.area = Area(area.top, area.left, area.bottom, area.right - 1)
    del table.columns[index - 1]
    shift_cells(Range(sheet, [band]), band, delete=True, vertical=False, through=table)
    _lose_columns(table, {gone.name.lower()})
    _changed(table)


def _lose_columns(table: Table, names: set[str]) -> None:
    """Every structured reference naming one of these columns of the table, now gone, becomes #REF!."""
    wanted = table.name.lower()

    def one(token: Token) -> str | None:
        if token.kind != "structured":
            return None
        node = read_structured(token.text)
        if node.table.lower() != wanted or node.first is None:
            return None
        return "#REF!" if {node.first.lower(), (node.last or node.first).lower()} & names else None

    rewrite_formulas(table.sheet.book, lambda text: rewritten(text, one))


# --- the whole table -------------------------------------------------------------------------------


def resize(table: Table, area: Area) -> None:
    """ListObject.Resize: the table over another block, the same header row. Measured: a block starting on
    another row is error 1004."""
    if area.top != table.area.top:
        raise error(1004, "The table's header row cannot move")
    if area.left != table.area.left or table.totals or not table.headers:
        raise VBAUnsupportedError("resizing a table from its left edge, or one with a totals row or no header "
                                  "row, is not implemented")
    if any(other is not table and other.area.top <= area.bottom and area.top <= other.area.bottom
           and other.area.left <= area.right and area.left <= other.area.right for other in table.sheet.tables):
        raise error(1004, "A table cannot overlap another table.")
    kept = min(area.columns, len(table.columns))
    lost = {one.name.lower() for one in table.columns[kept:]}
    table.columns = table.columns[:kept]
    table.area = Area(area.top, area.left, area.bottom, area.right)
    for _ in range(area.columns - kept):
        table.columns.append(TableColumn(name="", id=max(one.id for one in table.columns) + 1, uid=guid()))
    if len(table.columns) > kept:
        texts = [cell_text(table.sheet, area.top, area.left + offset) for offset in range(area.columns)]
        name_columns(table, column_names(texts))
    if lost:
        _lose_columns(table, lost)
    _changed(table)


def delete_table(table: Table) -> None:
    """ListObject.Delete: the table's cells cleared, formats and all, and every reference to it #REF!."""
    sheet, area = table.sheet, table.area
    wanted = table.name.lower()

    def one(token: Token) -> str | None:
        if token.kind == "structured":
            return "#REF!" if read_structured(token.text).table.lower() == wanted else None
        return "#REF!" if token.kind == "name" and token.text.lower() == wanted else None

    sheet.tables.remove(table)
    for row in range(area.top, area.bottom + 1):
        for column in range(area.left, area.right + 1):
            if sheet.cells_.pop((row, column), None) is not None:
                sheet.cell_changed(row, column)
    rewrite_formulas(sheet.book, lambda text: rewritten(text, one))
    sheet.touched()


def _changed(table: Table) -> None:
    table.changed = True
    table.sheet.touched()
    table.sheet.shape_changed()


# --- calculated columns ----------------------------------------------------------------------------


def calculated(sheet: Worksheet, written: list[Area]) -> None:
    """Formulas a macro wrote into a table: one written into a column otherwise empty fills the column, and one
    over all of it makes it calculated too."""
    for area in written:
        table = table_at(sheet, area.top, area.left)
        data = table.data if table is not None else None
        if table is None or data is None:
            continue
        for column in range(max(area.left, data.left), min(area.right, data.right) + 1):
            entry = table.columns[column - table.area.left]
            rows = range(data.top, data.bottom + 1)
            cells = {row: sheet.cell(row, column) for row in rows}
            formulas = {row: cell.formula for row, cell in cells.items() if cell is not None and cell.formula}
            filled = {row for row in rows if _filled(sheet, row, column)}
            if len(formulas) == 1 and filled == set(formulas) and area.top <= min(formulas) <= area.bottom:
                # One formula in a column otherwise empty: the column fills with it.
                formula = formulas[min(formulas)]
                for row in rows:
                    if row not in formulas:
                        cell = sheet.cell(row, column, create=True)
                        assert cell is not None
                        cell.value, cell.formula, cell.stale, cell.shared = EMPTY, formula, True, None
                        sheet.cell_changed(row, column)
                entry.calculated = formula
                _changed(table)
            elif len(set(formulas.values())) == 1 and len(formulas) == len(rows) and area.top <= data.top \
                    and data.bottom <= area.bottom:
                # One formula written over all of the column, the same in every row.
                entry.calculated = formulas[data.top]
                _changed(table)


# --- growing -------------------------------------------------------------------------------------


def grow(sheet: Worksheet, written: list[Area]) -> None:
    """A macro wrote into ``written``: a table takes the row under it, or the column right of it, where a value
    or a formula landed there, as far as the writing goes; a value stretches references down with it, and a
    formula alone does not (tests/fixtures/tables/growth.json)."""
    for table in list(sheet.tables):
        # References are stretched once, from the last row the table had to the last it has: SUM(C2:C5) stays
        # as it is when a write takes a table from row 4 to row 6.
        last, values = table.area.bottom, False
        while not table.totals and (landed := _landed(sheet, written, table.area.bottom + 1, table.area.left,
                                                      table.area.right, rows=True)):
            table.area = Area(table.area.top, table.area.left, table.area.bottom + 1, table.area.right)
            values = values or landed == "value"
            _calculate_row(table, table.area.bottom)
            _changed(table)
        if values:
            _stretch(table, last)
        while _landed(sheet, written, table.area.right + 1, table.area.top, table.area.bottom, rows=False):
            table.area = Area(table.area.top, table.area.left, table.area.bottom, table.area.right + 1)
            table.columns.append(TableColumn(name="", id=max(one.id for one in table.columns) + 1, uid=guid()))
            texts = [cell_text(sheet, table.area.top, column)
                     for column in range(table.area.left, table.area.right + 1)]
            name_columns(table, column_names(texts))
            _changed(table)


def _landed(sheet: Worksheet, written: list[Area], line: int, low: int, high: int, *, rows: bool) -> str:
    """What the writing put in row (or column) ``line`` between ``low`` and ``high``: "value" where a value is
    among it, "formula" where only formulas are, "" where nothing is."""
    if line > (MAX_ROWS if rows else MAX_COLUMNS):
        return ""
    found = ""
    for area in written:
        if rows and area.top <= line <= area.bottom:
            positions = [(line, column) for column in range(max(low, area.left), min(high, area.right) + 1)]
        elif not rows and area.left <= line <= area.right:
            positions = [(row, line) for row in range(max(low, area.top), min(high, area.bottom) + 1)]
        else:
            continue
        for position in positions:
            if _filled(sheet, *position):
                cell = sheet.cell(*position)
                if cell is not None and not cell.formula:
                    return "value"
                found = "formula"
    return found
