"""Refreshing a query: evaluating its M and landing the result on a sheet.

The evaluator in :mod:`pyopenvba.mlang` does not know what a workbook
is.  This is what connects the two: the other queries are in scope by
name, ``Excel.CurrentWorkbook()`` reads the tables and named ranges of
the workbook being refreshed, and what comes back is written into the
cells the query was loaded to.

A source that cannot be reached from here -- a database, a web service,
a folder on someone else's machine -- reports itself rather than
guessing, which is the same rule the rest of the runtime follows.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area, split_sheet
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang import MError, Record, Table, base_scope, evaluate, parse
from pyopenvba.mlang._eval import Scope, Thunk
from pyopenvba.mlang._values import Builtin, Duration, is_list
from pyopenvba.interpreter._values import EMPTY, VBADate

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet


@dataclass(slots=True)
class LoadTarget:
    """Where a query's rows go: the sheet, and the block they fill."""

    sheet: str
    area: Area
    table_part: str = ""
    table_name: str = ""
    #: The column names the last refresh wrote, so the table can be
    #: brought in line with them on save.
    headers: list[str] | None = None


def query_scope(book: Workbook) -> Scope:
    """Every name a query in this workbook can see."""
    scope = base_scope({"Excel.CurrentWorkbook": _current_workbook(book)})
    for entry in book.queries_.entries:
        scope.names[entry.name] = _QueryThunk(book, entry.name, scope)
    return scope


class _QueryThunk(Thunk):
    """One query, evaluated the first time another one names it."""

    def __init__(self, book: Workbook, name: str, scope: Scope) -> None:
        self.book = book
        self.query_name = name
        super().__init__(parse("null"), scope)

    def force(self) -> object:
        if self.state == "done":
            return self.value
        if self.state == "working":
            raise MError("Expression.Error", f"The query '{self.query_name}' refers to itself.")
        self.state = "working"
        for entry in self.book.queries_.entries:
            if entry.name == self.query_name:
                self.value = evaluate(parse(entry.formula), self.scope)
                break
        else:  # pragma: no cover - the thunk is only made from an entry
            self.value = None
        self.state = "done"
        return self.value


def _current_workbook(book: Workbook) -> Builtin:
    """``Excel.CurrentWorkbook()``: the workbook's own tables and names.

    Excel answers with a two-column table, Name and Content, one row per
    table or named range, with each Content a table of its own.
    """

    def call() -> object:
        rows: list[list[object]] = []
        for entry in book.names_.entries:
            if split_sheet(entry.name)[1].startswith("ExternalData_"):
                continue
            block = _named_block(book, entry.refers_to)
            if block is not None:
                rows.append([entry.name, block])
        for target in load_targets(book).items():
            name, where = target
            sheet = _sheet_named(book, where.sheet)
            if sheet is not None:
                rows.append([where.table_name or name, _block_of(sheet, where.area, headers=True)])
        return Table(["Name", "Content"], rows)

    return Builtin(name="Excel.CurrentWorkbook", call=call, minimum=0, maximum=0)


def _named_block(book: Workbook, refers_to: str) -> Table | None:
    text = refers_to.lstrip("=")
    try:
        area = parse_area(text)
    except ValueError:
        return None
    sheet = _sheet_named(book, area.sheet)
    if sheet is None:
        return None
    return _block_of(sheet, area, headers=True)


def _sheet_named(book: Workbook, name: str) -> Worksheet | None:
    for sheet in book.sheets_:
        if sheet.name.lower() == name.lower():
            return sheet
    return None


def _block_of(sheet: Worksheet, area: Area, *, headers: bool) -> Table:
    """A block of cells as an M table, its first row the column names."""
    rows: list[list[object]] = []
    for row in range(area.top, area.bottom + 1):
        line: list[object] = []
        for column in range(area.left, area.right + 1):
            line.append(_to_m(sheet, row, column))
        rows.append(line)
    if headers and rows:
        names = [str(one) if one is not None else f"Column{index + 1}" for index, one in enumerate(rows[0])]
        return Table(names, rows[1:])
    width = len(rows[0]) if rows else 0
    return Table([f"Column{index + 1}" for index in range(width)], rows)


def _to_m(sheet: Worksheet, row: int, column: int) -> object:
    """One cell as M sees it."""
    from pyopenvba.apps.excel._calc import as_vba
    from pyopenvba.formula._values import ExcelError

    cell = sheet.cell(row, column)
    if cell is None:
        return None
    value = cell.value
    if cell.formula:
        value = sheet.book.calculator.value_of(sheet.name, row, column)
    if isinstance(value, ExcelError):
        raise MError("Expression.Error", f"the cell holds {value.name}")
    value = as_vba(value, cell)
    if value is EMPTY:
        return None
    if isinstance(value, VBADate):
        return value.to_datetime()
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


# --- where a query's rows go ------------------------------------------------------------------


def load_targets(book: Workbook) -> dict[str, LoadTarget]:
    """Each query's block of cells, read from the package it came from."""
    if getattr(book, "_load_targets", None) is not None:
        return book._load_targets  # type: ignore[attr-defined]
    found: dict[str, LoadTarget] = {}
    package = book.package
    if package is not None:
        from pyopenvba.apps.excel._io import read_load_targets

        found = read_load_targets(package, book)
    book._load_targets = found  # type: ignore[attr-defined]
    return found


def refresh(book: Workbook, name: str) -> Table:
    """Evaluate one query and write what it answers onto its sheet."""
    scope = query_scope(book)
    for entry in book.queries_.entries:
        if entry.name.lower() == name.lower():
            formula = entry.formula
            break
    else:
        raise MError("Expression.Error", f"There is no query called '{name}'.")
    value = evaluate(parse(formula), scope)
    table = value if isinstance(value, Table) else _as_table(value)
    target = load_targets(book).get(name)
    if target is not None:
        write_table(book, target, table)
    return table


def _as_table(value: object) -> Table:
    if isinstance(value, Table):
        return value
    if is_list(value):
        if value and isinstance(value[0], Record):
            return Table.from_records([one for one in value if isinstance(one, Record)])
        return Table(["Column1"], [[one] for one in value])
    if isinstance(value, Record):
        return Table(value.names, [[value.get(name) for name in value.names]])
    return Table(["Value"], [[value]])


def write_table(book: Workbook, target: LoadTarget, table: Table) -> None:
    """Put a table's headers and rows into the cells it loads to."""
    sheet = _sheet_named(book, target.sheet)
    if sheet is None:
        raise MError("Expression.Error", f"There is no sheet called '{target.sheet}'.")
    top = target.area.top
    left = target.area.left
    for index, name in enumerate(table.columns):
        _put(sheet, top, left + index, name)
    for offset, row in enumerate(table.rows, start=1):
        for index in range(len(table.columns)):
            value = row[index] if index < len(row) else None
            _put(sheet, top + offset, left + index, value)
    # Anything the last result left behind goes, below the new rows and
    # beyond the new columns: a query that comes back narrower would
    # otherwise leave a column of the old one standing beside it.
    for row in range(top, target.area.bottom + 1):
        for column in range(left, target.area.right + 1):
            if row <= top + len(table.rows) and column < left + len(table.columns):
                continue
            sheet.cells_.pop((row, column), None)
    filled = Area(
        top,
        left,
        top + max(len(table.rows), 1),
        left + max(len(table.columns), 1) - 1,
        sheet.name,
    )
    book._load_targets[  # type: ignore[attr-defined]
        next(name for name, one in load_targets(book).items() if one is target)
    ] = LoadTarget(target.sheet, filled, target.table_part, target.table_name, list(table.columns))
    sheet.shape_changed()


def _put(sheet: Worksheet, row: int, column: int, value: object) -> None:
    from pyopenvba.apps.excel._model import as_cell_value

    cell = sheet.cell(row, column, create=True)
    assert cell is not None
    cell.formula = ""
    cell.stale = False
    if value is None:
        cell.value = EMPTY
    elif isinstance(value, _dt.datetime):
        cell.value = VBADate.from_datetime(value)
        if cell.number_format in ("General", ""):
            sheet.set_number_format(row, column, "m/d/yyyy h:mm")
    elif isinstance(value, _dt.date):
        cell.value = VBADate.from_datetime(value)
        if cell.number_format in ("General", ""):
            sheet.set_number_format(row, column, "m/d/yyyy")
    elif isinstance(value, Duration):
        cell.value = value.total_seconds / 86400.0
    elif isinstance(value, (Table, Record, list)):
        raise VBAUnsupportedError(
            "a query whose cells hold tables or records has to be expanded before it is loaded"
        )
    else:
        cell.value = as_cell_value(value)
