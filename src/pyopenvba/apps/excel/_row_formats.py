"""Formats a whole row, a whole column or the whole sheet carries.

Besides each cell's own format, Excel keeps one for a row (the row
element's ``s`` with ``customFormat``) and one for a column (the col
element's ``style``). A position with no cell shows its row's format if
the row has one, else its column's, else the default. Every rule here
was measured in live Excel (scripts/measure_row_formats.py):

- Formatting a range that spans every column formats its rows, one that
  spans every row its columns, and one that spans both every column.
  Each cell in them changes from its own format, and the row or column
  format from its own, or from the sheet's where it had none: the sheet's
  format is whatever column XFD has.
- A position whose row or column format would no longer show what the
  change made of it gets a cell of its own: a formatted row across a
  formatted column, in either order. A cell with nothing in it goes the
  moment its format is the one its row or column would give it.
- ClearFormats takes the row's or column's format away and puts every
  cell in it back to the default, giving one to each position that would
  otherwise show a column's or row's format.
- Borders on whole rows go on the row format, except that the left edge
  goes on the cells of column A and the right edge only marks the format
  as setting borders; on whole columns the top and bottom edges only mark
  it, bar the cells already in row 1, and on the whole sheet every edge
  only marks it.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel import _styles as S
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Worksheet

Change = Callable[[S.Style], S.Style]


def whole(area: Area) -> str:
    """"sheet", "rows" or "columns" for a range spanning every column, every row or both; else ""."""
    rows = area.left == 1 and area.right == MAX_COLUMNS
    columns = area.top == 1 and area.bottom == MAX_ROWS
    return "sheet" if rows and columns else "rows" if rows else "columns" if columns else ""


def sheet_format(sheet: Worksheet) -> S.Style:
    """The format a row takes a change from when it has none of its own: column XFD's."""
    return sheet.dims.column_style(MAX_COLUMNS) or sheet.book.stylesheet.default


def _cells_by_row(sheet: Worksheet, top: int, bottom: int) -> dict[int, dict[int, Cell]]:
    found: dict[int, dict[int, Cell]] = {}
    for (row, column), cell in sheet.cells_.items():
        if top <= row <= bottom:
            found.setdefault(row, {})[column] = cell
    return found


def _cells_by_column(sheet: Worksheet, left: int, right: int) -> dict[int, dict[int, Cell]]:
    found: dict[int, dict[int, Cell]] = {}
    for (row, column), cell in sheet.cells_.items():
        if left <= column <= right:
            found.setdefault(column, {})[row] = cell
    return found


def shared(change: Change) -> Change:
    """The change, made once for each format object it meets.

    The columns a macro formats together, or the cells one row format
    gave, go on sharing one format object, which keeps comparing them
    cheap.
    """
    made: dict[int, tuple[S.Style, S.Style]] = {}

    def changed(style: S.Style) -> S.Style:
        found = made.get(id(style))
        if found is None or found[0] is not style:
            found = made[id(style)] = (style, change(style))
        return found[1]

    return changed


def _put(sheet: Worksheet, row: int, column: int, style: S.Style) -> None:
    """A cell of its own for a position, in a format."""
    from pyopenvba.apps.excel._model import Cell

    default = sheet.book.stylesheet.default
    sheet.cells_[(row, column)] = Cell(style=None if style == default else style)
    sheet.book.stylesheet.meet(style)


def _restyle_cell(sheet: Worksheet, row: int, column: int, cell: Cell, style: S.Style) -> None:
    cell.style, cell.xf = (None if style == sheet.book.stylesheet.default else style), -1
    sheet.book.stylesheet.meet(cell.style)


def _set_row(sheet: Worksheet, row: int, style: S.Style | None) -> None:
    dims = sheet.dims
    record = dims.rows.get(row)
    if record is None and style is None:
        return
    record = dims.touch_row(row)
    record.style, record.xf = style, -1
    sheet.book.stylesheet.meet(style)
    dims.fonts_changed(row)
    dims.settle_row(row)


def _set_column(sheet: Worksheet, column: int, style: S.Style | None) -> None:
    dims = sheet.dims
    record = dims.columns.get(column)
    if record is None and style is None:
        return
    record = dims.touch_column(column)
    record.style, record.xf = (None if style == sheet.book.stylesheet.default else style), -1
    sheet.book.stylesheet.meet(record.style)
    dims.settle_column(column)


# --- formatting ------------------------------------------------------------------------------


def format_rows(sheet: Worksheet, top: int, bottom: int, change: Change) -> None:
    """Whole rows formatted: each row's format and each of its cells change from what they show."""
    change = shared(change)
    default = sheet.book.stylesheet.default
    base = sheet_format(sheet)
    # Where a row with no format of its own shows something besides the sheet's format.
    apart = sheet.dims.formatted_columns()
    cells = _cells_by_row(sheet, top, bottom)
    for row in range(top, bottom + 1):
        mine = cells.get(row, {})
        record = sheet.dims.rows.get(row)
        old = record.style if record is not None else None
        new = change(old if old is not None else base)
        if old is None:
            shown_by_row = new if new != base else None
            for column in apart:
                if column in mine:
                    continue
                shown = sheet.dims.column_style(column) or default
                wanted = change(shown)
                if wanted != (shown_by_row if shown_by_row is not None else shown):
                    _put(sheet, row, column, wanted)
            if shown_by_row is not None:
                _set_row(sheet, row, new)
        else:
            _set_row(sheet, row, new)
        for column, cell in mine.items():
            _restyle_cell(sheet, row, column, cell, change(cell.style or default))
            sheet.settle(row, column)
        sheet.dims.fonts_changed(row)


def format_columns(sheet: Worksheet, left: int, right: int, change: Change) -> None:
    """Whole columns formatted: a formatted row across one keeps its format there only in a cell."""
    change = shared(change)
    default = sheet.book.stylesheet.default
    styled_rows = [(row, record.style) for row, record in sheet.dims.rows.items() if record.style is not None]
    cells = _cells_by_column(sheet, left, right)
    for column in range(left, right + 1):
        mine = cells.get(column, {})
        for row, style in styled_rows:
            if row not in mine and change(style) != style:
                _put(sheet, row, column, change(style))
        _set_column(sheet, column, change(sheet.dims.column_style(column) or default))
        for row, cell in mine.items():
            _restyle_cell(sheet, row, column, cell, change(cell.style or default))
            sheet.settle(row, column)
            sheet.dims.fonts_changed(row)
    sheet.dims.column_fonts_changed()


def format_sheet(sheet: Worksheet, change: Change) -> None:
    """The whole sheet formatted: every column, every formatted row and every cell."""
    change = shared(change)
    default = sheet.book.stylesheet.default
    for column in range(1, MAX_COLUMNS + 1):
        _set_column(sheet, column, change(sheet.dims.column_style(column) or default))
    for row, record in list(sheet.dims.rows.items()):
        if record.style is not None:
            _set_row(sheet, row, change(record.style))
    for (row, column), cell in list(sheet.cells_.items()):
        _restyle_cell(sheet, row, column, cell, change(cell.style or default))
        sheet.settle(row, column)
        sheet.dims.fonts_changed(row)
    sheet.dims.column_fonts_changed()


def format_area(sheet: Worksheet, area: Area, change: Change) -> bool:
    """Format an area that spans whole rows, whole columns or the sheet; False for any other."""
    kind = whole(area)
    if kind == "sheet":
        format_sheet(sheet, change)
    elif kind == "rows":
        format_rows(sheet, area.top, area.bottom, change)
    elif kind == "columns":
        format_columns(sheet, area.left, area.right, change)
    else:
        return False
    sheet.touched()
    return True


# --- clearing ----------------------------------------------------------------------------------


def clear_area(sheet: Worksheet, area: Area) -> bool:
    """ClearFormats on whole rows, whole columns or the sheet; False for any other area."""
    kind = whole(area)
    if not kind:
        return False
    default = sheet.book.stylesheet.default
    if kind == "sheet":
        for column in list(sheet.dims.columns):
            _set_column(sheet, column, None)
        for row in list(sheet.dims.rows):
            _set_row(sheet, row, None)
        sheet.dims.column_fonts_changed()
    elif kind == "rows":
        styled = [column for column, record in sheet.dims.columns.items() if record.style is not None]
        cells = _cells_by_row(sheet, area.top, area.bottom)
        for row in range(area.top, area.bottom + 1):
            _set_row(sheet, row, None)
            for column in styled:
                if column not in cells.get(row, {}):
                    _put(sheet, row, column, default)
    else:
        styled_rows = [row for row, record in sheet.dims.rows.items()
                       if record.style is not None and record.style != default]
        cells = _cells_by_column(sheet, area.left, area.right)
        for column in range(area.left, area.right + 1):
            _set_column(sheet, column, None)
            for row in styled_rows:
                if row not in cells.get(column, {}):
                    _put(sheet, row, column, default)
        sheet.dims.column_fonts_changed()
    for (row, column), cell in list(sheet.cells_.items()):
        if area.contains(row, column):
            _restyle_cell(sheet, row, column, cell, default)
            sheet.settle(row, column)
            sheet.dims.fonts_changed(row)
    sheet.touched()
    return True


# --- copying ----------------------------------------------------------------------------------------


def shown_formats(source: Worksheet, area: Area, destination: Worksheet, top: int, left: int,
                  cells: dict[tuple[int, int], Cell], *, rows_follow: bool, columns_follow: bool,
                  ) -> list[tuple[tuple[int, int], S.Style]]:
    """What the positions of a copied area without a cell showed, where the destination may show another.

    A position shows its row's format, else its column's; where the rows
    or columns go along with the copy, their formats do too. Taken before
    anything is written, since the source and the destination can overlap.
    """
    down, across = top - area.top, left - area.left
    rows: set[int] = set()
    if not rows_follow:
        rows = {row for row in range(area.top, area.bottom + 1)
                if _row_style(source, row) != _row_style(destination, row + down)} \
            if area.rows <= 4096 else _differing_rows(source, destination, area, down)
    columns: set[int] = set()
    if not columns_follow:
        columns = {column for column in range(area.left, area.right + 1)
                   if source.dims.column_style(column) != destination.dims.column_style(column + across)}
    out: list[tuple[tuple[int, int], S.Style]] = []
    for row in sorted(rows):
        out.extend(((row + down, column + across), source.inherited_style(row, column))
                   for column in range(area.left, area.right + 1)
                   if (row - area.top, column - area.left) not in cells)
    for column in sorted(columns):
        out.extend(((row + down, column + across), source.inherited_style(row, column))
                   for row in range(area.top, area.bottom + 1)
                   if row not in rows and (row - area.top, column - area.left) not in cells)
    return out


def _row_style(sheet: Worksheet, row: int) -> S.Style | None:
    record = sheet.dims.rows.get(row)
    return record.style if record is not None else None


def _differing_rows(source: Worksheet, destination: Worksheet, area: Area, down: int) -> set[int]:
    """The rows of a tall area whose format differs between the source and where it lands."""
    candidates = {row for row, record in source.dims.rows.items() if record.style is not None}
    candidates.update(row - down for row, record in destination.dims.rows.items() if record.style is not None)
    return {row for row in candidates if area.top <= row <= area.bottom
            and _row_style(source, row) != _row_style(destination, row + down)}


# --- reading --------------------------------------------------------------------------------------


def distinct_styles(sheet: Worksheet, area: Area) -> set[S.Style]:
    """Every format some position of a large area shows, without visiting each position.

    A row with a format shows it wherever it has no cell; a row without
    one shows each column's format wherever one of those columns has no
    cell in it; and a row with neither cells nor a format shows every
    column format in the area.
    """
    default = sheet.book.stylesheet.default
    top, left = area.top, area.left
    bottom, right = min(area.bottom, MAX_ROWS), min(area.right, MAX_COLUMNS)
    width = right - left + 1
    groups = sheet.dims.column_groups(left, right)
    unstyled = width - sum(groups.values())
    if unstyled:
        groups[default] = groups.get(default, 0) + unstyled
    by_row: dict[int, list[tuple[int, Cell]]] = {}
    for (row, column), cell in sheet.cells_.items():
        if top <= row <= bottom and left <= column <= right:
            by_row.setdefault(row, []).append((column, cell))
    styled_rows = {row: record.style for row, record in sheet.dims.rows.items()
                   if top <= row <= bottom and record.style is not None}
    found: set[S.Style] = set()
    for row in set(by_row) | set(styled_rows):
        cells = by_row.get(row, [])
        found.update(cell.style or default for _, cell in cells)
        if len(cells) == width:
            continue
        own = styled_rows.get(row)
        if own is not None:
            found.add(own)
            continue
        covered: dict[S.Style, int] = {}
        for column, _ in cells:
            shown = sheet.dims.column_style(column) or default
            covered[shown] = covered.get(shown, 0) + 1
        found.update(style for style, count in groups.items() if count > covered.get(style, 0))
    if bottom - top + 1 > len(set(by_row) | set(styled_rows)):
        found.update(groups)
    return found


def sample_rows(sheet: Worksheet, top: int, bottom: int, left: int, right: int) -> list[int]:
    """Rows of a tall area that stand for all of them along a border.

    Those with cells or formats, their neighbours, and one row with neither.
    """
    special: set[int] = {row for (row, column) in sheet.cells_ if left - 1 <= column <= right + 1}
    special.update(row for row, record in sheet.dims.rows.items() if record.style is not None)
    special = {one for row in special for one in (row - 1, row, row + 1) if top <= one <= bottom}
    plain = next((row for row in range(top, bottom + 1) if row not in special), None)
    return sorted(special) + ([plain] if plain is not None else [])


def sample_columns(sheet: Worksheet, left: int, right: int, top: int, bottom: int) -> list[int]:
    """Columns of a wide area that stand for all of them along a border."""
    special: set[int] = {column for (row, column) in sheet.cells_ if top - 1 <= row <= bottom + 1}
    special.update(sheet.dims.formatted_columns())
    special = {one for column in special for one in (column - 1, column, column + 1) if left <= one <= right}
    plain = next((column for column in range(left, right + 1) if column not in special), None)
    return sorted(special) + ([plain] if plain is not None else [])


# --- borders ------------------------------------------------------------------------------------


def _sides(style: S.Style, **sides: S.Side) -> S.Style:
    return S.applying(style, "border", border=replace(style.border, **sides))


def _touch(style: S.Style) -> S.Style:
    return S.applying(style, "border")


def _current(found: Iterable[S.Side]) -> S.Side:
    """What a border along many cells is before a change.

    The line they share, or, where only some show one, that line in the
    automatic colour.
    """
    distinct = set(found)
    if len(distinct) == 1:
        return distinct.pop()
    lines = [side for side in distinct if side.style]
    if len(distinct) == 2 and len(lines) == 1:
        return S.Side(lines[0].style, S.Color("auto"))
    raise VBAUnsupportedError("a border along whole rows or columns whose cells show different lines is not "
                              "implemented")


def border_area(sheet: Worksheet, area: Area, index: int, change: Callable[[S.Side], S.Side]) -> bool:
    """A border set on whole rows, whole columns or the sheet; False for any other area."""
    from pyopenvba.apps.excel._formats import change_side, effective_side, with_diagonal

    kind = whole(area)
    if not kind:
        return False
    if index in (5, 6):
        format_area(sheet, area, lambda style: S.applying(style, "border",
                                                          border=with_diagonal(style.border, index, change)))
        return True
    if kind == "rows":
        columns = sample_columns(sheet, 1, MAX_COLUMNS, area.top - 1, area.bottom + 1)
        if index in (8, 9):
            row = area.top if index == 8 else area.bottom
            name, opposite, beyond = ("top", "bottom", row - 1) if index == 8 else ("bottom", "top", row + 1)
            new = change(_current(effective_side(sheet, row, column, name) for column in columns))
            format_rows(sheet, row, row, lambda style: _sides(style, **{name: new}))
            if 1 <= beyond <= MAX_ROWS:
                format_rows(sheet, beyond, beyond, lambda style: _sides(style, **{opposite: S.NO_SIDE})
                            if getattr(style.border, opposite).style else style)
        elif index == 12:
            for row in range(area.top, area.bottom):
                new = change(_current_pair(sheet, [(row, column, row + 1, column) for column in columns],
                                           "bottom", "top"))
                format_rows(sheet, row, row, lambda style: _sides(style, bottom=new))  # noqa: B023
                format_rows(sheet, row + 1, row + 1, lambda style: _sides(style, top=new))  # noqa: B023
        elif index == 11:
            pairs = [(row, column, row, column + 1) for row in range(area.top, area.bottom + 1)
                     for column in columns if column < MAX_COLUMNS]
            new = change(_current_pair(sheet, pairs, "right", "left"))
            format_rows(sheet, area.top, area.bottom, lambda style: _sides(style, left=new, right=new))
        else:
            format_rows(sheet, area.top, area.bottom, _touch)
            if index == 7:
                for row in range(area.top, area.bottom + 1):
                    change_side(sheet, row, 1, 7, change)
            elif any(area.top <= row <= area.bottom for row, column in sheet.cells_ if column == MAX_COLUMNS):
                raise VBAUnsupportedError("the right edge of whole rows with cells in column XFD is not implemented")
    elif kind == "columns":
        rows = sample_rows(sheet, 1, MAX_ROWS, area.left - 1, area.right + 1)
        if index in (7, 10):
            column = area.left if index == 7 else area.right
            name, opposite, beyond = ("left", "right", column - 1) if index == 7 else ("right", "left", column + 1)
            new = change(_current(effective_side(sheet, row, column, name) for row in rows))
            format_columns(sheet, column, column, lambda style: _sides(style, **{name: new}))
            if 1 <= beyond <= MAX_COLUMNS:
                format_columns(sheet, beyond, beyond, lambda style: _sides(style, **{opposite: S.NO_SIDE})
                               if getattr(style.border, opposite).style else style)
        elif index == 11:
            for column in range(area.left, area.right):
                new = change(_current_pair(sheet, [(row, column, row, column + 1) for row in rows], "right", "left"))
                format_columns(sheet, column, column, lambda style: _sides(style, right=new))  # noqa: B023
                format_columns(sheet, column + 1, column + 1, lambda style: _sides(style, left=new))  # noqa: B023
        elif index == 12:
            pairs = [(row, column, row + 1, column) for column in range(area.left, area.right + 1)
                     for row in rows if row < MAX_ROWS]
            new = change(_current_pair(sheet, pairs, "bottom", "top"))
            format_columns(sheet, area.left, area.right, lambda style: _sides(style, top=new, bottom=new))
        else:
            edge_row = 1 if index == 8 else MAX_ROWS
            record = sheet.dims.rows.get(edge_row)
            edge_cells = [column for row, column in sheet.cells_
                          if row == edge_row and area.left <= column <= area.right]
            if (record is not None and record.style is not None) or (edge_cells and index == 9):
                raise VBAUnsupportedError("the top or bottom edge of whole columns across a formatted row or cells in "
                                          "the last row is not implemented")
            format_columns(sheet, area.left, area.right, _touch)
            for column in edge_cells:
                change_side(sheet, edge_row, column, index, change)
    else:
        if index in (7, 8, 9, 10):
            if any(_on_edge(index, row, column) for row, column in sheet.cells_):
                raise VBAUnsupportedError("an edge of the whole sheet with cells along it is not implemented")
            format_sheet(sheet, _touch)
        else:
            styles = _shown_styles(sheet)
            first, second = ("bottom", "top") if index == 12 else ("right", "left")
            new = change(_current(side for style in styles
                                  for side in (getattr(style.border, first), getattr(style.border, second))))
            format_sheet(sheet, lambda style: _sides(style, **{first: new, second: new}))
    sheet.touched()
    return True


def _on_edge(index: int, row: int, column: int) -> bool:
    """Whether a position lies along the sheet's left (7), top (8), bottom (9) or right (10) edge."""
    return {7: column == 1, 8: row == 1, 9: row == MAX_ROWS, 10: column == MAX_COLUMNS}[index]


def _current_pair(sheet: Worksheet, pairs: Iterable[tuple[int, int, int, int]], first: str, second: str) -> S.Side:
    """What an inside border is along many pairs of cells: each pair's line, which must agree."""
    from pyopenvba.apps.excel._formats import inside_current, own_side

    return _current(inside_current(own_side(sheet, row, column, first),
                                   own_side(sheet, other_row, other_column, second))
                    for row, column, other_row, other_column in pairs)


def _shown_styles(sheet: Worksheet) -> set[S.Style]:
    """Every format anywhere on the sheet."""
    default = sheet.book.stylesheet.default
    groups = sheet.dims.column_groups()
    found = set(groups)
    if sum(groups.values()) < MAX_COLUMNS:
        found.add(default)
    found.update(record.style for record in sheet.dims.rows.values() if record.style is not None)
    found.update(cell.style or default for cell in sheet.cells_.values())
    return found
