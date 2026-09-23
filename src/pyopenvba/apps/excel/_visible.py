"""What an edit reaches on a filtered sheet.

Measured in live Excel (scripts/measure_autofilter_edits.py): while a sheet
is in filter mode -- some column of its AutoFilter has criteria -- an edit
of a range that has any visible cell reaches only its visible cells, the
ones in no hidden row and no hidden column, whatever hid them; a range
with no visible cell is edited whole, so a single hidden cell still takes
a value. Values, formulas, formats and borders, clearing, filling,
Replace, Sort, row heights and hiding, copying and deleting all hold back
so; AutoFill, Merge, PasteSpecial, SpecialCells and every read do not.
Each visible area is edited as a range of its own, with two exceptions
the callers keep: a formula written to the range moves from the range's
own top-left cell, and a fill down or up takes the first or last visible
row for every other one. Outside filter mode nothing is held back, and a
row hidden by hand takes writes like any other.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBAUnsupportedError

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def filtering(sheet: Worksheet) -> bool:
    """Whether the sheet is in filter mode: a column of its AutoFilter has criteria."""
    found = sheet.auto_filter
    return found is not None and bool(found.fields)


def bands(first: int, last: int, hidden: set[int]) -> list[tuple[int, int]]:
    """The runs of positions from ``first`` to ``last`` that are not hidden."""
    found: list[tuple[int, int]] = []
    start = first
    for position in sorted(one for one in hidden if first <= one <= last):
        if position > start:
            found.append((start, position - 1))
        start = position + 1
    if start <= last:
        found.append((start, last))
    return found


def hidden_lines(sheet: Worksheet) -> tuple[set[int], set[int]]:
    """The sheet's hidden rows and hidden columns."""
    dims = sheet.dims
    if dims.zero_height:
        raise VBAUnsupportedError("editing a filtered sheet whose rows are hidden by default is not implemented")
    return ({row for row, record in dims.rows.items() if record.hidden},
            {column for column, record in dims.columns.items() if record.hidden})


def visible_areas(target: Range) -> list[Area] | None:
    """The areas an edit of ``target`` reaches, when that is not the whole range.

    None stands for the whole range: outside filter mode, when nothing in
    it is hidden, and when nothing in it shows.
    """
    sheet = target.sheet
    if not filtering(sheet):
        return None
    hidden_rows, hidden_columns = hidden_lines(sheet)
    found: list[Area] = []
    cut = False
    for area in target.areas:
        rows = bands(area.top, area.bottom, hidden_rows)
        columns = bands(area.left, area.right, hidden_columns)
        cut = cut or rows != [(area.top, area.bottom)] or columns != [(area.left, area.right)]
        found += [Area(top, left, bottom, right, area.sheet) for top, bottom in rows for left, right in columns]
    return found if cut and found else None


def visible(target: Range) -> Range:
    """The range an edit of ``target`` reaches."""
    from pyopenvba.apps.excel._model import Range

    found = visible_areas(target)
    if found is None:
        return target
    # Hidden columns cut whole rows into pieces, and hidden rows whole columns.
    whole = target.whole if all(area.whole_rows if target.whole == "rows" else area.whole_columns
                                for area in found) else ""
    return Range(target.sheet, found, whole=whole)


def visible_rows(target: Range) -> list[tuple[int, int]] | None:
    """The runs of rows an edit of the range's rows reaches, when that is not all of them: the shown ones."""
    sheet = target.sheet
    if not filtering(sheet):
        return None
    hidden_rows, _ = hidden_lines(sheet)
    found: list[tuple[int, int]] = []
    cut = False
    for area in target.areas:
        runs = bands(area.top, area.bottom, hidden_rows)
        cut = cut or runs != [(area.top, area.bottom)]
        found += runs
    return found if cut and found else None


def unmeasured(target: Range, what: str, *, lines: str = "cells") -> None:
    """Refuse an edit whose filter-mode behaviour was not measured, where the range takes in hidden cells.

    ``lines`` is "rows" or "columns" for an edit of the range's rows or columns, which only those can hide.
    """
    sheet = target.sheet
    if not filtering(sheet):
        return
    hidden_rows, hidden_columns = hidden_lines(sheet)
    rows = lines != "columns" and any(area.top <= row <= area.bottom for area in target.areas for row in hidden_rows)
    columns = lines != "rows" and any(area.left <= column <= area.right
                                      for area in target.areas for column in hidden_columns)
    if rows or columns:
        raise VBAUnsupportedError(f"{what} on a filtered sheet, over hidden cells, is not implemented")
