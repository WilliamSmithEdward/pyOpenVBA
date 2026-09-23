"""Whole-row/column movement and structural formula reference updates."""
from __future__ import annotations

import re
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, column_letter, column_number
from pyopenvba._xml import attributes
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Cell, Worksheet, NameEntry, Workbook

_CORNER = re.compile(r"^(\$?)([A-Za-z]+)?(\$?)([0-9]+)?$")


def interval(low: int, high: int, start: int, count: int, delete: bool, limit: int) -> tuple[int, int] | None:
    """Where the rows or columns ``low`` to ``high`` end up when ``count`` go in or out at ``start``; None when
    every one of them is deleted."""
    if not delete:
        low += count if low >= start else 0
        high += count if high >= start else 0
        return (low, min(high, limit)) if low <= limit else None
    end = start + count - 1
    if start <= low <= high <= end:
        return None
    low = low - count if low > end else start if low >= start else low
    high = high - count if high > end else start - 1 if high >= start else high
    return low, high


def rewrite(formula: str, owner: str, edited: str, *, rows: bool, start: int, count: int, delete: bool) -> str:
    """Update references to moved cells regardless of absolute-dollar markers."""
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(formula):
        if token.kind != "ref":
            continue
        sheet, reference = split_sheet(token.text)
        if (sheet or owner).casefold() != edited.casefold():
            continue
        corners = [_CORNER.fullmatch(part) for part in reference.split(":")]
        assert all(corner is not None for corner in corners)
        groups = [corner.groups() for corner in corners if corner is not None]
        values = [int(g[3]) if rows and g[3] else column_number(g[1]) if not rows and g[1] else None for g in groups]
        if any(value is None for value in values):
            continue  # Whole columns survive row edits, and vice versa.
        positions = [value for value in values if value is not None]
        result = interval(min(positions), max(positions), start, count, delete, MAX_ROWS if rows else MAX_COLUMNS)
        prefix = token.text[:-len(reference)]
        if result is None:
            replacement = "#REF!"
        else:
            low, high = result
            mapped = [low] if len(groups) == 1 else [low, high] if positions[0] <= positions[-1] else [high, low]
            output: list[str] = []
            for (col_mark, letters, row_mark, digits), value in zip(groups, mapped):
                if rows:
                    digits = str(value)
                else:
                    letters = column_letter(value)
                output.append(col_mark + (letters or "") + row_mark + (digits or ""))
            replacement = ":".join(output)
        pieces.append((token.at, token.at + len(token.text), prefix + replacement))
    for first, last, replacement in reversed(pieces):
        formula = formula[:first] + replacement + formula[last:]
    return formula


def name_scope(book: Workbook, entry: NameEntry) -> str:
    """The sheet a defined name belongs to, by its qualified name or its localSheetId; "" for the workbook's own."""
    scope, _ = split_sheet(entry.name)
    local_id = attributes(f"<definedName {entry.attributes}>").get("localSheetId", "")
    if not scope and local_id.isdigit() and int(local_id) < len(book.sheets_):
        scope = book.sheets_[int(local_id)].name
    return scope


def edit(target: Range, *, delete: bool) -> None:
    sheet, area = target.sheet, target.first
    if len(target.areas) != 1:
        raise VBAUnsupportedError("Structural edits of multiple areas are not implemented")
    if sheet.merged_areas or sheet.shapes_:
        raise VBAUnsupportedError("Whole-row/column edits on sheets with merges or shapes are not implemented")
    rows = area.whole_rows
    start, count, limit = (area.top, area.rows, MAX_ROWS) if rows else (area.left, area.columns, MAX_COLUMNS)
    moved: dict[tuple[int, int], Cell] = {}
    for (row, column), cell in sheet.cells_.items():
        position = row if rows else column
        result = interval(position, position, start, count, delete, limit)
        if result is None:
            if not delete and not cell.is_blank():
                raise error(1004, "Insertion would move nonempty cells beyond the worksheet")
            continue
        moved[(result[0], column) if rows else (row, result[0])] = cell
    formulas: list[tuple[Worksheet, Cell, str]] = []
    for owner in sheet.book.sheets_:
        for cell in (moved if owner is sheet else owner.cells_).values():
            if cell.formula:
                text = rewrite(cell.formula, owner.name, sheet.name, rows=rows, start=start, count=count, delete=delete)
                if text != cell.formula:
                    formulas.append((owner, cell, text))
    names: list[tuple[NameEntry, str]] = []
    for entry in sheet.book.names_.entries:
        text = rewrite(entry.refers_to, name_scope(sheet.book, entry) or sheet.name, sheet.name,
                       rows=rows, start=start, count=count, delete=delete)
        names.append((entry, text))
    # All validation/conversion precedes mutation.
    sheet.cells_ = moved
    for owner, cell, text in formulas:
        cell.formula = text
        owner.touched()
    for entry, text in names:
        if text != entry.refers_to:
            entry.refers_to = text
            sheet.book.names_.changed = True
    if rows:
        sheet.dims.shift_rows(start, count, delete)
    else:
        sheet.dims.shift_columns(start, count, delete)
    if not delete and start > 1:
        _inherit_formats(sheet, rows=rows, start=start, count=count)
    from pyopenvba.apps.excel._autofilter import filter_edited

    filter_edited(sheet, rows=rows, start=start, count=count, delete=delete)
    sheet.shape_changed()


def _inherit_formats(sheet: Worksheet, *, rows: bool, start: int, count: int) -> None:
    """New rows take the formats of the cells in the row above; new columns those of the column to the left.

    The row's or column's own format came along with its height or width;
    a cell here is made only where it would show something else.
    """
    from pyopenvba.apps.excel._model import Cell

    limit = MAX_ROWS if rows else MAX_COLUMNS
    for (row, column), cell in list(sheet.cells_.items()):
        if (row if rows else column) != start - 1:
            continue
        for index in range(start, min(start + count, limit + 1)):
            position = (index, column) if rows else (row, index)
            sheet.cells_[position] = Cell(style=cell.style)
            sheet.settle(*position)
