"""Whole-row/column movement, cell shifts, and the formula references both update.

A cell shift -- Range.Delete or Range.Insert with a Shift, or none, which
shifts a range taller than it is wide across and any other range up or
down -- moves the cells in the band of columns (or rows) the range spans.
Measured in live Excel (scripts/measure_cell_shifts.py, 43 layouts): a
reference whose columns lie in the band moves and stretches as it would
for whole rows, becoming #REF! when every cell it reads goes; one reaching
outside the band stays as it was, unless a delete takes the whole band's
part of it, rows and all, at one edge -- then it keeps the rest, so
SUM(B2:C6) reads SUM(C2:C6) once B2:B6 is deleted up. A whole-column
reference stays through shifts up and down, a whole-row one through
shifts across. Names and the sheet's AutoFilter range follow the same
rule, and a delete that takes the filter's header cells removes it.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter, column_number
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


def _inherit_formats(sheet: Worksheet, *, rows: bool, start: int, count: int,
                     band: tuple[int, int] | None = None) -> None:
    """New rows take the formats of the cells in the row above; new columns those of the column to the left.

    The row's or column's own format came along with its height or width;
    a cell here is made only where it would show something else. Cells
    inserted into a ``band`` of columns (or rows) take theirs the same way.
    """
    from pyopenvba.apps.excel._model import Cell

    limit = MAX_ROWS if rows else MAX_COLUMNS
    for (row, column), cell in list(sheet.cells_.items()):
        if (row if rows else column) != start - 1:
            continue
        if band is not None and not band[0] <= (column if rows else row) <= band[1]:
            continue
        for index in range(start, min(start + count, limit + 1)):
            position = (index, column) if rows else (row, index)
            sheet.cells_[position] = Cell(style=cell.style)
            sheet.settle(*position)


# --- cell shifts ---------------------------------------------------------------------------------------


def shift_cells(target: Range, area: Area, *, delete: bool, vertical: bool) -> None:
    """``area`` deleted or inserted, the cells in its band of columns moving up or down -- ``vertical`` -- or in
    its band of rows across, and every reference to them following as the module docstring has it."""
    sheet = target.sheet
    if sheet.merged_areas or sheet.shapes_:
        raise VBAUnsupportedError("shifting cells on a sheet with merges or shapes is not implemented")
    shift = CellShift.of(area, delete=delete, vertical=vertical)
    band, start, count = shift.band, shift.start, shift.count
    limit = MAX_ROWS if vertical else MAX_COLUMNS
    moved: dict[tuple[int, int], Cell] = {}
    for (row, column), cell in sheet.cells_.items():
        along, across = (row, column) if vertical else (column, row)
        if band[0] <= across <= band[1] and along >= start:
            if delete and along < start + count:
                continue
            along = along - count if delete else along + count
            if along > limit:
                if not cell.is_blank():
                    raise error(1004, "Insertion would move nonempty cells beyond the worksheet")
                continue
        moved[(along, across) if vertical else (across, along)] = cell
    formulas: list[tuple[Worksheet, Cell, str]] = []
    for owner in sheet.book.sheets_:
        for cell in (moved if owner is sheet else owner.cells_).values():
            if cell.formula:
                text = rewrite_shift(cell.formula, owner.name, sheet.name, shift)
                if text != cell.formula:
                    formulas.append((owner, cell, text))
    names: list[tuple[NameEntry, str]] = []
    for entry in sheet.book.names_.entries:
        names.append((entry, rewrite_shift(entry.refers_to, name_scope(sheet.book, entry) or sheet.name, sheet.name,
                                           shift)))
    from pyopenvba.apps.excel._autofilter import filter_shifted

    follow = filter_shifted(sheet, shift)
    # All validation precedes mutation.
    sheet.cells_ = moved
    for owner, cell, text in formulas:
        cell.formula = text
        owner.touched()
    for entry, text in names:
        if text != entry.refers_to:
            entry.refers_to = text
            sheet.book.names_.changed = True
    if not delete and start > 1:
        # Measured: inserted cells take the formats of the cells above them, or to their left.
        _inherit_formats(sheet, rows=vertical, start=start, count=count, band=band)
    follow()
    sheet.shape_changed()


@dataclass(frozen=True)
class CellShift:
    """A partial Delete or Insert: ``count`` rows (``vertical``) or columns from ``start`` going out or in,
    across the ``band`` of columns (or rows) the range spans."""

    vertical: bool
    band: tuple[int, int]
    start: int
    count: int
    delete: bool

    @staticmethod
    def of(area: Area, *, delete: bool, vertical: bool) -> CellShift:
        if vertical:
            return CellShift(True, (area.left, area.right), area.top, area.rows, delete)
        return CellShift(False, (area.top, area.bottom), area.left, area.columns, delete)


def shifted_box(box: tuple[int, int, int, int], shift: CellShift, *,
                whole: bool = False) -> tuple[int, int, int, int] | None:
    """Where a reference over ``box`` -- top, left, bottom, right -- goes when cells shift; None for #REF!.

    ``whole`` says the box spans every column (or row) because the
    reference names none, which no shift trims: SUM(3:3) stays through
    A3:C4 deleted up.
    """
    vertical, band, start, count, delete = shift.vertical, shift.band, shift.start, shift.count, shift.delete
    top, left, bottom, right = box
    along, across = ((top, bottom), (left, right)) if vertical else ((left, right), (top, bottom))
    if band[0] <= across[0] and across[1] <= band[1]:
        moved = interval(along[0], along[1], start, count, delete, MAX_ROWS if vertical else MAX_COLUMNS)
        if moved is None:
            return None
        along = moved
    elif delete and across[0] <= band[1] and band[0] <= across[1] and start <= along[0] and \
            along[1] <= start + count - 1:
        # Every cell of it in the band goes: what is left stays, if it is still a block.
        if whole:
            return box
        if band[0] <= across[0]:
            across = (band[1] + 1, across[1])
        elif across[1] <= band[1]:
            across = (across[0], band[0] - 1)
        else:
            return box
    else:
        return box
    return (along[0], across[0], along[1], across[1]) if vertical else (across[0], along[0], across[1], along[1])


def rewrite_shift(formula: str, owner: str, edited: str, shift: CellShift) -> str:
    """References to cells a partial Delete or Insert moves, updated as shifted_box has them."""
    vertical = shift.vertical
    pieces: list[tuple[int, int, str]] = []
    for token in tokenize(formula):
        if token.kind != "ref":
            continue
        sheet, reference = split_sheet(token.text)
        if (sheet or owner).casefold() != edited.casefold():
            continue
        corners = [_CORNER.fullmatch(part) for part in reference.split(":")]
        groups = [corner.groups() for corner in corners if corner is not None]
        if len(groups) != len(corners):
            continue
        rows = [int(g[3]) if g[3] else None for g in groups]
        columns = [column_number(g[1]) if g[1] else None for g in groups]
        if any(value is None for value in (rows if vertical else columns)):
            continue  # A whole column stays through shifts up and down, a whole row through shifts across.
        whole = any(value is None for value in (columns if vertical else rows))
        known_rows = [value for value in rows if value is not None] or [1, MAX_ROWS]
        known_columns = [value for value in columns if value is not None] or [1, MAX_COLUMNS]
        box = (min(known_rows), min(known_columns), max(known_rows), max(known_columns))
        moved = shifted_box(box, shift, whole=whole)
        if moved == box:
            continue
        prefix = token.text[:-len(reference)]
        if moved is None:
            pieces.append((token.at, token.at + len(token.text), prefix + "#REF!"))
            continue
        top, left, bottom, right = moved
        new_rows = _laid(rows, top, bottom)
        new_columns = _laid(columns, left, right)
        output = [col_mark + (column_letter(column) if column is not None else "") + row_mark +
                  (str(row) if row is not None else "")
                  for (col_mark, _, row_mark, _), row, column in zip(groups, new_rows, new_columns)]
        pieces.append((token.at, token.at + len(token.text), prefix + ":".join(output)))
    for first, last, replacement in reversed(pieces):
        formula = formula[:first] + replacement + formula[last:]
    return formula


def _laid(values: list[int | None], low: int, high: int) -> list[int | None]:
    """A reference's corners along one axis moved to ``low`` and ``high``, each keeping its end of the range."""
    if any(value is None for value in values):
        return values
    if len(values) == 1:
        return [low]
    first, second = values[0], values[1]
    assert first is not None and second is not None
    return [low, high] if first <= second else [high, low]
