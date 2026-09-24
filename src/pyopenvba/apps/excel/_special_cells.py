"""Range.SpecialCells, as Excel answers it.

Measured in live Excel (scripts/measure_special_cells.py): 18 layouts,
eight of them seeded random grids, each asked for constants and formulas
of every kind, blanks, the last cell and visible cells.

- A range of one cell searches from A1 to the used range's last cell, even
  where the used range starts further in; a larger range, its own cells
  within that. A sheet with nothing on it finds nothing, not even
  blanks. Visible cells are the exception: a visible single cell answers
  itself, a hidden one the whole sheet's.
- Constants and formulas can be narrowed to numbers, text, logical
  values and errors, a formula by what it returns. Blanks are the cells
  in the used range with neither a value nor a formula, a format alone
  not counting. A merged area comes whole when any cell of it is found.
- The last cell is the used range's bottom-right cell, whatever range
  asked. Finding nothing is error 1004.

How Excel splits what it finds into areas is half the answer, and one
procedure accounts for every measured case: the cells are taken from
the last to the first, row by row; each joins the area below it when that
area is one column wide and in its column, else the one-row area to its
right, and a row that grows to span exactly the area below it merges
with it; a new area goes to the front of the list.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel._visible import bands
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBADate, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

CONSTANTS, FORMULAS, BLANKS, LAST_CELL, VISIBLE, COMMENTS = 2, -4123, 4, 11, 12, -4144
#: The kinds that ask about validation and conditional formats, which the model does not keep.
_UNMODELLED = {-4174: "data validation", -4175: "data validation",
               -4172: "conditional formats", -4173: "conditional formats"}
NUMBERS, TEXT, LOGICAL, ERRORS = 1, 2, 4, 16


def special_cells(target: Range, kind_value: object, which: object) -> Range:
    from pyopenvba.apps.excel._model import Range

    kind = int(to_integer(kind_value, "Long"))
    sheet = target.sheet
    if kind in _UNMODELLED:
        raise VBAUnsupportedError(f"SpecialCells for {_UNMODELLED[kind]} is not implemented")
    bounds = sheet.used_bounds()
    if kind == LAST_CELL:
        row, column = (bounds[2], bounds[3]) if bounds else (1, 1)
        return Range(sheet, [Area(row, column, row, column, sheet.name)])
    if kind == VISIBLE:
        return Range(sheet, _visible(target))
    if kind not in (CONSTANTS, FORMULAS, BLANKS, COMMENTS):
        raise error(1004, "SpecialCells has no such type")
    if bounds is None:
        # A sheet with nothing on it has not even blanks.
        raise error(1004, "No cells were found.")
    flags = 23 if which is MISSING else int(to_integer(which, "Long"))
    # The search runs from A1 to the last cell, even where the used range starts further in.
    searched = Area(1, 1, bounds[2], bounds[3], sheet.name)
    domains = [searched] if target.single else [_clipped(area, searched) for area in target.areas]
    found: list[tuple[int, int]] = []
    seen: set[tuple[int, int]] = set()
    for domain in domains:
        if domain is None:
            continue
        for position in _matching(sheet, domain, kind, flags):
            if position not in seen:
                seen.add(position)
                found.append(position)
    found = _with_merges(sheet, found, seen)
    if not found:
        raise error(1004, "No cells were found.")
    return Range(sheet, [Area(top, left, bottom, right, sheet.name) for top, left, bottom, right in areas_of(found)])


def _clipped(area: Area, used: Area) -> Area | None:
    top, left = max(area.top, used.top), max(area.left, used.left)
    bottom, right = min(area.bottom, used.bottom), min(area.right, used.right)
    return Area(top, left, bottom, right) if top <= bottom and left <= right else None


def _kind_of(value: object) -> int:
    if isinstance(value, bool):
        return LOGICAL
    if isinstance(value, (int, float, VBADate)):
        return NUMBERS
    if isinstance(value, ExcelError):
        return ERRORS
    return TEXT


def _matching(sheet: Worksheet, domain: Area, kind: int, flags: int) -> list[tuple[int, int]]:
    """The positions in ``domain`` of the kind asked for, row by row."""
    if kind == BLANKS:
        return [(row, column) for row in range(domain.top, domain.bottom + 1)
                for column in range(domain.left, domain.right + 1) if not _holds(sheet, row, column)]
    if kind == COMMENTS:
        # The cells with a note (tests/fixtures/excel_model/).
        return [position for position in sorted(sheet.notes) if domain.contains(*position)]
    out: list[tuple[int, int]] = []
    for (row, column), cell in sorted(sheet.cells_.items()):
        if not domain.contains(row, column):
            continue
        if kind == FORMULAS and cell.formula:
            value = sheet.book.calculator.value_of(sheet.name, row, column)
        elif kind == CONSTANTS and not cell.formula and cell.value is not EMPTY and cell.spilled_from is None:
            # A cell a formula spilled into is neither a formula nor a constant (tests/fixtures/dynamic_arrays.json).
            value = cell.value
        else:
            continue
        if _kind_of(value) & flags:
            out.append((row, column))
    return out


def _holds(sheet: Worksheet, row: int, column: int) -> bool:
    cell = sheet.cells_.get((row, column))
    return cell is not None and (cell.value is not EMPTY or bool(cell.formula))


def _with_merges(sheet: Worksheet, found: list[tuple[int, int]], seen: set[tuple[int, int]]) -> list[tuple[int, int]]:
    """What was found, with every merged area it touches taken whole."""
    extra: list[tuple[int, int]] = []
    for area in sheet.merged_areas:
        if any(area.contains(row, column) for row, column in found):
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    if (row, column) not in seen:
                        seen.add((row, column))
                        extra.append((row, column))
    return found + extra


def areas_of(cells: list[tuple[int, int]]) -> list[list[int]]:
    """Cells split into areas as SpecialCells splits them; each area is [top, left, bottom, right]."""
    owner: dict[tuple[int, int], list[int]] = {}
    ordered: list[list[int]] = []
    for row, column in sorted(set(cells), reverse=True):
        below = owner.get((row + 1, column))
        right = owner.get((row, column + 1))
        if below is not None and below[0] == row + 1 and below[1] == below[3] == column:
            below[0] = row
            owner[(row, column)] = below
            continue
        if right is not None and right[0] == right[2] == row and right[1] == column + 1:
            right[1] = column
            owner[(row, column)] = right
            under = owner.get((row + 1, column))
            if under is not None and under is not right and under[0] == row + 1 \
                    and under[1] == right[1] and under[3] == right[3]:
                # The row now spans exactly the area below it: they are one area, where the row stood.
                under[0] = row
                for at in range(right[1], right[3] + 1):
                    owner[(row, at)] = under
                stood = next(index for index, box in enumerate(ordered) if box is right)
                was = next(index for index, box in enumerate(ordered) if box is under)
                ordered[stood] = under
                del ordered[was]
            continue
        box = [row, column, row, column]
        owner[(row, column)] = box
        ordered.insert(0, box)
    return ordered


def _visible(target: Range) -> list[Area]:
    """Cells not in hidden rows or columns: each area's visible row bands by its visible column bands."""
    sheet = target.sheet
    dims = sheet.dims
    if dims.zero_height:
        raise VBAUnsupportedError("SpecialCells for visible cells on a sheet whose rows are hidden by default "
                                  "is not implemented")
    areas = target.areas
    if target.single:
        area = target.first
        if not (dims.row_hidden(area.top) or dims.column_hidden(area.left)):
            return [area]
        # A hidden single cell answers for the whole sheet.
        areas = [Area(1, 1, MAX_ROWS, MAX_COLUMNS, sheet.name)]
    hidden_rows = {row for row, record in dims.rows.items() if record.hidden}
    hidden_columns = {column for column, record in dims.columns.items() if record.hidden}
    out: list[Area] = []
    for area in areas:
        rows = bands(area.top, area.bottom, hidden_rows)
        columns = bands(area.left, area.right, hidden_columns)
        out += [Area(top, left, bottom, right, sheet.name) for top, bottom in rows for left, right in columns]
    if not out:
        raise error(1004, "No cells were found.")
    return out
