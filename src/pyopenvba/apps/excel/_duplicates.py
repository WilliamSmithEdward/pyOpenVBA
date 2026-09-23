"""Range.RemoveDuplicates, as Excel removes duplicate rows.

Measured in live Excel (scripts/measure_remove_duplicates.py):

- A row goes when the columns asked for hold what an earlier row's hold;
  the first of each stays. The rows left close up from the top of the
  range, formats and all, their formulas shifting as a copy's would, and
  the rows freed at the bottom are cleared. Nothing outside the range
  moves, and a reference from outside keeps pointing where it did.
- Two cells match when they are both blank, both the same error, both
  text that is the same ignoring case, or both numbers of equal value
  that show the same -- 1 and 1.00 differ, 0.3 and 0.1 + 0.2 differ, and
  TRUE counts as a number 1 that shows TRUE. Text never matches a number
  however it shows. Formulas match by what they give.
- Text compares as Windows compares words ignoring case, so æ matches ae
  and ß ss; that is reproduced for printable ASCII, and other text
  reports itself unsupported.
- Header is xlNo when left out; xlGuess takes the first row for a header
  as Sort's xlGuess does. A single cell works on its current region.
- Columns counts from the range's first column. Left out, or 0, nothing
  happens; outside the range it is error 1004. An array of them compares
  them all, an empty one does nothing, and one outside the range, 0
  included, is error 5.
- Over a filter's range (scripts/measure_autofilter_edits.py) every row
  counts, hidden or not. Over the whole range the range gives up the
  rows that went, and those left below it show; over all of it but its
  header the range stays. Either way a filter with criteria filters
  again. Other overlaps with a filter's range report themselves.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import _merges
from pyopenvba.apps.excel._sort import GUESS, NO, YES, guessed_header
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import UndisplayableError, shown
from pyopenvba.formula._parse import shift_text
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, VBADate, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def remove_duplicates(target: Range, columns: object, header: object) -> object:
    from pyopenvba.apps.excel._region import current_region

    if len(target.areas) != 1:
        raise error(1004, "RemoveDuplicates works on one block")
    sheet = target.sheet
    area = current_region(sheet, target.first.top, target.first.left) if target.single else target.first
    kept = NO if header is MISSING else int(to_integer(header, "Long"))
    if kept not in (GUESS, YES, NO):
        raise error(1004, "Header takes xlYes, xlNo or xlGuess")
    picked = _columns(columns, area.columns)
    if not picked:
        return EMPTY
    if any(_merges.intersects(area, one) for one in sheet.merged_areas):
        raise VBAUnsupportedError("RemoveDuplicates over merged cells is not implemented")
    dims = sheet.dims
    if any(area.top <= row <= area.bottom and record.style is not None for row, record in dims.rows.items()) \
            or any(dims.column_style(column) is not None for column in range(area.left, area.right + 1)):
        raise VBAUnsupportedError("RemoveDuplicates on rows or columns with formats of their own is not implemented")
    used = sheet.used_bounds()
    if used is None:
        return EMPTY
    box = sheet.auto_filter.area if sheet.auto_filter is not None else None
    overlaps = box is not None and box.top <= area.bottom and area.top <= box.bottom and \
        box.left <= area.right and area.left <= box.right
    if box is not None and overlaps and ((area.left, area.right, area.bottom) != (box.left, box.right, box.bottom)
                                         or area.top not in (box.top, box.top + 1)):
        # Measured over a filter's whole range and over all of it but its header; nothing else.
        raise VBAUnsupportedError("RemoveDuplicates over part of a filter's range is not implemented")
    first = area.top + (1 if kept == YES or (kept == GUESS and guessed_header(sheet, area, False)) else 0)
    last = min(area.bottom, used[2])
    seen: set[tuple[tuple[object, ...], ...]] = set()
    rows: list[int] = []
    for row in range(first, last + 1):
        found = tuple(_key(sheet, row, area.left + index - 1) for index in picked)
        if found not in seen:
            seen.add(found)
            rows.append(row)
    if len(rows) == last - first + 1:
        return EMPTY
    destination = {row: first + place for place, row in enumerate(rows)}
    block = Area(first, area.left, last, area.right)
    moving = {position: cell for position, cell in sheet.cells_.items() if block.contains(*position)}
    for position in moving:
        del sheet.cells_[position]
    for (row, column), cell in moving.items():
        if row not in destination:
            continue
        down = destination[row] - row
        if cell.formula and down:
            cell.formula = shift_text(cell.formula, down, 0)
            cell.stale, cell.value = True, EMPTY
        sheet.cells_[(row + down, column)] = cell
    if box is not None and overlaps:
        from pyopenvba.apps.excel._autofilter import filtered_again

        # Over the whole range the range gives up the rows that went; either way the filter applies again.
        filtered_again(sheet, last - first + 1 - len(rows) if area.top == box.top else 0)
    sheet.touched()
    sheet.book.calculator.rebuild()
    return EMPTY


def _columns(columns: object, width: int) -> list[int]:
    """The range's columns to compare, counted from 1; empty where nothing is to happen."""
    if columns is MISSING:
        return []
    if isinstance(columns, VBAArray):
        picked = [int(to_integer(item, "Long")) for item in columns.elements()]
        if any(not 1 <= index <= width for index in picked):
            raise error(5)
        return picked
    index = int(to_integer(columns, "Long"))
    if index == 0:
        return []
    if not 1 <= index <= width:
        raise error(1004, "Columns names a column outside the range")
    return [index]


def _key(sheet: Worksheet, row: int, column: int) -> tuple[object, ...]:
    """What a cell holds, as RemoveDuplicates compares it."""
    cell = sheet.cells_.get((row, column))
    if cell is None:
        return ("blank",)
    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY:
        return ("blank",)
    if isinstance(value, ExcelError):
        return ("error", value.name)
    if isinstance(value, bool):
        return ("number", float(value), "true" if value else "false")
    if isinstance(value, (int, float)):
        try:
            text, _ = shown(float(value), cell.number_format)
        except UndisplayableError:
            text = "#"
        return ("number", float(value), text.lower())
    text = str(value)
    if any(not 32 <= ord(char) <= 126 for char in text):
        raise VBAUnsupportedError("RemoveDuplicates on text beyond printable ASCII is not implemented")
    return ("text", text.lower())
