"""Range.AutoComplete: what Excel offers to finish a beginning typed into a cell.

Measured in live Excel (scripts/measure_range_text_helpers.py,
tests/fixtures/range_text_helpers.json).

- The candidates are the text in the cells of the cell's column that run
  without a gap from just above it and from just below it, top to bottom;
  numbers are none, and an empty cell ends a run.
- A candidate matches where it begins with the text given, case aside.
  One match, or several that differ only in case, gives the first's text;
  none, several, or "" given, gives "".
- Application.EnableAutoComplete makes no difference. A range of more
  than one cell is error 1004.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_ROWS
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, error, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


def auto_complete(target: Range, beginning: object) -> str:
    if len(target.areas) != 1 or not target.single:
        raise error(1004, "Unable to get the AutoComplete property of the Range class")
    typed = to_text(beginning).casefold()
    if not typed:
        return ""
    sheet, row, column = target.sheet, target.first.top, target.first.left
    above = _run(sheet, range(row - 1, 0, -1), column)
    below = _run(sheet, range(row + 1, MAX_ROWS + 1), column)
    matches = [text for text in [*reversed(above), *below] if text.casefold().startswith(typed)]
    if not matches or any(text.casefold() != matches[0].casefold() for text in matches):
        return ""
    return matches[0]


def _run(sheet: Worksheet, rows: range, column: int) -> list[str]:
    """The text of the cells from the first of ``rows`` on until an empty one, nearest first."""
    found: list[str] = []
    for row in rows:
        cell = sheet.cell(row, column)
        if cell is None or cell.value is EMPTY and not cell.formula:
            break
        if cell.formula:
            raise VBAUnsupportedError("AutoComplete beside a formula is not implemented")
        if isinstance(cell.value, str):
            found.append(cell.value)
    return found
