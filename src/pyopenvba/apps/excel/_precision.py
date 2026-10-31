"""Workbook.PrecisionAsDisplayed: every number a cell holds kept as its format shows it
(tests/fixtures/precision_as_displayed.json, scripts/measure_precision_as_displayed.py).

Turned on, it rounds every number in every sheet once, constants and what formulas came to; after that a number
written, a format given and a formula worked out are each rounded as they come. Turned off, the numbers stay as
they were rounded. The file says so with calcPr fullPrecision="0". The rounding itself is formula._display's
precision_shown: General, text, dates, times and fractions are kept whole.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook, Worksheet


def shown(sheet: Worksheet, row: int, column: int, value: object) -> object:
    """``value`` as the cell's format shows it, where the workbook keeps numbers so."""
    if not sheet.book.precision_as_displayed or not isinstance(value, float):
        return value
    from pyopenvba.formula._display import precision_shown

    rounded = precision_shown(value, sheet.style_at(row, column).number_format)
    return value if rounded is None else rounded


def settle(sheet: Worksheet, row: int, column: int) -> None:
    """Round what a cell holds, a constant or what its formula came to, as its format shows it."""
    if not sheet.book.precision_as_displayed:
        return
    cell = sheet.cells_.get((row, column))
    if cell is not None and isinstance(cell.value, float):
        cell.value = shown(sheet, row, column, cell.value)


def turned_on(book: Workbook) -> None:
    """Every number of every sheet rounded, as turning PrecisionAsDisplayed on rounds them; a formula reading a
    constant that changed is worked out again, at once or when manual calculation next comes round."""
    for sheet in book.sheets_:
        for (row, column), cell in list(sheet.cells_.items()):
            before = cell.value
            settle(sheet, row, column)
            if not cell.formula and cell.value != before:
                sheet.cell_changed(row, column)
        sheet.touched()
