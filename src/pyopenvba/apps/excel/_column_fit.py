"""Column AutoFit: how wide Excel makes a column for the text its cells show, on a 96-DPI display.

Measured in live Excel (scripts/measure_column_autofit.py,
tests/fixtures/column_autofit.json):

- A text is as wide as its characters, each as wide as GDI makes it in
  the cell's font (see _font_widths), whole pixels that add up along the
  text for ASCII and Latin-1. Italic reaches 5.91 pixels further in Aptos
  Narrow and 3.95 in Calibri: any value from 5.90 to 5.93, and from 3.94
  to 3.96, fits all 240 italic texts measured in each.
- A number in General is measured a digit wide for each digit, decimal
  point and exponent sign, its leading minus and its E as wide as they
  are; in any other format, as text, booleans and errors are, character
  by character.
- A text ``t`` pixels wide makes the column ceil(t) + 7 pixels, or
  floor(t * 29 / 28 + 5) where that is wider, and at most 1790 pixels,
  255 characters.
- Every cell of the column counts, those in hidden and filtered rows
  among them; a merged cell does not, nor one that shows nothing, and a
  column with no other keeps its width. A hidden column is shown. A file
  keeps the width as the column's own, bestFit="1" beside it.
- Range.Columns.AutoFit fits each column to the range's own cells.

Text that wraps, is indented or turned, a font or size other than Aptos
Narrow or Calibri at 11pt, a ``*`` fill and characters past Latin-1 are
not measured, and AutoFit over them reports itself unsupported.
"""
from __future__ import annotations

import math
from fractions import Fraction
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_ROWS, Area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Worksheet

#: How much further italic text reaches, in pixels, by font.
_ITALIC_REACH = {"aptos narrow": Fraction(591, 100), "calibri": Fraction(79, 20)}
#: The widest column AutoFit makes, in pixels: 255 characters.
WIDEST = 1790
#: Alignments that wrap or repeat a cell's text, which AutoFit was not measured over.
_UNMEASURED_ALIGNMENTS = ("justify", "distributed", "fill", "centerContinuous")


def column_pixels(text: Fraction) -> int:
    """The width AutoFit makes a column for a text this many pixels wide."""
    return min(WIDEST, max(math.ceil(text) + 7, math.floor(text * Fraction(29, 28) + 5)))


def fit(sheet: Worksheet, area: Area) -> None:
    """AutoFit of the columns of ``area``, each fitted to its cells in the area's rows."""
    if not sheet.dims.normal_known():
        normal = sheet.book.stylesheet.default.font
        raise VBAUnsupportedError(f"AutoFit of a column in a workbook whose Normal font is {normal.name} "
                                  f"{normal.size:g} is not implemented")
    whole = area.top == 1 and area.bottom == MAX_ROWS
    widest: dict[int, int] = {}
    from pyopenvba.apps.excel import _merges

    for (row, column), cell in list(sheet.cells_.items()):
        if not area.left <= column <= area.right or not (whole or area.top <= row <= area.bottom):
            continue
        if sheet.merged_areas and _merges.at(sheet, row, column) is not None:
            continue
        width = _width(sheet, row, column, cell)
        if width is not None:
            widest[column] = max(widest.get(column, 0), column_pixels(width))
    for column, pixels in widest.items():
        sheet.dims.fit_column(column, pixels)
    if widest:
        sheet.touched()


def _width(sheet: Worksheet, row: int, column: int, cell: Cell) -> Fraction | None:
    """How wide the text a cell shows is, in pixels; None for a cell that shows nothing."""
    from pyopenvba.apps.excel._font_widths import widths
    from pyopenvba.formula._display import UndisplayableError, shown
    from pyopenvba.formula._values import ExcelError
    from pyopenvba.interpreter._values import VBADate, to_text

    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    if value is EMPTY:
        return None
    style = sheet.style_at(row, column)
    font, alignment = style.font, style.alignment
    if alignment.wrap or alignment.indent or alignment.text_rotation or alignment.horizontal in _UNMEASURED_ALIGNMENTS:
        raise VBAUnsupportedError("AutoFit of a column whose text wraps, is indented or is turned is not implemented")
    table = widths(font.name, font.bold, font.italic) if float(font.size) == 11.0 and not font.vert_align else None
    if table is None:
        raise VBAUnsupportedError(f"AutoFit of text in {font.name} {font.size:g} is not implemented: AutoFit measures "
                                  f"Aptos Narrow and Calibri at 11pt")
    code = style.number_format
    if sheet.conditional_formats:
        from pyopenvba.apps.excel._conditional_eval import conditional_number_format

        code = conditional_number_format(sheet, row, column)
    number = value.serial if isinstance(value, VBADate) else value
    if isinstance(number, ExcelError):
        text = number.name
    elif isinstance(number, (bool, str, int, float)):
        try:
            text, fills = shown(number, code, epoch_1904=sheet.book.epoch_1904)
        except UndisplayableError:
            raise VBAUnsupportedError("AutoFit of a value its format cannot show is not implemented") from None
        if fills:
            raise VBAUnsupportedError("AutoFit of a number format with a * fill is not implemented")
    else:
        text = to_text(number)
    if not text:
        return None
    if "\n" in text or "\r" in text:
        raise VBAUnsupportedError("AutoFit of text holding a line break is not implemented")
    if any(char not in table for char in text):
        raise VBAUnsupportedError("AutoFit of text holding characters past Latin-1 is not implemented: GDI's widths "
                                  "for them depend on the characters beside them")
    general = isinstance(number, (int, float)) and not isinstance(number, bool) and code.lower() == "general"
    if general:
        # A number in General: each digit, point and exponent sign a digit wide, its sign and E as they are.
        signed = text.startswith("-")
        total = (table["-"] if signed else 0) + sum(
            table[char] if char == "E" else table["0"] for char in text[1 if signed else 0:])
    else:
        total = sum(table[char] for char in text)
    reach = _ITALIC_REACH[font.name.lower()] if font.italic else Fraction(0)
    return Fraction(total) + reach
