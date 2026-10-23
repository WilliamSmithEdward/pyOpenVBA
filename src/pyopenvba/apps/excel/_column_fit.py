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
  by character, a ``*`` fill as nothing, the space an ``_x`` leaves as wide
  as ``x`` and the one a ``?`` leaves for a digit as wide as a 0.
- A text ``t`` pixels wide makes the column ceil(t) + 7 pixels, or
  floor(t * 29 / 28 + 5) where that is wider, and at most 1790 pixels,
  255 characters.
- Every cell of the column counts, those in hidden and filtered rows
  among them; a merged cell does not, nor one that shows nothing, and a
  column with no other keeps its width. A hidden column is shown. A file
  keeps the width as the column's own, bestFit="1" beside it.
- Range.Columns.AutoFit fits each column to the range's own cells.
- A number format or style a macro sets widens a column the same way
  wherever that makes it wider (see widen).

Text that wraps, is indented or turned, a font or size other than Aptos
Narrow or Calibri at 11pt and characters past Latin-1 are not measured,
and AutoFit over them reports itself unsupported.

What a cell shows in a column too narrow for it, Range.Text, is here too
(see narrowed): it rests on the same widths.
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
    from pyopenvba.formula._display import Display

#: How much further italic text reaches, in pixels, by font.
_ITALIC_REACH = {"aptos narrow": Fraction(591, 100), "calibri": Fraction(79, 20)}
#: The widest column AutoFit makes, in pixels: 255 characters.
WIDEST = 1790
#: Alignments that wrap or repeat a cell's text, which AutoFit was not measured over.
_UNMEASURED_ALIGNMENTS = ("justify", "distributed", "fill", "centerContinuous")


def column_pixels(text: Fraction) -> int:
    """The width AutoFit makes a column for a text this many pixels wide."""
    return min(WIDEST, max(math.ceil(text) + 7, math.floor(text * Fraction(29, 28) + 5)))


# --- what a cell shows in a column too narrow for it --------------------------------------------------------------

#: How much further italic text reaches where a cell shows it, by font, in pixels: half what AutoFit allows, rounded
#: up (tests/fixtures/narrow_text.json).
_ITALIC_SHOWN = {"aptos narrow": 3, "calibri": 2}
#: The pixels an indent level takes from what a cell has to show its text in.
_INDENT = 9
#: Scientific notation needs this many characters: 1E+05.
_SHORTEST_SCIENTIFIC = 5
#: Range.Text holds at most this many characters, a fill's among them.
_LONGEST = 255


def narrowed(sheet: Worksheet, row: int, column: int, cell: Cell) -> str:
    """What Range.Text shows of a cell: its value through its number format, in a column perhaps too narrow for it
    (tests/fixtures/narrow_text.json).

    A value its format cannot show, a date before 1900 say, fills the column with #. A number shown by General loses decimals, then goes to scientific notation, then shows 0 where it rounds to
    that, as the column narrows, the format's other text kept; any other number, a date, a logical or an error
    fills the column with # once it no longer fits, as does a General number nothing else fits. Text shows whole,
    unless a section with a ``*`` fill lays it out, and then fills with # as a number would. A fill repeats its
    character as often as it fits, to 255 characters in all, each ``_x`` and ``?`` counted a space wide. What a
    cell has to show its text in is its column's width but 5 pixels, an indent level's 9 and italic's reach; a
    merged cell's is its columns'; a cell that shrinks its text shows it all, and its fill not at all."""
    from pyopenvba.apps.excel import _merges
    from pyopenvba.apps.excel._dimensions import characters_read
    from pyopenvba.apps.excel._font_widths import widths
    from pyopenvba.formula._values import ExcelError
    from pyopenvba.interpreter._values import to_text

    value, shown, undisplayable = _laid_out(sheet, cell.value, cell.number_format)
    if value is EMPTY:
        return ""
    if shown is not None:
        text = shown.text
    elif undisplayable:
        # As many # as the column is characters wide where nothing below measures it (measured at 40 and 60).
        text = "#" * int(characters_read(sheet.dims.column_shown_pixels(column)))
    else:
        text = value.name if isinstance(value, ExcelError) else to_text(value)
    if isinstance(value, str) and (shown is None or not shown.fill):
        return text
    style = cell.style or sheet.book.stylesheet.default
    font, alignment = style.font, style.alignment
    if alignment.shrink:
        return text
    merged = _merges.at(sheet, row, column) if sheet.merged_areas else None
    dims = sheet.dims
    if merged is not None:
        if (row, column) != (merged.top, merged.left):
            return text
        pixels = sum(dims.column_shown_pixels(one) for one in range(merged.left, merged.right + 1))
    elif dims.column_hidden(column):
        raise VBAUnsupportedError("Range.Text of a value in a hidden column is not implemented: what Excel shows "
                                  "there depends on what it drew there before")
    else:
        pixels = dims.column_shown_pixels(column)
    fill = shown.fill if shown is not None else ""
    table = widths(font.name, font.bold, font.italic) if float(font.size) == 11.0 and not font.vert_align else None
    if table is None or alignment.text_rotation or any(char not in table for char in text + fill):
        # Not measured, but wide enough for any font's characters an em and a quarter wide.
        em = round(float(font.size) * 4 / 3)
        if not fill and len(text) * em * 5 / 4 + 5 + _INDENT * alignment.indent + em <= pixels:
            return text
        raise VBAUnsupportedError(f"Range.Text of a value in {font.name} {font.size:g}, which may not fit its column "
                                  f"or fills it, is not implemented: Excel measures Aptos Narrow and Calibri at 11pt")
    reach = _ITALIC_SHOWN[font.name.lower()] if font.italic else 0
    room = pixels - 5 - reach
    space = room - _INDENT * alignment.indent
    hashes = "#" * max(0, room // table["#"])
    if shown is None:
        return hashes if undisplayable or _glyphs(table, text) > space else text
    laid, fill_at = shown.text, shown.fill_at
    if shown.general is not None and isinstance(value, (int, float)) and not isinstance(value, bool):
        start, end = shown.general
        piece = laid[start:end]
        spare = space - _glyphs(table, laid[:start] + laid[end:])
        if _shown_width(table, piece) > spare:
            found = _general_within(shown.general_value, piece, spare, table)
            if found is None:
                return hashes
            if fill_at >= end:
                fill_at += len(found) - len(piece)
            laid = laid[:start] + found + laid[end:]
    elif _glyphs(table, laid) > space:
        return hashes
    if fill:
        count = max(0, min((space - _glyphs(table, laid)) // table[fill], _LONGEST - len(laid)))
        laid = laid[:fill_at] + fill * count + laid[fill_at:]
    return laid


def _laid_out(sheet: Worksheet, value: object, code: str) -> tuple[object, Display | None, bool]:
    """A cell's value, a date as its serial; how its format lays it out, None for a value it does not lay out; and
    whether that is because the format cannot show it."""
    from pyopenvba.formula._display import UndisplayableError, displayed
    from pyopenvba.interpreter._values import VBADate

    if isinstance(value, VBADate):
        value = value.serial
    if value is EMPTY or not isinstance(value, (bool, str, int, float)):
        return value, None, False
    try:
        return value, displayed(value, code, epoch_1904=sheet.book.epoch_1904), False
    except UndisplayableError:
        return value, None, True


def _glyphs(table: dict[str, int], text: str) -> int:
    """How wide a text is, its characters side by side."""
    return sum(table[char] for char in text)


def _shown_width(table: dict[str, int], text: str) -> int:
    """How wide a number in General is where a cell shows it: as wide as its characters, but at least its sign and
    a digit's width for every other character (tests/fixtures/narrow_text.json)."""
    signed = text.startswith("-")
    return max(sum(table[char] for char in text), (table["-"] if signed else 0) + table["0"] * (len(text) - signed))


def _significant(text: str) -> int:
    """How many significant digits a General number shows: from its first digit that is not 0 to its last."""
    digits = text.partition("E")[0].replace("-", "").replace(".", "").strip("0")
    return len(digits)


def _general_within(value: float, full: str, space: int, table: dict[str, int]) -> str | None:
    """The General number that fits ``space`` pixels: of the most decimals of it that keep a digit that is not 0,
    and the most digits of it in scientific notation, the one showing more digits, the decimals where they show as
    many; else 0 where it rounds to 0 and a scientific one would not have room; None where nothing does."""
    fixed = scientific = None
    if "E" not in full:
        decimals = len(full.partition(".")[2])
        fixed = next((shown for places in range(decimals, -1, -1) for shown in [_fixed(value, places)]
                      if _significant(shown) and _shown_width(table, shown) <= space), None)
    scientific = next((shown for places in range(9, -1, -1) for shown in [_scientific(value, places)]
                       if _shown_width(table, shown) <= space), None)
    if fixed is not None and (scientific is None or _significant(fixed) >= _significant(scientific)):
        return fixed
    if scientific is not None:
        return scientific
    zero = _fixed(value, 0)
    room = (space - (table["-"] if value < 0 else 0)) // table["0"]
    if zero.lstrip("-") == "0" and 1 <= room < _SHORTEST_SCIENTIFIC:
        return zero
    return None


def _fixed(value: float, places: int) -> str:
    """``value`` to ``places`` decimals, halves away from 0, as General spells it: no trailing 0s, a sign kept on a
    negative that rounds to 0 (-0)."""
    from decimal import ROUND_HALF_UP, Decimal, localcontext

    with localcontext() as context:
        # Room for every digit of a double's largest, 1.8E+308, and its decimals.
        context.prec = 400
        rounded = Decimal(f"{value:.15g}").quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    shown = f"{abs(rounded):f}"
    if "." in shown:
        shown = shown.rstrip("0").rstrip(".")
    return ("-" if value < 0 else "") + shown


def _scientific(value: float, places: int) -> str:
    """``value`` in scientific notation with ``places`` decimals, halves away from 0, as General spells it."""
    from decimal import ROUND_HALF_UP, Decimal, localcontext

    number = abs(Decimal(f"{value:.15g}"))
    if not number:
        return "0"
    exponent = number.adjusted()
    with localcontext() as context:
        context.prec = 400
        mantissa = number.scaleb(-exponent).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
        if mantissa >= 10:
            exponent += 1
            mantissa = number.scaleb(-exponent).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
    digits = f"{mantissa:f}"
    if "." in digits:
        digits = digits.rstrip("0").rstrip(".")
    sign = "-" if value < 0 else ""
    return f"{sign}{digits}E{'+' if exponent >= 0 else '-'}{abs(exponent):02d}"


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


def _shows(sheet: Worksheet, row: int, column: int, cell: Cell) -> tuple[object, Display] | None:
    """A cell's value, a date as its serial, and how its format lays it out; None for a cell that shows nothing."""
    from pyopenvba.formula._display import Display
    from pyopenvba.formula._values import ExcelError
    from pyopenvba.interpreter._values import to_text

    value = sheet.book.calculator.value_of(sheet.name, row, column) if cell.formula else cell.value
    code = sheet.style_at(row, column).number_format
    if sheet.conditional_formats:
        from pyopenvba.apps.excel._conditional_eval import conditional_number_format

        code = conditional_number_format(sheet, row, column)
    number, shown, undisplayable = _laid_out(sheet, value, code)
    if undisplayable:
        raise VBAUnsupportedError("measuring a value its format cannot show is not implemented")
    if number is EMPTY:
        return None
    if shown is None:
        shown = Display(number.name if isinstance(number, ExcelError) else to_text(number))
    return (number, shown) if shown.text else None


def _measured(table: dict[str, int], shown: Display) -> int:
    """How wide AutoFit measures a text a format laid out: each space an _ leaves as wide as the character it
    stands for and a fill as nothing (tests/fixtures/narrow_text.json, cell_style), and what General shows of a
    number a digit wide for each digit, point and exponent sign, its sign and E as they are."""
    total = _glyphs(table, shown.text) + sum(table[pad] - table[" "] for pad in shown.pads)
    if shown.general is not None:
        piece = shown.text[shown.general[0]:shown.general[1]]
        signed = piece.startswith("-")
        total += (table["-"] if signed else 0) - _glyphs(table, piece) + sum(
            table[char] if char == "E" else table["0"] for char in piece[1 if signed else 0:])
    return total


def _width(sheet: Worksheet, row: int, column: int, cell: Cell, *, formatted: bool = False) -> Fraction | None:
    """How wide AutoFit measures the text a cell shows, in pixels; None for a cell that shows nothing. ``formatted``
    measures only what a format set on the cell widens its column for: a number General does not show, wrapped
    or not, as numbers do not wrap."""
    from pyopenvba.apps.excel._font_widths import widths

    found = _shows(sheet, row, column, cell)
    if found is None:
        return None
    number, shown = found
    if formatted and (not isinstance(number, (int, float)) or isinstance(number, bool) or shown.general is not None):
        return None
    style = sheet.style_at(row, column)
    font, alignment = style.font, style.alignment
    wraps = alignment.wrap and not formatted
    if wraps or alignment.indent or alignment.text_rotation or alignment.horizontal in _UNMEASURED_ALIGNMENTS:
        raise VBAUnsupportedError("AutoFit of a column whose text wraps, is indented or is turned is not implemented")
    table = widths(font.name, font.bold, font.italic) if float(font.size) == 11.0 and not font.vert_align else None
    if table is None:
        raise VBAUnsupportedError(f"AutoFit of text in {font.name} {font.size:g} is not implemented: AutoFit measures "
                                  f"Aptos Narrow and Calibri at 11pt")
    if "\n" in shown.text or "\r" in shown.text:
        raise VBAUnsupportedError("AutoFit of text holding a line break is not implemented")
    if any(char not in table for char in shown.text + shown.pads):
        raise VBAUnsupportedError("AutoFit of text holding characters past Latin-1 is not implemented: GDI's widths "
                                  "for them depend on the characters beside them")
    reach = _ITALIC_REACH[font.name.lower()] if font.italic else Fraction(0)
    return Fraction(_measured(table, shown)) + reach


def widen(sheet: Worksheet, areas: list[Area]) -> None:
    """A macro gave the cells of ``areas`` a number format or a style: where AutoFit would make a column wider for
    a number one of them now shows, Excel widens the column to that, even where the number shows whole already;
    not a column sized by hand or hidden, nor for a merged cell, and never narrower. A number General shows, text,
    logicals and errors widen nothing (tests/fixtures/narrow_text.json, widening). A file keeps the width as
    AutoFit's, bestFit="1" beside it.

    A number AutoFit does not measure -- in a font or size other than Aptos Narrow and Calibri at 11pt, indented
    or turned, or holding characters past Latin-1 -- leaves its column as it is, where Excel may widen it."""
    from pyopenvba.apps.excel import _merges

    if not sheet.dims.normal_known():
        return
    dims = sheet.dims
    widest: dict[int, int] = {}
    for row, column, cell in _cells_in(sheet, areas):
        if dims.column_hidden(column):
            continue
        record = dims.columns.get(column)
        if record is not None and record.custom and record.extra.get("bestFit") != "1":
            continue
        if sheet.merged_areas and _merges.at(sheet, row, column) is not None:
            continue
        try:
            width = _width(sheet, row, column, cell, formatted=True)
        except VBAUnsupportedError:
            continue
        if width is not None:
            widest[column] = max(widest.get(column, 0), column_pixels(width))
    for column, pixels in widest.items():
        if pixels > dims.column_shown_pixels(column):
            dims.fit_column(column, pixels)
            sheet.touched()


def _cells_in(sheet: Worksheet, areas: list[Area]) -> list[tuple[int, int, Cell]]:
    """The cells of ``areas`` that hold something, each once: looked up where the areas are smaller than the sheet's
    cells, and picked from them where they are not."""
    cells = sheet.cells_
    if sum(area.rows * area.columns for area in areas) > len(cells):
        return [(row, column, cell) for (row, column), cell in cells.items()
                if any(area.contains(row, column) for area in areas)]
    found: dict[tuple[int, int], Cell] = {}
    for area in areas:
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                cell = cells.get((row, column))
                if cell is not None:
                    found[row, column] = cell
    return [(row, column, cell) for (row, column), cell in found.items()]
