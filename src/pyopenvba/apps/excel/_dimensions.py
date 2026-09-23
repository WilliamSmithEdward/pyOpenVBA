"""Row heights, column widths, and hidden rows and columns.

Excel keeps a row's height in quarters of a pixel and a column's width in
whole pixels of the display it runs on, and rounds what a macro asks for
to them. The model is Excel on a 96-DPI display (Windows at 100%) with
Aptos Narrow 11 as the Normal font: a pixel is 0.75pt, a standard row is
20 pixels (15pt), a standard column 64 pixels (8.43 characters), the
widest digit 7 pixels and a column's padding 5. Every rule here was
measured in live Excel on such a display (scripts/measure_dimensions.py
and scripts/measure_dimension_file.py):

- A height a macro sets is rounded to twips, half up, and then to quarter
  pixels. RowHeight reads it back to the nearest quarter point, Height
  counts whole pixels, and the file stores it rounded up to twips, spelled
  the way :func:`~pyopenvba.apps.excel._styles.excel_number` spells a
  measurement.
- A width is rounded to whole pixels, half up: 7 to the character plus 5,
  or 12 to the character below one. ColumnWidth reads it back to two
  decimals, and the file stores ``pixels * 256 / 7`` truncated to 1/256.
- A height of zero hides a row, and a width of zero a column; each keeps
  the size it had for when it is shown again. A standard column hidden
  this way is written with a width of 0 and comes back at the standard
  width, now its own.
- A read across several rows or columns compares only the part that lies
  inside the sheet's used cells (see :func:`_compared`), which is why the
  same read comes back Null on one sheet and not on another.
"""

from __future__ import annotations

import copy
import math
import re
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from fractions import Fraction
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba._xml import attributes as _attributes
from pyopenvba._xml import escape as _escape
from pyopenvba.apps.excel import _font_rows
from pyopenvba.apps.excel._styles import excel_number
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL, error, to_number

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Range, Worksheet
    from pyopenvba.apps.excel._styles import Side, Style

#: Points in a pixel on the display the model emulates.
PIXEL = 0.75
#: Pixels in the widest digit of the Normal font, and in a column's padding.
DIGIT = 7
PADDING = 5
#: A standard row, in quarter pixels: 20 pixels, 15pt.
STANDARD_ROW = 80
#: A standard column, in pixels: 8.43 characters.
STANDARD_COLUMN = 64
#: The tallest row a macro may ask for (409.5pt) and the tallest Excel writes (409.6pt), in twips.
MOST_TWIPS = 8190
MOST_STORED_TWIPS = 8192
#: The widest column a macro may ask for, in characters.
MOST_CHARACTERS = 255
#: Where a row's descent is written, when the sheet declares Excel 2010's namespace.
_DESCENT = "x14ac:dyDescent"
#: The Normal fonts whose metrics on this display were measured: Aptos
#: Narrow 11, today's, and Calibri 11, the default before it. Both make a
#: 20-pixel row, a 7-pixel digit and a 0.25pt descent.
_KNOWN_NORMAL = {("aptos narrow", 11.0), ("calibri", 11.0)}
#: The descent Excel writes for a row in the Normal font.
NORMAL_DESCENT = "0.25"
#: Alignments that wrap a cell's text onto more lines, as WrapText does.
_WRAPPING = ("justify", "distributed")

_ROW_ORDER = ("r", "spans", "s", "customFormat", "ht", "hidden", "customHeight", "outlineLevel", "collapsed",
              "thickTop", "thickBot", "ph")
_COLUMN_ORDER = ("min", "max", "width", "style", "hidden", "bestFit", "customWidth", "phonetic", "outlineLevel",
                 "collapsed")
_FORMAT_ORDER = ("baseColWidth", "defaultColWidth", "defaultRowHeight", "customHeight", "zeroHeight", "thickTop",
                 "thickBottom", "outlineLevelRow", "outlineLevelCol")
#: A row's own attributes, which the model works out; everything else it carries.
#: A row's s means nothing without customFormat, and Excel drops it.
_ROW_OWN = {"r", "spans", "s", "customFormat", "ht", "hidden", "customHeight", "thickTop", "thickBot", _DESCENT}
_COLUMN_OWN = {"min", "max", "width", "style", "hidden", "customWidth"}


# --- the pixel rules -------------------------------------------------------------------------


def _c_round(value: float) -> int:
    """Round half up, the way ``(int)(x + 0.5)`` does in C."""
    return math.floor(value + 0.5)


def quarters_for(points: float) -> int:
    """The quarter pixels a height a macro sets comes to: through twips, half up at each step."""
    twips = _c_round(points * 20)
    return (twips * 8 + 15) // 30


def quarters_from_file(text: str) -> int | None:
    """The quarter pixels a height in a file comes to; Excel ignores its sign."""
    try:
        points = abs(float(text))
    except ValueError:
        return None
    if math.isnan(points) or math.isinf(points):
        return None
    return quarters_for(points)


def row_points_read(quarters: int) -> float:
    """What RowHeight reads for a row this tall: the nearest quarter point, at most 409.5."""
    return min((3 * quarters + 2) // 4, 1638) / 4


def row_pixels(quarters: int) -> int:
    return quarters // 4


def height_text(quarters: int) -> str:
    """The ht a file stores for a row this tall: rounded up to twips, at most 409.6pt."""
    twips = min((15 * quarters + 3) // 4, MOST_STORED_TWIPS)
    return excel_number(twips / 20)


def pixels_for(characters: float) -> int:
    """The pixels a width a macro sets comes to."""
    if characters < 1:
        return _c_round(characters * (DIGIT + PADDING))
    return _c_round(characters * DIGIT + PADDING)


def pixels_from_file(text: str) -> int | None:
    """The pixels a width in a file comes to, by the formula ECMA-376 gives.

    A negative width wraps round to a very wide column, as a 16-bit pixel
    count does: Excel reads ``width="-3"`` as 65516 pixels.
    """
    try:
        width = float(text)
    except ValueError:
        return None
    if math.isnan(width) or math.isinf(width) or abs(width) > 1e9:
        return None
    sixteenths = int(width * 256)
    return int(Fraction((sixteenths + 128 // DIGIT) * DIGIT, 256)) % 65536


def characters_read(pixels: int) -> float:
    """What ColumnWidth reads for a column this wide: characters, to two decimals."""
    exact = Fraction(pixels, DIGIT + PADDING) if pixels < DIGIT + PADDING else Fraction(pixels - PADDING, DIGIT)
    return float(Fraction(math.floor(exact * 100 + Fraction(1, 2)), 100))


def width_text(pixels: int) -> str:
    """The width a file stores for a column this wide."""
    return excel_number(float(Fraction(pixels * 256 // DIGIT, 256)))


def standard_pixels(base_width: int = 8) -> int:
    """A sheet's standard column when its file names only a base width: rounded up to 8 pixels."""
    pixels = base_width * DIGIT + PADDING
    return (pixels + 7) // 8 * 8


def font_pixels(points: float) -> int:
    """A font's size in pixels, which is what its row height depends on: 12.375pt is 17."""
    return _c_round(points * 4 / 3)


# --- what a sheet holds ------------------------------------------------------------------------


def _no_attributes() -> dict[str, str]:
    return {}


@dataclass(slots=True)
class RowRecord:
    """What a sheet says about one row, beyond its cells."""

    #: The row's own height, in quarter pixels (customHeight in the file).
    height: int | None = None
    hidden: bool = False
    #: The height the row's fonts gave it when the file was saved, which a
    #: file stores without customHeight.
    grown: int | None = None
    #: Every other attribute of the row element (outlineLevel, ...), as read.
    extra: dict[str, str] = field(default_factory=_no_attributes)
    #: The ht and descent the file wrote, reused while they still hold.
    height_as_read: str = ""
    descent: str = ""
    #: The row's own format (customFormat in the file), which every cell
    #: of the row shows that has no format of its own; None when it has
    #: none. The default format is a format too: it hides a column's.
    style: Style | None = None
    #: The xf the file gave the row's format, reused while the format holds.
    xf: int = -1
    #: The thickTop and thickBot the file gave the row, kept where the model cannot work them out.
    flags_as_read: tuple[bool, bool] = (False, False)

    def empty(self) -> bool:
        return (self.height is None and not self.hidden and self.grown is None and not self.extra
                and self.style is None)


@dataclass(slots=True)
class ColumnRecord:
    """What a sheet says about one column."""

    #: Width in pixels; None when the file gives none, which is the standard width.
    width: int | None = None
    custom: bool = False
    hidden: bool = False
    extra: dict[str, str] = field(default_factory=_no_attributes)
    width_as_read: str = ""
    #: The column's format, which its cells show where neither they nor
    #: their row has one; None is the default format.
    style: Style | None = None
    xf: int = -1

    def empty(self) -> bool:
        return self.width is None and not self.custom and not self.hidden and not self.extra and self.style is None

    def attributes(self, standard: str, index_of: Callable[[Style, int], int]) -> dict[str, str]:
        """The col element's attributes; a column with a format but no width of its own states the standard one."""
        out: dict[str, str] = {}
        if self.width is not None:
            out["width"] = self.width_as_read or width_text(self.width)
        elif self.style is not None:
            out["width"] = standard
        if self.style is not None:
            index = index_of(self.style, self.xf)
            if index:
                out["style"] = str(index)
        out.update(self.extra)
        if self.hidden:
            out["hidden"] = "1"
        if self.custom:
            out["customWidth"] = "1"
        return {key: out[key] for key in _ordered(out, _COLUMN_ORDER)}


def _ordered(attrs: dict[str, str], order: tuple[str, ...]) -> list[str]:
    known = [key for key in order if key in attrs]
    return known + [key for key in attrs if key not in order]


class SheetDimensions:
    """A worksheet's row heights, column widths and hidden rows and columns."""

    def __init__(self, sheet: Worksheet) -> None:
        self.sheet = sheet
        self.rows: dict[int, RowRecord] = {}
        self.columns: dict[int, ColumnRecord] = {}
        #: The sheetFormatPr attributes the file had.
        self.format: dict[str, str] = {}
        #: Every row's height, when a macro sized all of them at once (customHeight on sheetFormatPr).
        self.all_rows: int | None = None
        #: Rows nobody gave a record are hidden (zeroHeight on sheetFormatPr).
        self.zero_height = False
        #: The standard column in pixels, once StandardWidth has been set or read from the file.
        self.standard: int | None = None
        #: What changed since the file was read, so an untouched part is written back as it came.
        self.changed_rows: set[int] = set()
        self.columns_changed = False
        self.format_changed = False
        #: Rows were inserted or deleted, so no row's old start tag is where it was.
        self.moved = False
        #: Reading the file dropped something Excel drops when it opens it, so
        #: the sheet is written again rather than kept as it came.
        self.tidied = False
        #: Columns hidden while they kept a width of their own. Excel's used
        #: block takes them in and keeps them, whatever the column does next.
        self.kept_columns: set[int] = set()
        #: Rows their fonts make another height than the standard row, or
        #: give another descent: row -> (quarter pixels, descent in twips).
        self._grown: dict[int, tuple[int, int]] = {}
        #: Rows whose height rests on something not measured, and what.
        self._unmeasured: dict[int, str] = {}
        #: Rows whose cells changed since the file was read, so a height it
        #: recorded for them no longer holds.
        self._touched: set[int] = set()
        #: A column format changed since the file was read, which reaches every row.
        self._columns_touched = False
        #: The standard row the column formats make, or why the model cannot tell.
        self._default_row: tuple[int, int] | None = None
        self._default_reason: str | None = None
        #: The thickTop and thickBot of rows the model worked out, and of the standard row.
        self._flags: dict[int, tuple[bool, bool]] = {}
        self._default_flags: tuple[bool, bool] = (False, False)
        self._grown_valid = False

    # -- reading a file

    def load(self, xml: str) -> None:
        head = re.search(r"<sheetFormatPr\b[^>]*?/?>", xml)
        if head is not None:
            self.format = _attributes(head.group(0))
            if self.format.get("customHeight") in ("1", "true"):
                self.all_rows = quarters_from_file(self.format.get("defaultRowHeight", "15"))
            self.zero_height = self.format.get("zeroHeight") in ("1", "true")
            if "defaultColWidth" in self.format:
                self.standard = pixels_from_file(self.format["defaultColWidth"])
            elif "baseColWidth" in self.format and self.format["baseColWidth"].isdigit():
                self.standard = standard_pixels(int(self.format["baseColWidth"]))
        cols = re.search(r"<cols\b[^>]*>(.*?)</cols>", xml, re.DOTALL)
        if cols is not None:
            for tag in re.findall(r"<col\b[^>]*?/?>", cols.group(1)):
                self._load_column(_attributes(tag))
        data = re.search(r"<sheetData\b[^>]*>(.*?)</sheetData>", xml, re.DOTALL)
        if data is not None:
            for tag in re.findall(r"<row\b[^>]*?/?>", data.group(1)):
                self._load_row(_attributes(tag))
        self._drop_redundant_formats()

    def _style_read(self, text: str) -> tuple[Style, int]:
        stylesheet = self.sheet.book.stylesheet
        index = int(text) if text.isdigit() else 0
        return stylesheet.style(index), index

    def _load_column(self, attrs: dict[str, str]) -> None:
        try:
            first, last = int(attrs.get("min", "0")), int(attrs.get("max", "0"))
        except ValueError:
            return
        # Excel reads a column with no width as one 0 wide, and a column 0 wide as hidden.
        width = pixels_from_file(attrs["width"]) if "width" in attrs else 0
        if "width" not in attrs or (width == 0 and attrs.get("hidden") not in ("1", "true")):
            self.tidied = self.columns_changed = True
        extra = {key: value for key, value in attrs.items() if key not in _COLUMN_OWN}
        style, xf = self._style_read(attrs["style"]) if "style" in attrs else (None, -1)
        for column in range(max(first, 1), min(last, MAX_COLUMNS) + 1):
            self.columns[column] = ColumnRecord(
                width=width, custom=attrs.get("customWidth") in ("1", "true"),
                hidden=attrs.get("hidden") in ("1", "true") or width == 0, extra=dict(extra),
                width_as_read=attrs.get("width", ""), style=style, xf=xf)

    def _load_row(self, attrs: dict[str, str]) -> None:
        try:
            row = int(attrs.get("r", "0"))
        except ValueError:
            return
        if not 1 <= row <= MAX_ROWS:
            return
        record = RowRecord(descent=attrs.get(_DESCENT, ""),
                           flags_as_read=(attrs.get("thickTop") in ("1", "true"), attrs.get("thickBot") in ("1", "true")))
        if attrs.get("customFormat") in ("1", "true"):
            record.style, record.xf = self._style_read(attrs.get("s", "0"))
        elif "s" in attrs:
            self.tidied = True
        custom = attrs.get("customHeight") in ("1", "true")
        if "ht" in attrs:
            quarters = quarters_from_file(attrs["ht"])
            if custom:
                record.height = quarters
            else:
                record.grown = quarters
            record.height_as_read = attrs["ht"]
        elif custom:
            record.height = self.default_quarters()
        record.hidden = attrs.get("hidden") in ("1", "true")
        record.extra = {key: value for key, value in attrs.items() if key not in _ROW_OWN}
        # When every row is hidden by default, a row the file lists is one it shows.
        if not record.empty() or self.zero_height:
            self.rows[row] = record

    def _drop_redundant_formats(self) -> None:
        """What Excel lets go of when it opens a file: formats that change nothing a cell shows.

        A column in the default format has none, and a row format goes when
        every column shows it already. Excel takes a format that differs from
        another only in its apply flags as that format, so a row a macro set
        to not bold, saved and opened again, has no format left.
        """
        default = self.sheet.book.stylesheet.default
        for column, record in list(self.columns.items()):
            if record.style == default:
                record.style, record.xf = None, -1
                self.tidied = self.columns_changed = True
                if record.empty():
                    del self.columns[column]
        shown: set[Style] = set()
        styled = 0
        for record in self.columns.values():
            if record.style is not None:
                shown.add(record.style)
                styled += 1
        if styled < MAX_COLUMNS:
            shown.add(default)
        if len(shown) != 1:
            return
        (everywhere,) = shown
        for row, record in list(self.rows.items()):
            if record.style == everywhere:
                record.style, record.xf = None, -1
                self.tidied = True
                if record.empty() and not self.zero_height:
                    del self.rows[row]

    # -- the standard sizes

    def normal_known(self) -> bool:
        font = self.sheet.book.stylesheet.default.font
        return (font.name.lower(), float(font.size)) in _KNOWN_NORMAL

    def default_quarters(self) -> int:
        """How tall a row with no record is: what a macro sized every row to, or what the column formats make it.

        With no column formats that is the standard row. Where the model
        cannot tell -- a column format in a font it has not measured -- the
        file's height holds until a column format changes.
        """
        if self.all_rows is not None:
            return self.all_rows
        self._settle()
        if self._default_reason is not None and self._columns_touched:
            raise VBAUnsupportedError(f"the height of the sheet's rows rests on {self._default_reason}, "
                                      "which is not measured")
        return self.standard_written()[0]

    def standard_written(self) -> tuple[int, str]:
        """The standard row's height and descent as the sheet part states them."""
        self._settle()
        filed = quarters_from_file(self.format.get("defaultRowHeight", "")) or None
        if self._default_row is not None and self._default_reason is None:
            quarters, descent = self._default_row
            return quarters, excel_number(descent / 20)
        return filed or STANDARD_ROW, self.format.get(_DESCENT, NORMAL_DESCENT)

    def default_descent(self) -> str:
        """The descent a row with nothing taller in it is written with."""
        return self.standard_written()[1]

    def standard_width(self) -> int:
        return self.standard if self.standard is not None else STANDARD_COLUMN

    # -- rows

    def record(self, row: int) -> RowRecord | None:
        return self.rows.get(row)

    def row_quarters(self, row: int) -> int:
        """A row's height when shown, in quarter pixels.

        A row that keeps no height of its own is as tall as the fonts its
        cells show make it, and those come from the cells' own formats, the
        row's or the columns'. Where the model cannot tell -- a font it has
        not measured, wrapped or turned text, a mix of fonts -- it takes the
        height the file recorded while nothing in the row has changed, and
        reports the height unsupported once something has.
        """
        record = self.rows.get(row)
        if record is not None and record.height is not None:
            return record.height
        if self.all_rows is not None:
            return self.all_rows
        self._settle()
        reason = self._unmeasured.get(row)
        if reason is not None:
            if row in self._touched or self._columns_touched:
                raise VBAUnsupportedError(f"the height of row {row} rests on {reason}, which is not measured")
            if record is not None and record.grown is not None:
                return record.grown
            return self.default_quarters()
        grown = self._grown.get(row)
        return grown[0] if grown is not None else self.default_quarters()

    def fonts_changed(self, row: int) -> None:
        """A cell in a row changed its value or its format, or went away, or the row's format changed."""
        self._touched.add(row)
        self._grown_valid = False

    def column_fonts_changed(self) -> None:
        """A column's format changed, which every row without a format of its own shows."""
        self._columns_touched = True
        self._grown_valid = False

    def invalidate_growth(self) -> None:
        """Something that decides which fonts count changed: a merge."""
        self._grown_valid = False

    def _settle(self) -> None:
        """Work out how tall each row's fonts and borders make it, and the standard row the column formats make.

        Every position of a row shows a font: its cell's, else the row's
        format's, else its column's, else the Normal one. The measured tables
        give a font's row beside the Normal font and on a row to itself, and
        two fonts from different tables together are not measured. A medium
        or heavier line along a row's bottom, and a thick or double line
        along its top, draw it a pixel taller (see :func:`_edge_flags`).
        """
        if self._grown_valid:
            return
        self._grown_valid = True
        self._grown.clear()
        self._unmeasured.clear()
        self._flags.clear()
        default = self.sheet.book.stylesheet.default
        normal = default.font
        normal_entry = (_font_rows.table_of(normal.name, normal.bold, normal.italic), font_pixels(normal.size))
        unknown = None if self.normal_known() else f"the Normal font {normal.name} {normal.size:g}"
        # Each column format and the columns showing it; columns with none show the default format.
        groups = self.column_groups()
        unstyled = MAX_COLUMNS - sum(groups.values())
        if unstyled:
            groups[default] = groups.get(default, 0) + unstyled
        # The columns whose format has a line heavy enough to count, in order.
        edged_columns = sorted((column, _lines(record.style)) for column, record in self.columns.items()
                               if record.style is not None and _lines(record.style) != (0, 0))
        default_fonts = _row_fonts([_font_entry(style, False) for style in groups], normal_entry, unknown)
        default_lines = _first_lines([], edged_columns)
        default_flags = _edge_flags(default_lines, default_lines, default_lines)
        self._default_row, self._default_reason, self._default_flags = None, None, (False, False)
        if isinstance(default_fonts, str):
            self._default_reason = default_fonts
        elif _lines(default) != (0, 0):
            self._default_reason = "borders in the default format"
        elif default_fonts is not None:
            self._default_row, self._default_flags = _with_flags(default_fonts, default_flags), default_flags
        # A font in a merged cell over several rows makes none of them taller.
        tall = [area for area in self.sheet.merged_areas if area.rows > 1]
        by_row: dict[int, list[tuple[int, Cell]]] = {}
        for (row, column), cell in self.sheet.cells_.items():
            by_row.setdefault(row, []).append((column, cell))
        special = set(by_row) | {row for row, record in self.rows.items() if record.style is not None}
        fonts: dict[int, tuple[int, int] | str | None] = {}
        lines: dict[int, tuple[int, int]] = {}
        for row in special:
            cells = sorted(by_row.get(row, []), key=lambda found: found[0])
            entries = [_font_entry(cell.style or default, cell.value is not EMPTY or bool(cell.formula))
                       for column, cell in cells
                       if not (tall and any(area.contains(row, column) for area in tall))]
            record = self.rows.get(row)
            own = record.style if record is not None and record.style is not None else None
            if len(cells) < MAX_COLUMNS:
                if own is not None:
                    others = [own]
                else:
                    # A column format shows wherever one of its columns has no cell in the row.
                    covered: dict[Style, int] = {}
                    for column, _ in cells:
                        column_format = self.column_style(column) or default
                        covered[column_format] = covered.get(column_format, 0) + 1
                    others = [style for style, count in groups.items() if count > covered.get(style, 0)]
                entries.extend(_font_entry(style, False) for style in others)
            fonts[row] = _row_fonts(entries, normal_entry, unknown)
            taken = {column for column, _ in cells}
            beyond = [(0, _lines(own))] if own is not None and len(cells) < MAX_COLUMNS else \
                [(column, found) for column, found in edged_columns if column not in taken]
            lines[row] = _first_lines([_lines(cell.style or default) for _, cell in cells], beyond)
        # Rows whose borders can differ from the standard row's: those with cells or a format, their
        # neighbours, and the first and last rows, which have no neighbour on one side.
        edged = {one for row in special for one in (row - 1, row, row + 1) if 1 <= one <= MAX_ROWS}
        edged.update((1, MAX_ROWS))
        for row in edged:
            found = fonts.get(row, default_fonts)
            flags = _edge_flags(lines.get(row - 1, default_lines) if row > 1 else None,
                                lines.get(row, default_lines),
                                lines.get(row + 1, default_lines) if row < MAX_ROWS else None)
            if isinstance(found, str):
                self._unmeasured[row] = found
            elif found is None:
                # The Normal font was not measured: the file's heights stand, unless a border moves them.
                if flags != self._default_flags:
                    self._unmeasured[row] = unknown or "the Normal font"
            else:
                self._flags[row] = flags
                shape = _with_flags(found, flags)
                if shape != self._default_row:
                    self._grown[row] = shape
        self._settle_hidden(lines, default_lines)

    def _settle_hidden(self, lines: dict[int, tuple[int, int]], default_lines: tuple[int, int]) -> None:
        """Leave unmeasured the rows around a hidden row that a border draws taller.

        Excel carries such a border on past the hidden row to the next row
        shown, which was measured once and not closely enough to follow.
        """
        reason = "a medium or thick border beside a hidden row"
        if self.zero_height and (self._flags or self._default_flags != (False, False)):
            for row in set(self._flags) | set(self._grown):
                self._border_unmeasured(row, reason)
            return

        def between(upper: int) -> bool:
            """Whether a line heavy enough to count runs between a row and the next."""
            return bool(lines.get(upper, default_lines)[1] or lines.get(upper + 1, default_lines)[0])

        hidden = {row for row, record in self.rows.items() if record.hidden or record.height == 0}
        for row in hidden:
            if not ((row > 1 and between(row - 1)) or (row < MAX_ROWS and between(row))):
                continue
            first, last = row, row
            while first - 1 in hidden:
                first -= 1
            while last + 1 in hidden:
                last += 1
            for one in range(max(first - 1, 1), min(last + 1, MAX_ROWS) + 1):
                self._border_unmeasured(one, reason)

    def _border_unmeasured(self, row: int, reason: str) -> None:
        """A row whose borders the model cannot follow; a change beside it counts as a change to it,
        since the borders it shares with its neighbours decide its height."""
        self._unmeasured[row] = reason
        if self._touched & {row - 1, row, row + 1}:
            self._touched.add(row)

    def settle_growth(self) -> None:
        """After loading: a height a file recorded for a row matters only where the model cannot work it out.

        Excel works a row's height out from its fonts when it opens a
        file, so a recorded height left on a row whose fonts the model
        measures, or on a row with none, goes.
        """
        self._grown_valid = False
        self._settle()
        for row, record in list(self.rows.items()):
            if record.grown is not None and row not in self._unmeasured:
                record.grown, record.height_as_read = None, ""
                if record.empty() and not self.zero_height:
                    del self.rows[row]

    def sized_rows(self) -> set[int]:
        """Every row whose height can differ from the default: records, and rows their fonts set."""
        self._settle()
        return set(self.rows) | set(self._grown) | set(self._unmeasured)

    def written_growth(self, row: int) -> tuple[str | None, str | None]:
        """The ht a row keeping no height of its own is written with, and the descent its fonts give it.

        A row whose height rests on something not measured is written as
        its file had it while untouched, and otherwise with neither: Excel
        works both out again when it opens the file.
        """
        self._settle()
        record = self.rows.get(row)
        if row in self._unmeasured:
            if row in self._touched or self._columns_touched or record is None:
                return None, None
            return (record.height_as_read or None) if record.grown is not None else None, record.descent or None
        grown = self._grown.get(row)
        if grown is None:
            return None, None
        quarters, descent = grown
        standard = self.standard_written()[0]
        return (height_text(quarters) if quarters != standard else None), excel_number(descent / 20)

    def written_flags(self, row: int) -> tuple[bool, bool]:
        """The thickTop and thickBot a row is written with: its borders', or its file's where not measured."""
        self._settle()
        if row in self._unmeasured:
            record = self.rows.get(row)
            if record is None or row in self._touched or self._columns_touched:
                return False, False
            return record.flags_as_read
        return self._flags.get(row, self._default_flags)

    def default_flags(self) -> tuple[bool, bool]:
        """The thickTop and thickBot of a row the column formats alone decide."""
        self._settle()
        return self._default_flags

    def shaped_rows(self) -> set[int]:
        """Rows the model worked out as taller, shorter or edged otherwise than the standard row.

        Excel keeps a record for each, so each is written and each counts
        in the used block, even with no cell: a row under a double border
        does both.
        """
        self._settle()
        found = set(self._grown)
        found.update(row for row, flags in self._flags.items() if flags != self._default_flags)
        return found

    def row_hidden(self, row: int) -> bool:
        record = self.rows.get(row)
        if record is not None:
            return record.hidden
        return self.zero_height

    def row_shown_pixels(self, row: int) -> int:
        """How many pixels a row takes on screen: none when hidden."""
        if self.row_hidden(row):
            return 0
        return row_pixels(self.row_quarters(row))

    def row_is_custom(self, row: int) -> bool:
        record = self.rows.get(row)
        return self.all_rows is not None or (record is not None and record.height is not None)

    def touch_row(self, row: int) -> RowRecord:
        record = self.rows.get(row)
        if record is None:
            record = self.rows[row] = RowRecord()
        self.changed_rows.add(row)
        return record

    def settle_row(self, row: int) -> None:
        record = self.rows.get(row)
        # A row shown on a sheet whose rows are hidden by default keeps its record.
        if record is not None and record.empty() and not self.zero_height:
            del self.rows[row]
        self.changed_rows.add(row)

    def set_row_height(self, row: int, quarters: int) -> None:
        record = self.touch_row(row)
        if quarters == 0:
            record.hidden = True
        else:
            record.height, record.hidden, record.height_as_read = quarters, False, ""
        self.settle_row(row)

    def hide_row(self, row: int, hidden: bool) -> None:
        record = self.touch_row(row)
        record.hidden = hidden
        self.settle_row(row)

    def standard_row(self, row: int) -> None:
        """UseStandardHeight = True: the row loses the height it kept, and is shown."""
        record = self.touch_row(row)
        record.height, record.hidden, record.height_as_read = None, False, ""
        self.settle_row(row)

    def hide_rows_to_the_end(self, top: int) -> None:
        """Every row from one down hidden: Excel hides rows by default and lists the ones above as shown."""
        self.zero_height = True
        for row in range(1, top):
            self.rows.setdefault(row, RowRecord())
            self.changed_rows.add(row)
        for row in [row for row in self.rows if row >= top]:
            self.rows[row].hidden = True
            self.changed_rows.add(row)
        self.format_changed = True

    def show_every_row(self) -> None:
        self.zero_height = False
        for row, record in list(self.rows.items()):
            record.hidden = False
            self.settle_row(row)
        self.format_changed = True

    def hide_every_column(self) -> None:
        """Every column hidden: the standard width becomes 0, and each column keeps the width it showed, hidden.

        None keeps a width of its own afterwards: Excel drops customWidth
        from all of them.
        """
        before = self.standard_width()
        for column in range(1, MAX_COLUMNS + 1):
            record = self.touch_column(column)
            if record.custom and record.width:
                self.kept_columns.add(column)
            else:
                record.width, record.width_as_read = before, ""
            record.custom, record.hidden = False, True
        self.set_standard_width(0)

    def keep_row_height(self, row: int) -> None:
        """UseStandardHeight = False: the row keeps the height it shows now."""
        if self.row_is_custom(row):
            return
        record = self.touch_row(row)
        record.height, record.height_as_read = self.row_quarters(row), ""
        self.settle_row(row)

    def autofit_row(self, row: int) -> None:
        """AutoFit: the row loses the height it kept and is shown, as tall as its fonts make it."""
        self._settle()
        reason = self._unmeasured.get(row)
        if reason is not None:
            raise VBAUnsupportedError(f"AutoFit of row {row} rests on {reason}, which is not measured")
        record = self.touch_row(row)
        record.height, record.hidden, record.grown, record.height_as_read = None, False, None, ""
        self.settle_row(row)

    def size_all_rows(self, quarters: int) -> None:
        """Every row sized at once, which Excel keeps as the sheet's default rather than row by row."""
        self.all_rows = quarters
        for row, record in list(self.rows.items()):
            record.height, record.hidden, record.height_as_read = None, False, ""
            self.settle_row(row)
        self.format_changed = True

    # -- columns

    def column_style(self, column: int) -> Style | None:
        """The column's own format; None when it has the default one."""
        record = self.columns.get(column)
        return record.style if record is not None else None

    def column_groups(self, left: int = 1, right: int = MAX_COLUMNS) -> dict[Style, int]:
        """Each column format between two columns, and how many columns have it.

        Columns formatted together share one format object, so they are
        counted by object first and compared as formats only once each.
        """
        by_object: dict[int, tuple[Style, int]] = {}
        for column, record in self.columns.items():
            if record.style is not None and left <= column <= right:
                found = by_object.get(id(record.style))
                by_object[id(record.style)] = (record.style, found[1] + 1 if found is not None else 1)
        groups: dict[Style, int] = {}
        for style, count in by_object.values():
            groups[style] = groups.get(style, 0) + count
        return groups

    def column_pixels(self, column: int) -> int:
        """A column's width when shown, in pixels."""
        record = self.columns.get(column)
        if record is not None and record.width is not None:
            return record.width
        return self.standard_width()

    def column_hidden(self, column: int) -> bool:
        record = self.columns.get(column)
        return record is not None and record.hidden

    def column_shown_pixels(self, column: int) -> int:
        return 0 if self.column_hidden(column) else self.column_pixels(column)

    def column_is_custom(self, column: int) -> bool:
        record = self.columns.get(column)
        return record is not None and record.custom

    def touch_column(self, column: int) -> ColumnRecord:
        record = self.columns.get(column)
        if record is None:
            record = self.columns[column] = ColumnRecord()
        self.columns_changed = True
        return record

    def settle_column(self, column: int) -> None:
        record = self.columns.get(column)
        if record is None:
            return
        if record.style is None and not record.custom and not record.hidden and \
                record.width == self.standard_width():
            # The standard width stated beside a format that has gone says nothing: Excel drops it.
            record.width, record.width_as_read = None, ""
        if record.empty():
            del self.columns[column]

    def set_column_width(self, column: int, pixels: int) -> None:
        record = self.touch_column(column)
        if pixels == 0:
            # Hidden, keeping a width of its own if it had one: a standard
            # column is written with a width of 0.
            if record.custom and record.width:
                self.kept_columns.add(column)
            else:
                record.width, record.width_as_read, record.custom = 0, "", True
            record.hidden = True
        elif pixels == self.standard_width() and not record.custom:
            # The standard width is no width of its own: the column follows
            # StandardWidth from now on.  One that had a width of its own
            # keeps one, at the standard width.
            record.width, record.hidden, record.width_as_read = None, False, ""
        else:
            record.width, record.custom, record.hidden, record.width_as_read = pixels, True, False, ""
        self.settle_column(column)

    def hide_column(self, column: int, hidden: bool) -> None:
        record = self.touch_column(column)
        if hidden:
            if record.width is None or not record.custom:
                record.width, record.custom, record.width_as_read = 0, True, ""
            elif record.width:
                self.kept_columns.add(column)
            record.hidden = True
        else:
            record.hidden = False
            if record.width == 0:
                # A standard column hidden by Excel comes back at the
                # standard width, now as a width of its own.
                record.width, record.custom, record.width_as_read = self.standard_width(), True, ""
        self.settle_column(column)

    def standard_column(self, column: int) -> None:
        """UseStandardWidth = True."""
        record = self.touch_column(column)
        record.width, record.custom, record.width_as_read = None, False, ""
        self.settle_column(column)

    def set_standard_width(self, pixels: int) -> None:
        self.standard = pixels
        self.format_changed = True

    def size_all_columns(self, pixels: int) -> None:
        """Every column sized at once: the standard width changes, and no column keeps its own."""
        self.set_standard_width(pixels)
        for column in list(self.columns):
            record = self.touch_column(column)
            record.width, record.custom, record.hidden, record.width_as_read = None, False, False, ""
            self.settle_column(column)

    # -- where things are

    def default_row_pixels(self) -> int:
        """How many pixels a row with no record of its own takes."""
        return 0 if self.zero_height else row_pixels(self.default_quarters())

    def rows_pixels(self, top: int, bottom: int) -> int:
        """Pixels from the top of one row to the bottom of another."""
        if bottom < top:
            return 0
        default = self.default_row_pixels()
        total = (bottom - top + 1) * default
        for row in self.sized_rows():
            if top <= row <= bottom:
                total += self.row_shown_pixels(row) - default
        return total

    def columns_pixels(self, left: int, right: int) -> int:
        if right < left:
            return 0
        default = self.standard_width()
        total = (right - left + 1) * default
        for column in self.columns:
            if left <= column <= right:
                total += self.column_shown_pixels(column) - default
        return total

    # -- rows and columns moving

    def shift_rows(self, start: int, count: int, delete: bool) -> None:
        moved: dict[int, RowRecord] = {}
        for row, record in self.rows.items():
            if delete:
                if start <= row < start + count:
                    continue
                target = row - count if row >= start + count else row
            else:
                target = row + count if row >= start else row
            if target <= MAX_ROWS:
                moved[target] = record
        if not delete and start > 1:
            # A new row takes the height and the format of the row above it, but is shown.
            above = self.rows.get(start - 1)
            if above is not None and (above.height is not None or above.style is not None):
                for row in range(start, min(start + count, MAX_ROWS + 1)):
                    moved[row] = RowRecord(height=above.height, style=above.style, xf=above.xf)
        self.rows = moved
        self.moved = True
        touched: set[int] = set()
        for row in self._touched:
            if delete and start <= row < start + count:
                continue
            target = (row - count if row >= start + count else row) if delete else (row + count if row >= start else row)
            if target <= MAX_ROWS:
                touched.add(target)
        self._touched = touched
        self._grown_valid = False

    def shift_columns(self, start: int, count: int, delete: bool) -> None:
        moved: dict[int, ColumnRecord] = {}
        for column, record in self.columns.items():
            if delete:
                if start <= column < start + count:
                    continue
                target = column - count if column >= start + count else column
            else:
                target = column + count if column >= start else column
            if target <= MAX_COLUMNS:
                moved[target] = record
        if not delete and start > 1:
            # A new column takes the width and the format of the one to its left.
            left = self.columns.get(start - 1)
            width = left.width if left is not None and left.custom and left.width else None
            if left is not None and (width is not None or left.style is not None):
                for column in range(start, min(start + count, MAX_COLUMNS + 1)):
                    moved[column] = ColumnRecord(width=width, custom=width is not None, style=left.style, xf=left.xf)
        last = self.columns.get(MAX_COLUMNS)
        if delete and last is not None and last.style is not None:
            # The last column's format is the sheet's own: the columns a
            # deletion opens up at the end take it.
            for column in range(MAX_COLUMNS - count + 1, MAX_COLUMNS + 1):
                moved.setdefault(column, ColumnRecord(style=last.style, xf=last.xf))
        self.columns = moved
        self.columns_changed = True
        self._grown_valid = False
        kept: set[int] = set()
        for column in self.kept_columns:
            if delete and start <= column < start + count:
                continue
            target = (column - count if column >= start + count else column) if delete else \
                (column + count if column >= start else column)
            if target <= MAX_COLUMNS:
                kept.add(target)
        self.kept_columns = kept

    def copy_rows_from(self, source: SheetDimensions, rows: list[tuple[int, int]]) -> None:
        """Whole rows copied: each destination row takes its source row's height and format."""
        same_book = source.sheet.book is self.sheet.book
        for source_row, row in rows:
            found = source.rows.get(source_row)
            record = self.touch_row(row)
            record.height, record.height_as_read = (found.height, found.height_as_read) if found else (None, "")
            record.style = found.style if found is not None else None
            record.xf = found.xf if found is not None and same_book else -1
            self.fonts_changed(row)
            self.settle_row(row)

    def copy_columns_from(self, source: SheetDimensions, columns: list[tuple[int, int]]) -> None:
        """Whole columns copied: each destination column takes its source column's width and format."""
        same_book = source.sheet.book is self.sheet.book
        for source_column, column in columns:
            found = source.columns.get(source_column)
            record = self.touch_column(column)
            if found is not None and found.custom and found.width:
                record.width, record.custom, record.width_as_read = found.width, True, found.width_as_read
            else:
                record.width, record.custom, record.width_as_read = None, False, ""
            record.style = found.style if found is not None else None
            record.xf = found.xf if found is not None and same_book else -1
            self.settle_column(column)
        self.column_fonts_changed()

    def copied(self, sheet: Worksheet) -> SheetDimensions:
        """The same sizes, for a copy of the sheet."""
        other = SheetDimensions(sheet)
        other.rows = copy.deepcopy(self.rows)
        other.columns = copy.deepcopy(self.columns)
        other.format = dict(self.format)
        other.all_rows, other.zero_height, other.standard = self.all_rows, self.zero_height, self.standard
        other.kept_columns = set(self.kept_columns)
        other._touched = set(self._touched)
        other._columns_touched = self._columns_touched
        other.changed_rows = set(other.rows)
        other.columns_changed = other.format_changed = other.moved = True
        return other

    def record_rows(self) -> Iterable[int]:
        return self.rows.keys()

    def used_columns(self) -> list[int]:
        """Columns that count toward the sheet's used block without a cell.

        A hidden column keeping a width of its own counts; so does one
        that was hidden that way at any point while the workbook is open,
        since Excel's block does not give it back; and so does a column
        with a format apart from the sheet's.
        """
        found = set(self.kept_columns)
        found.update(column for column, record in self.columns.items()
                     if record.hidden and record.custom and record.width)
        found.update(self.formatted_columns())
        return sorted(found)

    def formatted_columns(self) -> list[int]:
        """Columns whose format is not the sheet's own, which the last column's is.

        Formatting the whole sheet formats every column, and Excel counts
        none of them in the used block; a column formatted apart from the
        rest counts, and so does one left in the default format beside them.
        Excel reads a file's columns the same way: whatever format column XFD
        has is the sheet's.
        """
        sheet_format = self.column_style(MAX_COLUMNS)
        if sheet_format is None:
            return [column for column, record in self.columns.items() if record.style is not None]
        return [column for column in range(1, MAX_COLUMNS + 1) if self.column_style(column) != sheet_format]

    def columns_written(self) -> bool:
        """Whether the sheet's cols element has to be written again."""
        return self.columns_changed

    def saved(self) -> None:
        """The sheet part now says what the model does: nothing is pending."""
        self.changed_rows.clear()
        self.columns_changed = self.format_changed = self.moved = False


#: Lines that make the row above them a pixel taller, and lines that also make the row below one.
_MEDIUM_LINES = frozenset({"medium", "mediumDashed", "mediumDashDot", "mediumDashDotDot", "slantDashDot"})
_THICK_LINES = frozenset({"thick", "double"})


def _weight(side: Side) -> int:
    """1 for a medium line, 2 for a thick or double one, 0 for anything lighter."""
    return 2 if side.style in _THICK_LINES else 1 if side.style in _MEDIUM_LINES else 0


def _lines(style: Style) -> tuple[int, int]:
    """The weights of a format's top and bottom lines."""
    return _weight(style.border.top), _weight(style.border.bottom)


def _first_lines(cells: list[tuple[int, int]], beyond: list[tuple[int, tuple[int, int]]]) -> tuple[int, int]:
    """The weight that counts along a row's top and along its bottom.

    Excel takes the first cell along the row, by column, with a line heavy
    enough to count; failing that, the row's own format, or the first
    column whose format has one. So a medium line in B2 and a double one in
    C2 make row 3 no taller, while a double one in B2 and a medium one in
    C2 do, whichever was set first.
    """
    top = next((found for found, _ in cells if found), 0)
    bottom = next((found for _, found in cells if found), 0)
    if not top:
        top = next((found for _, (found, _) in beyond if found), 0)
    if not bottom:
        bottom = next((found for _, (_, found) in beyond if found), 0)
    return top, bottom


def _edge_flags(above: tuple[int, int] | None, own: tuple[int, int],
                below: tuple[int, int] | None) -> tuple[bool, bool]:
    """A row's thickTop and thickBot, from the lines that count along its top and bottom and its neighbours'.

    A medium or heavier line between a row and the next draws the upper
    one a pixel taller (thickBot); a thick or double one draws the lower
    one a pixel taller too (thickTop). A line is between the two rows
    whichever of them keeps it, as MS-XLS says of the ROW record's flags;
    the first row has nothing above it, and the last nothing below.
    """
    thick_top = own[0] == 2 or (above is not None and above[1] == 2)
    thick_bottom = own[1] > 0 or (below is not None and below[0] > 0)
    return thick_top, thick_bottom


def _with_flags(shape: tuple[int, int], flags: tuple[bool, bool]) -> tuple[int, int]:
    """A row's height in quarter pixels and descent in twips once its borders' pixels are added.

    Each flag is a pixel, and thickBot deepens the descent by a twip.
    """
    top, bottom = flags
    return shape[0] + 4 * (top + bottom), shape[1] + bottom


def _font_entry(style: Style, text: bool) -> tuple[int, int] | str | None:
    """What a cell showing a format adds to its row's height.

    The font's measured table and pixel size; the reason when the model
    cannot tell; or None when it adds nothing, as turned text in an empty
    cell does.
    """
    font, alignment = style.font, style.alignment
    wraps = alignment.wrap or alignment.horizontal in _WRAPPING or alignment.vertical in _WRAPPING
    if alignment.text_rotation:
        # Turned text in an empty cell takes no room; with text it takes its length.
        return "turned text" if text else None
    if wraps and text:
        return "wrapped text"
    if font.vert_align in ("superscript", "subscript"):
        return f"{font.vert_align} text"
    table = _font_rows.table_of(font.name, font.bold, font.italic)
    pixels = font_pixels(font.size)
    if table is None or not 1 <= pixels <= _font_rows.MOST_PIXELS:
        return f"the font {font.name}"
    return table, pixels


def _row_fonts(entries: Iterable[tuple[int, int] | str | None], normal: tuple[int | None, int],
               unknown: str | None) -> tuple[int, int] | str | None:
    """How tall a row showing these fonts is, in quarter pixels, and its descent in twips.

    Beside the Normal font a font makes the row its measured height beside
    it; with no position left in the Normal font, its height alone, which a
    small font leaves shorter than the standard row. Fonts from different
    tables together add the tallest ascent to the deepest descent, which
    the heights do not tell apart. The reason comes back when the model
    cannot tell, and None for a row in the Normal font alone when that
    font was not measured, which is as tall as the file says.
    """
    fonts: set[tuple[int, int]] = set()
    for entry in entries:
        if isinstance(entry, str):
            return entry
        if entry is not None:
            fonts.add(entry)
    normal_table, normal_pixels = normal
    if unknown is not None or normal_table is None:
        return None if fonts <= {normal} else (unknown or "the Normal font")
    if not fonts:
        return "turned text"
    alone = normal not in fonts
    if not alone:
        # Beside the Normal font, the same font no bigger adds nothing.
        fonts = {(table, pixels) for table, pixels in fonts
                 if not (table == normal_table and pixels <= normal_pixels)} or {(normal_table, normal_pixels)}
    tables = {table for table, _ in fonts}
    if len(tables) > 1:
        return "a mix of fonts"
    (table,) = tables
    rows, descent = _font_rows.row_of(table, max(pixels for _, pixels in fonts), alone=alone)
    return rows * 4, descent


# --- what a macro reads and sets -----------------------------------------------------------------


def _size(value: object) -> float:
    """A height or width a macro gives: True counts as 1, and text must read as a number."""
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    try:
        return float(to_number(value))
    except (VBARuntimeError, TypeError, ValueError, OverflowError):
        raise error(1004, "a height or width must be a number") from None


def _flag(value: object) -> bool:
    """A Hidden or UseStandard setting: a number is True unless zero, and text must say True or False."""
    if value is EMPTY:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        folded = value.strip().lower()
        if folded in ("true", "false"):
            return folded == "true"
        raise error(1004, "that setting takes True or False")
    try:
        return float(to_number(value)) != 0
    except (VBARuntimeError, TypeError, ValueError):
        raise error(1004, "that setting takes True or False") from None


def _dims(target: Range) -> SheetDimensions:
    return target.sheet.dims


def _all_rows(area: Area) -> bool:
    return area.top == 1 and area.bottom == MAX_ROWS


def _all_columns(area: Area) -> bool:
    return area.left == 1 and area.right == MAX_COLUMNS


def _used_area(sheet: Worksheet) -> tuple[int, int, int, int] | None:
    """The block a multi-row or multi-column read compares inside.

    Its columns are those of the sheet's cells; its rows reach every row
    that has a record of its own, which is Excel's UsedRange. A sheet
    with no cells compares nothing, whatever sizes its rows have.
    """
    cells = [(row, column) for (row, column), cell in sheet.cells_.items() if sheet.holds(row, column, cell)]
    for area in sheet.merged_areas:
        cells.extend(((area.top, area.left), (area.bottom, area.right)))
    if not cells:
        return None
    rows = [row for row, _ in cells]
    rows.extend(sheet.dims.record_rows())
    rows.extend(sheet.dims.shaped_rows())
    columns = [column for _, column in cells]
    return min(rows), min(columns), max(rows), max(columns)


def _compared(target: Range, along_rows: bool, value: Callable[[int], object], records: Iterable[int]) -> object:
    """A property read across several rows (or columns): the first one's, or Null when they differ.

    Excel compares only the part of the range inside the used block, and
    walks it the way a loop with a slip in it would: the first line from
    the block's own edge, every later line from the range's edge. So on a
    sheet with no cells nothing is compared and the first row answers; a
    range whose part in the block is a single line compares that line
    alone; and a taller part pulls in the range's first row or column
    too. When everything compared agrees, the answer is still the range's
    first row or column, compared or not. Fitted to 459 live reads.
    """
    area = target.first
    first = value(area.top if along_rows else area.left)
    used = _used_area(target.sheet)
    if used is None:
        return first
    top, left = max(area.top, used[0]), max(area.left, used[1])
    bottom, right = min(area.bottom, used[2]), min(area.right, used[3])
    if top > bottom or left > right:
        return first
    if along_rows:
        low, high = (area.top if right > left else top), bottom
    else:
        low, high = (area.left if bottom > top else left), right
    inside = {index for index in records if low <= index <= high}
    seen = {value(index) for index in inside}
    if len(inside) < high - low + 1:
        # Every line without a record answers the same, so one stands for all.
        index = low
        while index in inside:
            index += 1
        seen.add(value(index))
    return first if len(seen) <= 1 else NULL


def _rows_of(target: Range) -> Iterable[int]:
    for area in target.areas:
        yield from range(area.top, area.bottom + 1)


def _columns_of(target: Range) -> Iterable[int]:
    for area in target.areas:
        yield from range(area.left, area.right + 1)


def read_row_height(target: Range) -> object:
    dims = _dims(target)

    def height(row: int) -> object:
        return 0.0 if dims.row_hidden(row) else row_points_read(dims.row_quarters(row))

    return _compared(target, True, height, dims.sized_rows())


def write_row_height(target: Range, value: object) -> None:
    points = _size(value)
    if points < 0 or _c_round(points * 20) > MOST_TWIPS:
        raise error(1004, "a row can be from 0 to 409.5 points tall")
    quarters = quarters_for(points)
    dims = _dims(target)
    if any(_all_rows(area) for area in target.areas):
        if not quarters:
            raise VBAUnsupportedError("hiding every row of a sheet is not implemented")
        dims.size_all_rows(quarters)
    else:
        for row in _rows_of(target):
            dims.set_row_height(row, quarters)
    target.sheet.touched()


def read_column_width(target: Range) -> object:
    dims = _dims(target)

    def width(column: int) -> object:
        return 0.0 if dims.column_hidden(column) else characters_read(dims.column_pixels(column))

    return _compared(target, False, width, dims.columns.keys())


def write_column_width(target: Range, value: object) -> None:
    characters = _size(value)
    if characters < 0 or characters > MOST_CHARACTERS:
        raise error(1004, "a column can be from 0 to 255 characters wide")
    pixels = pixels_for(characters)
    dims = _dims(target)
    if any(_all_columns(area) for area in target.areas):
        dims.size_all_columns(pixels)
    else:
        for column in _columns_of(target):
            dims.set_column_width(column, pixels)
    target.sheet.touched()


def _whole(target: Range) -> tuple[Area, bool]:
    """The one area Hidden and AutoFit work on, and whether it is whole rows (else whole columns)."""
    if len(target.areas) != 1:
        raise error(1004, "that needs whole rows or whole columns")
    area = target.first
    rows, columns = area.left == 1 and area.right == MAX_COLUMNS, area.top == 1 and area.bottom == MAX_ROWS
    if rows and columns:
        # The whole sheet: rows or columns only when it was made as them.
        rows, columns = target.whole == "rows", target.whole == "columns"
    if rows == columns:
        raise error(1004, "that needs whole rows or whole columns")
    return area, rows


def read_hidden(target: Range) -> object:
    area, rows = _whole(target)
    dims = _dims(target)
    if rows:
        return dims.row_hidden(area.top) or row_pixels(dims.row_quarters(area.top)) == 0
    return dims.column_hidden(area.left) or dims.column_pixels(area.left) == 0


def write_hidden(target: Range, value: object) -> None:
    """Hidden on whole rows or columns.

    Hiding rows down to the last one makes hidden the sheet's default
    (zeroHeight), with the rows above listed as shown: measured for every
    row and for rows 3 down. Hiding every column makes the standard width
    0, and showing every column again then does nothing, as in Excel.
    """
    area, rows = _whole(target)
    hidden = _flag(value)
    dims = _dims(target)
    if rows:
        if area.bottom == MAX_ROWS and area.top - 1 < MAX_ROWS - area.top + 1:
            if hidden:
                dims.hide_rows_to_the_end(area.top)
            elif _all_rows(area):
                dims.show_every_row()
            else:
                for row in range(area.top, area.bottom + 1):
                    dims.hide_row(row, False)
        else:
            for row in range(area.top, area.bottom + 1):
                dims.hide_row(row, hidden)
    elif _all_columns(area):
        if hidden:
            dims.hide_every_column()
        elif dims.standard_width():
            for column in range(area.left, area.right + 1):
                dims.hide_column(column, False)
    else:
        for column in range(area.left, area.right + 1):
            dims.hide_column(column, hidden)
    target.sheet.touched()


def read_use_standard_height(target: Range) -> object:
    dims = _dims(target)
    return _compared(target, True, lambda row: not dims.row_is_custom(row), dims.sized_rows())


def write_use_standard_height(target: Range, value: object) -> None:
    standard = _flag(value)
    dims = _dims(target)
    for row in _rows_of(target):
        if standard:
            dims.standard_row(row)
        else:
            dims.keep_row_height(row)
    target.sheet.touched()


def read_use_standard_width(target: Range) -> object:
    dims = _dims(target)
    return _compared(target, False, lambda column: not dims.column_is_custom(column), dims.columns.keys())


def write_use_standard_width(target: Range, value: object) -> None:
    # Setting False does nothing: a column only gets a width of its own by being given one.
    if not _flag(value):
        return
    dims = _dims(target)
    for column in _columns_of(target):
        dims.standard_column(column)
    target.sheet.touched()


def read_height(target: Range) -> object:
    area = target.first
    return _dims(target).rows_pixels(area.top, area.bottom) * PIXEL


def read_width(target: Range) -> object:
    area = target.first
    return _dims(target).columns_pixels(area.left, area.right) * PIXEL


def read_top(target: Range) -> object:
    return _dims(target).rows_pixels(1, target.first.top - 1) * PIXEL


def read_left(target: Range) -> object:
    return _dims(target).columns_pixels(1, target.first.left - 1) * PIXEL


def autofit(target: Range) -> None:
    """AutoFit on whole rows or columns: a row loses the height it kept; an empty column keeps its width."""
    area, rows = _whole(target)
    dims = _dims(target)
    if rows:
        for row in range(area.top, area.bottom + 1):
            dims.autofit_row(row)
        target.sheet.touched()
        return
    for (_, column), cell in target.sheet.cells_.items():
        if area.left <= column <= area.right and (cell.value is not EMPTY or cell.formula):
            raise VBAUnsupportedError("AutoFit of a column holding values is not implemented: it measures their text")


def read_standard_width(sheet: Worksheet) -> object:
    return characters_read(sheet.dims.standard_width())


def write_standard_width(sheet: Worksheet, value: object) -> None:
    characters = _size(value)
    if characters < 0 or characters > MOST_CHARACTERS:
        raise error(1004, "the standard width can be from 0 to 255 characters")
    sheet.dims.set_standard_width(pixels_for(characters))
    sheet.touched()


def read_standard_height(sheet: Worksheet) -> object:
    """The standard row, or 0 on a sheet whose rows are hidden by default."""
    dims = sheet.dims
    return 0.0 if dims.zero_height else row_points_read(dims.default_quarters())


# --- writing the sheet back ------------------------------------------------------------------------


def block_spans(sheet: Worksheet) -> dict[int, str]:
    """The spans Excel writes on the rows of each 16-row block: the block's first and last column with a cell."""
    reach: dict[int, list[int]] = {}
    for (row, column), cell in sheet.cells_.items():
        if not sheet.holds(row, column, cell):
            continue
        block = (row - 1) // 16
        found = reach.get(block)
        if found is None:
            reach[block] = [column, column]
        else:
            found[0], found[1] = min(found[0], column), max(found[1], column)
    return {block: f"{low}:{high}" for block, (low, high) in reach.items()}


def row_start_tag(sheet: Worksheet, row: int, original: dict[str, str] | None, spans: str | None,
                  descent: str | None, has_cells: bool) -> str | None:
    """A row's start tag as Excel writes it, or None when the row has nothing left to say.

    ``original`` is the tag's attributes as the file had them at this row,
    which carry over while no row has moved; ``descent`` is what a row as
    tall as the sheet's standard row gets, or None when the sheet does not
    declare the namespace it lives in. A row with a format of its own
    names it with customFormat, leaving s out for the default format. Once
    every row has been sized at once, a row with a record states that
    height as its own.
    """
    dims = sheet.dims
    record = dims.rows.get(row)
    attrs: dict[str, str] = {"r": str(row)}
    if spans is not None:
        attrs["spans"] = spans
    kept = original if original is not None and not dims.moved else {}
    attrs.update({key: value for key, value in kept.items() if key not in _ROW_OWN})
    grown_height, grown_descent = dims.written_growth(row)
    if record is not None:
        attrs.update(record.extra)
        if record.style is not None:
            index = sheet.book.stylesheet.index_as_read(record.style, record.xf)
            if index:
                attrs["s"] = str(index)
            attrs["customFormat"] = "1"
        height = record.height if record.height is not None else dims.all_rows
        if height is not None:
            attrs["ht"] = record.height_as_read or height_text(height)
        if record.hidden or record.height == 0:
            attrs["hidden"] = "1"
        if height is not None:
            attrs["customHeight"] = "1"
    if "ht" not in attrs and grown_height is not None and dims.all_rows is None:
        attrs["ht"] = grown_height
    thick_top, thick_bottom = dims.written_flags(row)
    if thick_top:
        attrs["thickTop"] = "1"
    if thick_bottom:
        attrs["thickBot"] = "1"
    listed = record is not None and dims.zero_height
    if len(attrs) == 1 + (spans is not None) and not has_cells and not listed:
        return None
    if descent is not None:
        # The descent comes from the row's fonts, even when it keeps a height of its own.
        attrs[_DESCENT] = grown_descent or descent
    names = _ordered(attrs, _ROW_ORDER)
    body = " ".join(f'{key}="{_escape(attrs[key])}"' for key in names)
    return f"<row {body}>"


def columns_element(dims: SheetDimensions) -> str:
    """The sheet's cols element, with neighbouring columns that read the same written as one."""
    spans: list[tuple[int, int, dict[str, str]]] = []
    stylesheet, standard = dims.sheet.book.stylesheet, width_text(dims.standard_width())
    # Columns formatted together share a format object: look its xf up once.
    found: dict[tuple[int, int], tuple[Style, int]] = {}

    def index_of(style: Style, xf: int) -> int:
        known = found.get((id(style), xf))
        if known is None or known[0] is not style:
            known = found[(id(style), xf)] = (style, stylesheet.index_as_read(style, xf))
        return known[1]

    for column in sorted(dims.columns):
        attrs = dims.columns[column].attributes(standard, index_of)
        if spans and spans[-1][1] == column - 1 and spans[-1][2] == attrs:
            spans[-1] = (spans[-1][0], column, attrs)
        else:
            spans.append((column, column, attrs))
    if not spans:
        return ""
    body = "".join(
        f'<col min="{first}" max="{last}"' + "".join(f' {key}="{_escape(value)}"' for key, value in attrs.items())
        + "/>"
        for first, last, attrs in spans)
    return f"<cols>{body}</cols>"


def format_element(sheet: Worksheet, has_descent: bool) -> str | None:
    """The sheet's sheetFormatPr as Excel on this display writes it, or None to keep the file's."""
    dims = sheet.dims
    known = dims.normal_known()
    if not (dims.format_changed or known):
        return None
    attrs = dict(dims.format)
    if dims.standard is not None and (dims.format_changed or "defaultColWidth" not in attrs):
        attrs["defaultColWidth"] = width_text(dims.standard)
    if dims.all_rows is not None:
        attrs["defaultRowHeight"] = height_text(dims.all_rows)
        attrs["customHeight"] = "1"
    elif known:
        attrs["defaultRowHeight"] = height_text(dims.standard_written()[0])
        attrs.pop("customHeight", None)
    if known:
        # Borders every row shows make the standard row taller, and the sheet says which.
        for name, on in zip(("thickTop", "thickBottom"), dims.default_flags(), strict=True):
            if on:
                attrs[name] = "1"
            else:
                attrs.pop(name, None)
    elif "defaultRowHeight" not in attrs:
        attrs["defaultRowHeight"] = height_text(dims.standard_written()[0])
    if dims.zero_height:
        attrs["zeroHeight"] = "1"
    else:
        attrs.pop("zeroHeight", None)
    if has_descent and known:
        attrs[_DESCENT] = dims.default_descent()
    names = _ordered(attrs, _FORMAT_ORDER)
    if _DESCENT in names:
        names.remove(_DESCENT)
        names.append(_DESCENT)
    # What was written is what the part says from now on.
    dims.format = {key: attrs[key] for key in names}
    return "<sheetFormatPr " + " ".join(f'{key}="{_escape(attrs[key])}"' for key in names) + "/>"
