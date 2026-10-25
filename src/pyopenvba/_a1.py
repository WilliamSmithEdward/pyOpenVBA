"""A1 references: the spelling Excel uses for a cell, a block, or a list of them.

Pure string and number handling, with no host and no errors of its own
beyond :class:`ValueError`, so both the Power Query writer and the
in-memory workbook can use it and each raise the error its own caller
expects.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Final

#: The grid Excel has had since 2007.
MAX_ROWS: Final = 1048576
MAX_COLUMNS: Final = 16384

_CELL = re.compile(r"^([A-Za-z]{1,3})([0-9]{1,7})$")
_COLUMN = re.compile(r"^([A-Za-z]{1,3})$")
_ROW = re.compile(r"^([0-9]{1,7})$")
#: Past ASCII, the characters Excel quotes a sheet's name for wherever they stand in it, and those it quotes a name
#: for only when the name starts with them; every other character past ASCII reads as a letter. pyOfficeEditor
#: measured them a character at a time over eight blocks of Unicode (its commit a74926e,
#: scripts/measure_sheet_names.py), and a sample agrees here (tests/fixtures/sheet_quoting.json).
_QUOTED_ANYWHERE: Final = (
    (0x00A0, 0x00A0), (0x00A2, 0x00A3), (0x00A5, 0x00A6), (0x00A9, 0x00A9), (0x00AB, 0x00AC),
    (0x00AE, 0x00AE), (0x00BB, 0x00BB), (0x2011, 0x2012), (0x2017, 0x2017), (0x201A, 0x201B),
    (0x201E, 0x201F), (0x2022, 0x2024), (0x2031, 0x2031), (0x2034, 0x2034), (0x2036, 0x203A),
    (0x203C, 0x2043), (0x2045, 0x2051), (0x2053, 0x205E), (0x2065, 0x2069), (0x20B6, 0x20C0),
    (0x3018, 0x301C), (0x3030, 0x3030), (0x303D, 0x303D), (0xFF5F, 0xFF60),
)
_QUOTED_FIRST: Final = (
    (0x2000, 0x200F), (0x2028, 0x202F), (0x2044, 0x2044), (0x2052, 0x2052), (0x205F, 0x2064),
    (0x206A, 0x206F), (0x20A0, 0x20B5), (0x2100, 0x2101), (0x2104, 0x2104), (0x2106, 0x2106),
    (0x2108, 0x2108), (0x2114, 0x2114), (0x2117, 0x2118), (0x211E, 0x2120), (0x2123, 0x2123),
    (0x2125, 0x2125), (0x2127, 0x2127), (0x2129, 0x2129), (0x212E, 0x212E), (0x213A, 0x213B),
    (0x2140, 0x2144), (0x214A, 0x214D), (0x214F, 0x214F), (0x219A, 0x219F), (0x2201, 0x2201),
    (0x2204, 0x2206), (0x2209, 0x220A), (0x220C, 0x220E), (0x2210, 0x2210), (0x2212, 0x2214),
    (0x2216, 0x2219), (0x221B, 0x221C), (0x3004, 0x3004), (0x3020, 0x3020), (0x302A, 0x302F),
    (0x3036, 0x3037), (0x303E, 0x303F),
)
#: A name that may read as a cell: one to three letters and a number.
_CELL_LIKE = re.compile(r"([A-Za-z]{1,3})([0-9]+)")
#: A row, a column or a cell in R1C1 at the start of a name: R2, C3, RC1, R1C1.
_R1C1_START = re.compile(r"(RC|R|C)([0-9]+)", re.IGNORECASE)
#: Another workbook's sheet as a file spells it, by its link's number, [1]Sheet1 or [1]Sheet1:Other, with no quotes.
_LINKED_NAME = re.compile(r"\[[0-9]+\](?:[A-Za-z_][A-Za-z0-9_.]*(?::[A-Za-z_][A-Za-z0-9_.]*)?)?")


def column_letter(index: int) -> str:
    """``1`` is ``A``, ``27`` is ``AA``."""
    if index < 1:
        raise ValueError("a column number starts at one")
    out = ""
    while index:
        index, rest = divmod(index - 1, 26)
        out = chr(ord("A") + rest) + out
    return out


def column_number(letters: str) -> int:
    """``A`` is ``1``, ``AA`` is ``27``."""
    if not letters or not letters.isalpha():
        raise ValueError(f"{letters!r} is not a column")
    out = 0
    for char in letters.upper():
        out = out * 26 + (ord(char) - ord("A") + 1)
    return out


def quote_sheet(name: str) -> str:
    """A sheet's name as a formula reference spells it.

    A plain name a formula would not read as something else is written
    bare. Anything else is wrapped in apostrophes, and an apostrophe
    inside the name is doubled, because a single one would close the
    quoting early and leave a reference to something else.
    """
    if _LINKED_NAME.fullmatch(name) or (_plain(name) and not _read_otherwise(name)):
        return name
    return "'" + name.replace("'", "''") + "'"


def quote_workbook(book: str, sheet: str = "") -> str:
    """A workbook's sheet as a formula names it from outside the workbook, ``[Book1]Data``, or the workbook alone
    for one of its names, ``Book1``: in apostrophes where a character of the workbook's name or a sheet's needs
    them, though a sheet named like a cell stands bare after the workbook's name. ``sheet`` may be a pair of
    sheets, ``Jan:Mar`` (tests/fixtures/convert_formula.json)."""
    text = f"[{book}]{sheet}" if sheet else book
    sheets = sheet.split(":") if sheet else []
    if _plain(book) and all(_plain(part) for part in sheets):
        return text
    return "'" + text.replace("'", "''") + "'"


def _plain(name: str) -> bool:
    """Whether every character of a name may stand bare where it is: in ASCII a letter or an underscore first, then
    letters, digits, underscores and full stops; past ASCII any character the two tables above leave out."""
    if not name:
        return False
    first = ord(name[0])
    if first < 128:
        if not (name[0].isalpha() or name[0] == "_"):
            return False
    elif _within(first, _QUOTED_ANYWHERE) or _within(first, _QUOTED_FIRST):
        return False
    return all((char.isalnum() or char in "_.") if char.isascii() else not _within(ord(char), _QUOTED_ANYWHERE)
               for char in name[1:])


def _within(point: int, spans: tuple[tuple[int, int], ...]) -> bool:
    return any(low <= point <= high for low, high in spans)


def _read_otherwise(name: str) -> bool:
    """Whether a formula would read a plain name as something other than a sheet, so that Excel writes it in
    apostrophes: a logical; a cell, Q4 or A01, though no name starting LOG10, the name of a function, LOG100 among
    them; R, C or RC; or R, C or RC with a number from 1 to 1048576, alone or going on with anything but a full
    stop, R2, C3, RC1, R1x and C1A, though R1.x is plain (tests/fixtures/sheet_quoting.json, and pyOfficeEditor's
    1,246 names)."""
    if name.upper() in ("TRUE", "FALSE", "R", "C", "RC"):
        return True
    cell = _CELL_LIKE.fullmatch(name)
    if cell is not None and 1 <= int(cell.group(2)) <= MAX_ROWS and column_number(cell.group(1)) <= MAX_COLUMNS:
        return not name.upper().startswith("LOG10")
    start = _R1C1_START.match(name)
    return start is not None and not name[start.end() :].startswith(".") and 1 <= int(start.group(2)) <= MAX_ROWS


@dataclass(frozen=True, slots=True)
class Area:
    """One rectangle of cells, and the sheet it was written against."""

    top: int
    left: int
    bottom: int
    right: int
    sheet: str = ""

    @property
    def rows(self) -> int:
        return self.bottom - self.top + 1

    @property
    def columns(self) -> int:
        return self.right - self.left + 1

    @property
    def whole_columns(self) -> bool:
        return self.top == 1 and self.bottom == MAX_ROWS

    @property
    def whole_rows(self) -> bool:
        return self.left == 1 and self.right == MAX_COLUMNS

    def address(
        self,
        *,
        absolute: bool | None = None,
        rows_fixed: bool = True,
        columns_fixed: bool = True,
        with_sheet: bool = False,
    ) -> str:
        """The reference as Excel's Address property spells it.

        The two dollar signs are independent: Address(False, True) gives
        ``$A1``, with the column fixed and the row free.
        """
        if absolute is not None:
            rows_fixed = columns_fixed = absolute
        row_mark = "$" if rows_fixed else ""
        column_mark = "$" if columns_fixed else ""
        if self.whole_columns and not self.whole_rows:
            body = f"{column_mark}{column_letter(self.left)}:{column_mark}{column_letter(self.right)}"
        elif self.whole_rows:
            # The whole sheet is spelt as its rows: $1:$1048576.
            body = f"{row_mark}{self.top}:{row_mark}{self.bottom}"
        else:
            first = f"{column_mark}{column_letter(self.left)}{row_mark}{self.top}"
            last = f"{column_mark}{column_letter(self.right)}{row_mark}{self.bottom}"
            body = first if (self.top, self.left) == (self.bottom, self.right) else f"{first}:{last}"
        if with_sheet and self.sheet:
            return f"{quote_sheet(self.sheet)}!{body}"
        return body

    def contains(self, row: int, column: int) -> bool:
        return self.top <= row <= self.bottom and self.left <= column <= self.right


def split_sheet(text: str) -> tuple[str, str]:
    """``'My Sheet'!A1`` split into the sheet's name and the reference."""
    body = text.strip()
    if "!" not in body:
        return "", body
    name, _, rest = body.rpartition("!")
    name = name.strip()
    if name.startswith("'") and name.endswith("'") and len(name) >= 2:
        name = name[1:-1].replace("''", "'")
    if name.startswith("[") and "]" in name:
        name = name.partition("]")[2]
    return name, rest.strip()


def parse_area(text: str, *, sheet: str = "") -> Area:
    """One rectangle from ``A1``, ``A1:B2``, ``A:C`` or ``3:5``."""
    named, body = split_sheet(text)
    body = body.replace("$", "").strip()
    if not body:
        raise ValueError("an empty reference")
    first, _, last = body.partition(":")
    start = _corner(first)
    if not last:
        top, left = start
        if top is None or left is None:
            # A column or a row on its own is not a reference: Excel
            # refuses Range("ZZ") and wants Range("ZZ:ZZ").
            raise ValueError(f"{text!r} names a whole column or row, which needs a colon")
        return Area(top, left, top, left, named or sheet)
    stop = _corner(last)
    tops = [value for value in (start[0], stop[0]) if value is not None]
    lefts = [value for value in (start[1], stop[1]) if value is not None]
    if start[0] is None and stop[0] is None:
        top, bottom = 1, MAX_ROWS
    else:
        top, bottom = min(tops), max(tops)
    if start[1] is None and stop[1] is None:
        left, right = 1, MAX_COLUMNS
    else:
        left, right = min(lefts), max(lefts)
    return Area(top, left, bottom, right, named or sheet)


def _corner(text: str) -> tuple[int | None, int | None]:
    """A corner as (row, column); either may be missing for a whole row or column."""
    body = text.strip()
    cell = _CELL.match(body)
    if cell:
        row = int(cell.group(2))
        column = column_number(cell.group(1))
        if not 1 <= row <= MAX_ROWS or not 1 <= column <= MAX_COLUMNS:
            raise ValueError(f"{text!r} is outside the grid")
        return row, column
    if _COLUMN.match(body):
        column = column_number(body)
        if column > MAX_COLUMNS:
            raise ValueError(f"{text!r} is outside the grid")
        return None, column
    if _ROW.match(body):
        row = int(body)
        if not 1 <= row <= MAX_ROWS:
            raise ValueError(f"{text!r} is outside the grid")
        return row, None
    raise ValueError(f"{text!r} is not a cell reference")


def parse_reference(text: str, *, sheet: str = "") -> list[Area]:
    """Every rectangle in a reference, which a comma may list several of. A space between two references is where
    they meet, and binds tighter than a comma: A1:C3 B2:D4,E5 is B2:C3 and E5; references that do not meet are no
    reference (tests/fixtures/excel_model/)."""
    pieces = [piece for piece in _split_areas(text) if piece.strip()]
    if not pieces:
        raise ValueError("an empty reference")
    return [_intersection(piece, sheet) for piece in pieces]


def _intersection(text: str, sheet: str) -> Area:
    """The rectangle where the references a space separates in ``text`` meet."""
    first, *rest = [parse_area(part, sheet=sheet) for part in _split_areas(text, " ") if part]
    for area in rest:
        top, left = max(first.top, area.top), max(first.left, area.left)
        bottom, right = min(first.bottom, area.bottom), min(first.right, area.right)
        if area.sheet != first.sheet or top > bottom or left > right:
            raise ValueError(f"{text!r} names references that do not meet")
        first = Area(top, left, bottom, right, first.sheet)
    return first


def _split_areas(text: str, separator: str = ",") -> list[str]:
    """Split on each ``separator`` that is not inside a quoted sheet name."""
    out: list[str] = []
    current: list[str] = []
    quoted = False
    for char in text:
        if char == "'":
            quoted = not quoted
        if char == separator and not quoted:
            out.append("".join(current))
            current = []
            continue
        current.append(char)
    out.append("".join(current))
    return out
