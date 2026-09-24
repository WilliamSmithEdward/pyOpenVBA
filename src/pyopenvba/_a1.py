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
_PLAIN_NAME = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


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

    Anything but a plain identifier is wrapped in apostrophes, and an
    apostrophe inside the name is doubled, because a single one would
    close the quoting early and leave a reference to something else.
    """
    if _PLAIN_NAME.fullmatch(name):
        return name
    return "'" + name.replace("'", "''") + "'"


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
