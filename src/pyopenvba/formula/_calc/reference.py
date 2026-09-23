"""A1 notation: cell and range references.

The addressing layer everything in the Excel surface sits on. It is pure
arithmetic over strings and integers, with no package or XML in sight, so it
is cheap to test exhaustively.

Rows and columns are 1-based throughout, because that is what the file
format uses and what a user reading a spreadsheet sees. There is no
0-based variant anywhere in this library; mixing the two is how off-by-one
bugs get into a cell address.

Absolute markers are carried but do not affect equality of *position*: a
reference knows whether it was written ``A1``, ``$A1``, ``A$1`` or
``$A$1``, because shared-formula translation has to shift the relative
parts and leave the absolute ones alone.

Scope: this module addresses a sheet. The forms that only appear inside
formulas, whole-column ``A:A``, sheet-qualified ``Sheet1!A1``, R1C1, and
external workbook references, belong to the formula layer, which needs a
real parser rather than this.
"""

from __future__ import annotations

import re
from collections.abc import Iterator
from dataclasses import dataclass

#: The last column Excel has, ``XFD``. A reference past it is not a
#: reference Excel can store, so it is refused rather than clamped.
MAX_COLUMN = 16384
#: The last row Excel has.
MAX_ROW = 1048576

_CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$")
_AXIS_ROWS = re.compile(r"^(\$?)([1-9][0-9]*):(\$?)([1-9][0-9]*)$")
_AXIS_COLUMNS = re.compile(r"^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$")
_LETTERS = re.compile(r"^[A-Za-z]{1,3}$")


def column_letter(index: int) -> str:
    """``1`` -> ``'A'``, ``27`` -> ``'AA'``, ``16384`` -> ``'XFD'``.

    Spreadsheet columns are bijective base-26: there is no zero digit, so
    the usual base conversion is off by one at every carry.
    """
    if not 1 <= index <= MAX_COLUMN:
        raise ValueError(f"column {index} is outside 1..{MAX_COLUMN} (A..{column_letter(MAX_COLUMN)})")
    letters = ""
    remaining = index
    while remaining > 0:
        remaining, digit = divmod(remaining - 1, 26)
        letters = chr(ord("A") + digit) + letters
    return letters


def column_index(letters: str) -> int:
    """``'A'`` -> ``1``, ``'aa'`` -> ``27``. Case does not matter."""
    if not _LETTERS.match(letters):
        raise ValueError(f"{letters!r} is not a column name")
    index = 0
    for character in letters.upper():
        index = index * 26 + (ord(character) - ord("A") + 1)
    if index > MAX_COLUMN:
        raise ValueError(
            f"column {letters.upper()!r} is past {column_letter(MAX_COLUMN)}, the last column Excel has"
        )
    return index


@dataclass(frozen=True)
class CellRef:
    """One cell's address, with the absolute markers it was written with."""

    row: int
    column: int
    absolute_row: bool = False
    absolute_column: bool = False

    def __post_init__(self) -> None:
        if not 1 <= self.row <= MAX_ROW:
            raise ValueError(f"row {self.row} is outside 1..{MAX_ROW}")
        if not 1 <= self.column <= MAX_COLUMN:
            raise ValueError(f"column {self.column} is outside 1..{MAX_COLUMN}")

    @classmethod
    def parse(cls, text: str) -> CellRef:
        """Read ``A1``, ``$A1``, ``A$1`` or ``$A$1``."""
        match = _CELL.match(text.strip())
        if not match:
            raise ValueError(f"{text!r} is not a cell reference")
        dollar_column, letters, dollar_row, digits = match.groups()
        return cls(
            row=int(digits),
            column=column_index(letters),
            absolute_row=dollar_row == "$",
            absolute_column=dollar_column == "$",
        )

    @property
    def letter(self) -> str:
        """The column's name, without any absolute marker."""
        return column_letter(self.column)

    @property
    def a1(self) -> str:
        """The reference as written, absolute markers included."""
        return (
            f"{'$' if self.absolute_column else ''}{self.letter}"
            f"{'$' if self.absolute_row else ''}{self.row}"
        )

    @property
    def relative(self) -> CellRef:
        """The same position with both absolute markers dropped."""
        return CellRef(self.row, self.column)

    @property
    def sort_key(self) -> tuple[int, int]:
        """Reading order: down the rows, then across the columns.

        Excel requires a worksheet's rows to be in ascending order and each
        row's cells to be in ascending column order, so writers sort by this.
        """
        return (self.row, self.column)

    def offset(self, rows: int = 0, columns: int = 0) -> CellRef:
        """A reference moved by a delta, keeping the absolute markers."""
        return CellRef(
            row=self.row + rows,
            column=self.column + columns,
            absolute_row=self.absolute_row,
            absolute_column=self.absolute_column,
        )

    def translated(self, rows: int, columns: int) -> CellRef:
        """A reference shifted the way a formula's is when it is copied:
        relative parts move, absolute parts stay.

        This is what a shared formula needs. Excel stores the text once on
        the group's first cell and leaves the rest pointing at it by index,
        so every other cell's formula has to be derived by translating this
        way.
        """
        return CellRef(
            row=self.row if self.absolute_row else self.row + rows,
            column=self.column if self.absolute_column else self.column + columns,
            absolute_row=self.absolute_row,
            absolute_column=self.absolute_column,
        )

    def __str__(self) -> str:
        return self.a1

    def __lt__(self, other: CellRef) -> bool:
        return self.sort_key < other.sort_key

    def __le__(self, other: CellRef) -> bool:
        return self.sort_key <= other.sort_key

    def __gt__(self, other: CellRef) -> bool:
        return self.sort_key > other.sort_key

    def __ge__(self, other: CellRef) -> bool:
        return self.sort_key >= other.sort_key


@dataclass(frozen=True)
class RangeRef:
    """A rectangular block of cells, such as ``A1:C3``.

    A single cell is a range of one, so ``A1`` parses too and round-trips
    as ``A1`` rather than ``A1:A1``: that is how Excel writes a one-cell
    ``dimension``.
    """

    start: CellRef
    end: CellRef

    @classmethod
    def parse(cls, text: str) -> RangeRef:
        """Read ``A1:C3``, or a bare ``A1`` as a range of one."""
        cleaned = text.strip()
        if ":" not in cleaned:
            cell = CellRef.parse(cleaned)
            return cls(cell, cell)
        left, _, right = cleaned.partition(":")
        return cls(CellRef.parse(left), CellRef.parse(right))

    @classmethod
    def bounding(cls, cells: list[CellRef]) -> RangeRef:
        """The smallest range containing every cell given."""
        if not cells:
            raise ValueError("an empty set of cells has no bounding range")
        top = min(c.row for c in cells)
        bottom = max(c.row for c in cells)
        left = min(c.column for c in cells)
        right = max(c.column for c in cells)
        return cls(CellRef(top, left), CellRef(bottom, right))

    @property
    def normalized(self) -> RangeRef:
        """The same block with ``start`` at the top left.

        ``C3:A1`` and ``A1:C3`` denote one block; Excel writes the second.
        """
        top, bottom = sorted((self.start.row, self.end.row))
        left, right = sorted((self.start.column, self.end.column))
        return RangeRef(
            CellRef(top, left, self.start.absolute_row, self.start.absolute_column),
            CellRef(bottom, right, self.end.absolute_row, self.end.absolute_column),
        )

    @property
    def is_single_cell(self) -> bool:
        return self.start.sort_key == self.end.sort_key

    @property
    def top(self) -> int:
        return min(self.start.row, self.end.row)

    @property
    def bottom(self) -> int:
        return max(self.start.row, self.end.row)

    @property
    def left(self) -> int:
        return min(self.start.column, self.end.column)

    @property
    def right(self) -> int:
        return max(self.start.column, self.end.column)

    @property
    def height(self) -> int:
        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        return self.right - self.left + 1

    @property
    def size(self) -> int:
        """How many cells the block covers."""
        return self.height * self.width

    @property
    def a1(self) -> str:
        """The reference as Excel writes it."""
        if self.is_single_cell:
            return self.start.a1
        return f"{self.start.a1}:{self.end.a1}"

    def __contains__(self, cell: CellRef) -> bool:
        return self.top <= cell.row <= self.bottom and self.left <= cell.column <= self.right

    def cells(self) -> Iterator[CellRef]:
        """Every cell in the block, in reading order."""
        for row in range(self.top, self.bottom + 1):
            for column in range(self.left, self.right + 1):
                yield CellRef(row, column)

    def rows(self) -> Iterator[list[CellRef]]:
        """The block a row at a time, in reading order."""
        for row in range(self.top, self.bottom + 1):
            yield [CellRef(row, column) for column in range(self.left, self.right + 1)]

    def intersects(self, other: RangeRef) -> bool:
        """Whether two blocks share any cell.

        Two merged ranges may not overlap, and Excel repairs a worksheet
        where they do rather than rendering it, so this is checked before a
        merge is recorded.
        """
        return not (
            self.right < other.left
            or other.right < self.left
            or self.bottom < other.top
            or other.bottom < self.top
        )

    def contains(self, other: RangeRef) -> bool:
        """Whether this block covers all of another."""
        return (
            self.top <= other.top
            and self.left <= other.left
            and self.bottom >= other.bottom
            and self.right >= other.right
        )

    @property
    def absolute(self) -> RangeRef:
        """The same block with every marker set: ``$A$1:$C$3``.

        Which is how a defined name writes a range, because a relative one
        would move with whichever cell happened to be selected.
        """
        return RangeRef(
            CellRef(self.top, self.left, True, True),
            CellRef(self.bottom, self.right, True, True),
        )

    def expanded(self, cell: CellRef) -> RangeRef:
        """The smallest range covering this block and one more cell."""
        return RangeRef(
            CellRef(min(self.top, cell.row), min(self.left, cell.column)),
            CellRef(max(self.bottom, cell.row), max(self.right, cell.column)),
        )

    def __str__(self) -> str:
        return self.a1


@dataclass(frozen=True)
class AxisRef:
    """A whole-row or whole-column reference, such as ``A:A`` or ``2:4``.

    It names a span on one axis and every cell on the other, so it cannot be
    written as a :class:`CellRef` pair: ``A:A`` is not ``A1:A1048576``, and
    Excel keeps the shorter form when it rewrites a formula.

    The two ends still move like the ends of a range. Inserting rows at 2
    turns ``2:4`` into ``4:6``, and deleting rows 2 to 4 turns it into
    ``#REF!``, both measured against Excel.
    """

    #: ``True`` for ``2:4``, ``False`` for ``A:C``.
    is_row: bool
    low: int
    high: int
    absolute_low: bool = False
    absolute_high: bool = False

    def __post_init__(self) -> None:
        limit = MAX_ROW if self.is_row else MAX_COLUMN
        what = "row" if self.is_row else "column"
        for end in (self.low, self.high):
            if not 1 <= end <= limit:
                raise ValueError(f"{what} {end} is outside 1..{limit}")
        if self.low > self.high:
            raise ValueError(f"{what}s {self.low}..{self.high} run backwards")

    @classmethod
    def parse(cls, text: str) -> AxisRef:
        """Read ``A:C``, ``$A:$C``, ``2:4`` or ``$2:$4``."""
        match = _AXIS_ROWS.match(text.strip())
        if match:
            low_dollar, low, high_dollar, high = match.groups()
            return cls(True, int(low), int(high), low_dollar == "$", high_dollar == "$")
        match = _AXIS_COLUMNS.match(text.strip())
        if match:
            low_dollar, low, high_dollar, high = match.groups()
            return cls(
                False,
                column_index(low),
                column_index(high),
                low_dollar == "$",
                high_dollar == "$",
            )
        raise ValueError(f"{text!r} is not a whole-row or whole-column reference")

    def _end(self, value: int, absolute: bool) -> str:
        name = str(value) if self.is_row else column_letter(value)
        return f"{'$' if absolute else ''}{name}"

    @property
    def a1(self) -> str:
        """The reference as written, absolute markers included."""
        return (
            f"{self._end(self.low, self.absolute_low)}:"
            f"{self._end(self.high, self.absolute_high)}"
        )

    def with_span(self, low: int, high: int) -> AxisRef:
        """The same reference over a different span, keeping the markers."""
        return AxisRef(self.is_row, low, high, self.absolute_low, self.absolute_high)

    def __str__(self) -> str:
        return self.a1


__all__ = [
    "MAX_COLUMN",
    "MAX_ROW",
    "AxisRef",
    "CellRef",
    "RangeRef",
    "column_index",
    "column_letter",
]
