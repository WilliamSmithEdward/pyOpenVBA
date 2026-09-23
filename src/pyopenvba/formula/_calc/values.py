"""The values a formula computes with, and how Excel converts between them.

A scalar is a ``float``, a ``str``, a ``bool``, a :class:`CellError`, or
:data:`EMPTY`, what a blank cell holds. Excel has no integers, so every
number is a float here, and there is no ``None``: a blank cell and an empty
argument are :data:`EMPTY`, which reads as 0, ``""`` or FALSE as needed.

Besides scalars an expression can give a :class:`Reference` to cells, which
a function such as ROW or OFFSET needs as such, and an :class:`Array` of
scalars, from ``{1,2;3,4}`` or from arithmetic on a range in an array
formula.

**Coercion** is measured, not assumed. Text becomes a number when it reads
as one typed into a cell would: ``"1,000"``, ``"$5"``, ``"(5)"`` for -5,
``"50%"``, ``"1 1/2"``, ``"1/2/2020"``, ``"12:30"``. Only spaces are
trimmed, not tabs; ``"TRUE"`` is not a number, and neither is ``""``.

**Order** across types is numbers, then text, then logicals, so ``1<"a"``
and ``"a"<TRUE``. Text compares as the filters do, by
:mod:`~pyopenvba.formula._calc.collate`: case and hyphens aside, and ``"\u00df"``
equal to ``"ss"``. A blank compares as the zero of the other side's type,
so a blank cell equals 0, ``""`` and FALSE alike.
"""

from __future__ import annotations

import datetime as dt
import re
from collections.abc import Iterator
from dataclasses import dataclass

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.dates import parse_date_time
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.numbers import compare_numbers, number_text, text_value
from pyopenvba.formula._calc.cells import CellError


class Empty:
    """What a blank cell holds, and an argument left empty."""

    __slots__ = ()
    _instance: Empty | None = None

    def __new__(cls) -> Empty:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __repr__(self) -> str:
        return "EMPTY"


EMPTY = Empty()

Scalar = float | str | bool | CellError | Empty

NULL = CellError("#NULL!")
DIV0 = CellError("#DIV/0!")
VALUE = CellError("#VALUE!")
REF = CellError("#REF!")
NAME = CellError("#NAME?")
NUM = CellError("#NUM!")
NA = CellError("#N/A")
SPILL = CellError("#SPILL!")
CALC = CellError("#CALC!")

#: The error codes ERROR.TYPE numbers, in its order.
ERROR_NUMBERS = {
    "#NULL!": 1,
    "#DIV/0!": 2,
    "#VALUE!": 3,
    "#REF!": 4,
    "#NAME?": 5,
    "#NUM!": 6,
    "#N/A": 7,
    "#GETTING_DATA": 8,
    "#SPILL!": 9,
    "#CONNECT!": 10,
    "#BLOCKED!": 11,
    "#UNKNOWN!": 12,
    "#FIELD!": 13,
    "#CALC!": 14,
}

#: The longest text a cell holds.
MAX_TEXT = 32767


class ExcelError(Exception):
    """Raised inside a calculation to make its result an error value."""

    def __init__(self, error: CellError) -> None:
        super().__init__(error.code)
        self.error = error


@dataclass(frozen=True, slots=True)
class Area:
    """A block of cells on one sheet, by its edges."""

    sheet: str
    top: int
    left: int
    bottom: int
    right: int

    @property
    def height(self) -> int:
        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        return self.right - self.left + 1

    @property
    def is_cell(self) -> bool:
        return self.top == self.bottom and self.left == self.right

    def contains(self, row: int, column: int) -> bool:
        return self.top <= row <= self.bottom and self.left <= column <= self.right

    def intersection(self, other: Area) -> Area | None:
        if self.sheet != other.sheet:
            return None
        top, bottom = max(self.top, other.top), min(self.bottom, other.bottom)
        left, right = max(self.left, other.left), min(self.right, other.right)
        if top > bottom or left > right:
            return None
        return Area(self.sheet, top, left, bottom, right)

    def cells(self) -> Iterator[tuple[int, int]]:
        for row in range(self.top, self.bottom + 1):
            for column in range(self.left, self.right + 1):
                yield row, column


@dataclass(frozen=True, slots=True)
class Reference:
    """Cells, as one or more areas: more than one only from the union
    operator or a 3D reference."""

    areas: tuple[Area, ...]

    @classmethod
    def of(cls, area: Area) -> Reference:
        return cls((area,))

    @property
    def area(self) -> Area | None:
        """The one area, or ``None`` for a reference with several."""
        return self.areas[0] if len(self.areas) == 1 else None


class Array:
    """A rectangle of scalars, ``rows[row][column]``."""

    __slots__ = ("rows",)

    def __init__(self, rows: list[list[Scalar]]) -> None:
        self.rows = rows

    @classmethod
    def filled(cls, height: int, width: int, value: Scalar) -> Array:
        return cls([[value] * width for _ in range(height)])

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    def at(self, row: int, column: int) -> Scalar:
        """The item at a position of a broadcast: a single row or column
        stretches, and past the end of a longer side is ``#N/A``."""
        if self.height == 1:
            row = 0
        if self.width == 1:
            column = 0
        if row >= self.height or column >= self.width:
            return NA
        return self.rows[row][column]

    def items(self) -> Iterator[Scalar]:
        for row in self.rows:
            yield from row

    def __eq__(self, other: object) -> bool:
        return isinstance(other, Array) and self.rows == other.rows

    def __hash__(self) -> int:
        return hash(tuple(tuple(row) for row in self.rows))

    def __repr__(self) -> str:
        return f"Array({self.rows!r})"


Value = Scalar | Reference | Array


@dataclass(frozen=True, eq=False)
class Scope:
    """The names a LET binds, or a LAMBDA's call: upper-cased, as a formula
    stores them, ``_XLPM.X``. ``omitted`` are the parameters a call left
    out, which ISOMITTED asks about."""

    values: dict[str, Value]
    omitted: frozenset[str] = frozenset()


@dataclass(frozen=True, eq=False)
class Lambda(CellError):  # noqa: N818 - an error value, not an exception
    """A function LAMBDA made, carried where a value goes.

    It is an error value, ``#CALC!``, because that is what Excel shows for
    a LAMBDA left uncalled in a cell; MAP, REDUCE and the rest, and a call
    such as ``LAMBDA(x,x*2)(4)``, use it as the function it is. It keeps
    the names in scope where it was made.
    """

    parameters: tuple[str, ...] = ()
    body: Node | None = None
    closure: tuple[Scope, ...] = ()

    @classmethod
    def make(cls, parameters: tuple[str, ...], body: Node, closure: tuple[Scope, ...]) -> Lambda:
        return cls(CALC.code, parameters, body, closure)


# ----------------------------------------------------------------------
# Text as a number
# ----------------------------------------------------------------------

_CURRENCY = "$\u20ac"
_MANTISSA = re.compile(r"(?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)?(?:\.[0-9]*)?")
_EXPONENT = re.compile(r"[eE][+-]?[0-9]+")
_FRACTION = re.compile(r" ([0-9]+)/([0-9]+)")


def text_to_number(text: str, today: dt.date, *, epoch_1904: bool = False) -> float | None:
    """The number text stands for, as typing it into a cell would make it,
    or ``None`` when it is not one."""
    body = text.strip(" ")
    if not body:
        return None
    number = _plain_number(body)
    if number is not None:
        return number
    return parse_date_time(body, today, epoch_1904=epoch_1904)


def _skip(text: str, index: int) -> int:
    while index < len(text) and text[index] == " ":
        index += 1
    return index


def _plain_number(text: str) -> float | None:
    """A number with its sign, currency, percent and parentheses."""
    index = 0
    sign = ""
    currency = percent = parenthesis = False
    while True:
        index = _skip(text, index)
        if index >= len(text):
            return None
        char = text[index]
        if char in "+-" and not sign and not parenthesis:
            sign = char
        elif char in _CURRENCY and not currency:
            currency = True
        elif char == "(" and not parenthesis and not sign:
            parenthesis = True
        elif char == "%" and not percent:
            percent = True
        else:
            break
        index += 1

    mantissa = _MANTISSA.match(text, index)
    digits = mantissa.group(0) if mantissa else ""
    if not any(char.isdigit() for char in digits):
        return None
    index += len(digits)
    exponent = _EXPONENT.match(text, index)
    fraction = None
    if exponent:
        index = exponent.end()
    elif "," not in digits and "." not in digits:
        fraction = _FRACTION.match(text, index)
        if fraction:
            index = fraction.end()
    found = text_value(digits.replace(",", "") + (exponent.group(0) if exponent else ""))
    if found is None:
        return None
    value = found
    if fraction:
        denominator = int(fraction.group(2))
        if denominator == 0:
            return None
        value = value + int(fraction.group(1)) / denominator

    while True:
        index = _skip(text, index)
        if index >= len(text):
            break
        char = text[index]
        if char == "%" and not percent:
            percent = True
        elif char == ")" and parenthesis:
            parenthesis = False
            sign = "-"
        else:
            return None
        index += 1
    if parenthesis:
        return None
    if percent:
        value = value / 100
    # Adding 0.0 turns -0.0, which Excel does not have, into 0.
    return (-value if sign == "-" else value) + 0.0


# ----------------------------------------------------------------------
# Comparison
# ----------------------------------------------------------------------


def _rank(value: Scalar) -> int:
    if isinstance(value, bool):
        return 2
    if isinstance(value, str):
        return 1
    return 0


def _zero_like(value: Scalar) -> Scalar:
    if isinstance(value, bool):
        return False
    if isinstance(value, str):
        return ""
    return 0.0


def compare(left: Scalar, right: Scalar) -> int:
    """-1, 0 or 1 as ``left`` sorts before, level with or after ``right``.
    Neither may be an error."""
    if isinstance(left, Empty):
        left = _zero_like(right)
    if isinstance(right, Empty):
        right = _zero_like(left)
    rank_left, rank_right = _rank(left), _rank(right)
    if rank_left != rank_right:
        return -1 if rank_left < rank_right else 1
    if isinstance(left, bool) and isinstance(right, bool):
        return (left > right) - (left < right)
    if isinstance(left, str) and isinstance(right, str):
        return _collate.compare(left, right)
    assert isinstance(left, float) and isinstance(right, float)
    return compare_numbers(left, right)


def scalar_text(value: Scalar) -> str:
    """A non-error scalar as text: numbers to fifteen digits, logicals in
    capitals, a blank as nothing."""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return number_text(value)
    if isinstance(value, str):
        return value
    return ""


__all__ = [
    "CALC",
    "DIV0",
    "EMPTY",
    "ERROR_NUMBERS",
    "MAX_TEXT",
    "NA",
    "NAME",
    "NULL",
    "NUM",
    "REF",
    "SPILL",
    "VALUE",
    "Area",
    "Array",
    "Empty",
    "ExcelError",
    "Lambda",
    "Reference",
    "Scalar",
    "Scope",
    "Value",
    "compare",
    "scalar_text",
    "text_to_number",
]
