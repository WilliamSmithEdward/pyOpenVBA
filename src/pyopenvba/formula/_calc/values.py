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
from dataclasses import dataclass, field

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.dates import parse_date_time
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.numbers import compare_numbers, number_text, text_value
from pyopenvba.formula._calc.cells import CellError, UnsupportedFormulaError
from pyopenvba.interpreter._locale import EN_US, Locale


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


class Omitted(Empty):
    """An argument left out, as a function hands one back: CHOOSE(1,). It is blank to every formula,
    ``CHOOSE(1,)&"x"`` being x, and VBA alone reads it back as 0, where it reads a blank cell as Empty
    (pyOpenVBA's tests/fixtures/formula/ and tests/fixtures/worksheet_functions.json)."""

    __slots__ = ()
    _instance: Empty | None = None

    def __repr__(self) -> str:
        return "OMITTED"


OMITTED = Omitted()

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

    @property
    def external(self) -> bool:
        """Whether the area is on a sheet of another workbook, read from its
        link's cache: such a sheet is named ``[1]Sheet1``, which no sheet of
        the workbook itself can be, a sheet's name never holding ``[``."""
        return self.sheet.startswith("[")

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
    #: INDIRECT keeps its reference identity for most consumers, but Excel's
    #: Evaluate passes it differently to GETPIVOTDATA. This is provenance only.
    indirect: bool = field(default=False, compare=False, repr=False)
    #: A 3D reference, Jan:Mar!A1, an area on each of two or more sheets:
    #: no value, only what the functions that read every sheet take
    #: (pyOpenVBA's tests/fixtures/three_d_functions.json).
    three_d: bool = False

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

    A function's name passed where a LAMBDA goes, ``BYROW(A1:B2,SUM)``,
    which a file writes ``_xleta.SUM``, is one too: ``builtin`` names the
    function, called with whatever arguments the caller passes.

    ``required`` is how many arguments a call must pass: the parameters
    before the optional ones, which a file writes ``_xlop.y`` and the
    screen ``[y]``, or the fewest a function passed by name takes.
    """

    parameters: tuple[str, ...] = ()
    body: Node | None = None
    closure: tuple[Scope, ...] = ()
    builtin: str | None = None
    required: int = 0

    @classmethod
    def make(cls, parameters: tuple[str, ...], body: Node, closure: tuple[Scope, ...], required: int) -> Lambda:
        return cls(CALC.code, parameters, body, closure, required=required)

    @classmethod
    def eta(cls, name: str, required: int) -> Lambda:
        """The function ``name`` as a value, as ``_xleta.SUM`` gives it."""
        return cls(CALC.code, builtin=name, required=required)


# ----------------------------------------------------------------------
# Text as a number
# ----------------------------------------------------------------------

_CURRENCY = "$\u20ac"
#: Measured: digits may be grouped by commas when the first group is not
#: all zeros and every later group has three digits or more, so
#: ``1,0000`` is 10000 and ``0,123`` is not a number.
_MANTISSA = re.compile(r"(?:0*[1-9][0-9]*(?:,[0-9]{3,})+|[0-9]+)?(?:\.[0-9]*)?")
_EXPONENT = re.compile(r"[eE][+-]?[0-9]+")
_FRACTION = re.compile(r" ([0-9]+)/([0-9]+)")
#: The largest numerator or denominator a fraction may have, measured.
_FRACTION_PART = 32767


def text_to_number(text: str, today: dt.date, *, epoch_1904: bool = False, locale: Locale = EN_US) -> float | None:
    """The number text stands for, as typing it into a cell would make it,
    or ``None`` when it is not one; under ``locale``'s Regional format."""
    number = plain_number(text, locale=locale)
    if number is not None:
        return number
    return read_date_time(text, today, epoch_1904=epoch_1904, locale=locale)


#: Three numbers apart by slashes or dashes, alone; and at the start of text, a time perhaps after them.
_THREE_NUMBERS = re.compile(r" *([0-9]{1,4}) *([/-]) *([0-9]{1,4}) *\2 *([0-9]{1,4}) *")
_STARTS_THREE_NUMBERS = re.compile(r" *[0-9]{1,2} *[/-] *[0-9]{1,4} *[/-] *[0-9]{1,4}")


def read_date_time(text: str, today: dt.date, *, epoch_1904: bool = False, locale: Locale = EN_US) -> float | None:
    """The serial a date or time typed as text stands for, as DATEVALUE and a
    formula's text read it, or ``None``.

    Under ja-JP three numbers are the year, the month and the day and
    nothing else: 26/1/5 is 5 January 2026 and 1/5/2026 no date
    (tests/fixtures/excel_locale_live.json). A month's name, which ja-JP
    reads by rules not measured, reports itself unsupported.
    """
    if locale.currency_symbol == "$":
        return parse_date_time(text, today, epoch_1904=epoch_1904)
    if re.search(r"[A-Za-z]{3}", text):
        raise UnsupportedFormulaError(f"reading a month's name in text as a date under {locale.name}")
    found = _THREE_NUMBERS.fullmatch(text)
    if found is not None:
        year, month, day = (int(found.group(place)) for place in (1, 3, 4))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return None
        year = year + (2000 if year < 30 else 1900) if year < 100 else year
        return parse_date_time(f"{year}/{month}/{day}", today, epoch_1904=epoch_1904)
    if _STARTS_THREE_NUMBERS.match(text):
        raise UnsupportedFormulaError(f"reading a date with no four-digit year first and a time under {locale.name}")
    return parse_date_time(text, today, epoch_1904=epoch_1904)


def plain_number(text: str, *, locale: Locale = EN_US) -> float | None:
    """The number text stands for when it is written as one, with its sign,
    currency or percent, rather than as a date or a time; ``None`` when it
    is not. Measured, this is the text Excel's error checking calls a
    number stored as text. Under ja-JP the currency sign is ¥, not $."""
    body = text.strip(" ")
    return _plain_number(body, locale.currency_symbol + _CURRENCY[1:]) if body else None


def _skip(text: str, index: int) -> int:
    while index < len(text) and text[index] == " ":
        index += 1
    return index


def _plain_number(text: str, signs: str = _CURRENCY) -> float | None:
    """A number with its sign, currency or percent, and parentheses.
    Measured: a currency sign and a percent sign together are not one."""
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
        elif char in signs and not currency and not percent:
            currency = True
        elif char == "(" and not parenthesis and not sign:
            parenthesis = True
        elif char == "%" and not percent and not currency:
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
    elif "." not in digits:
        # Measured: a fraction may follow digits grouped by thousands.
        fraction = _FRACTION.match(text, index)
        if fraction:
            index = fraction.end()
    found = text_value(digits.replace(",", "") + (exponent.group(0) if exponent else ""))
    if found is None:
        return None
    value = found
    if fraction:
        numerator, denominator = int(fraction.group(1)), int(fraction.group(2))
        if denominator == 0 or numerator > _FRACTION_PART or denominator > _FRACTION_PART:
            return None
        value = value + numerator / denominator

    while True:
        index = _skip(text, index)
        if index >= len(text):
            break
        char = text[index]
        if char == "%" and not percent and not currency:
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
    "OMITTED",
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
    "Omitted",
    "Reference",
    "Scalar",
    "Scope",
    "Value",
    "compare",
    "plain_number",
    "scalar_text",
    "read_date_time",
    "text_to_number",
]
