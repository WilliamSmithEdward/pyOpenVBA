"""What a formula computes with.

Excel's value model is not VBA's.  There is one number type rather than
six, an error is a value that travels through an expression rather than
an exception, an empty cell is its own thing, and a block of cells is a
value in its own right.  Keeping it separate from
:mod:`pyopenvba.interpreter._values` is what stops VBA's Integer widths
and Null leaking into a grid that has neither.
"""

from __future__ import annotations

import math
import struct
from dataclasses import dataclass
from decimal import ROUND_HALF_DOWN, ROUND_HALF_UP, Decimal, localcontext
from typing import Final, Iterator


class ExcelError(Exception):
    """One of the seven values a cell shows when a formula cannot finish.

    An exception so that raising one unwinds an expression the way Excel
    abandons it, and a value so that IFERROR and ISERROR can hold one.
    """

    __slots__ = ("name",)

    def __init__(self, name: str) -> None:
        super().__init__(name)
        self.name = name

    def __eq__(self, other: object) -> bool:
        return isinstance(other, ExcelError) and other.name == self.name

    def __hash__(self) -> int:
        return hash(("ExcelError", self.name))

    def __repr__(self) -> str:
        return self.name


DIV0: Final = ExcelError("#DIV/0!")
NA: Final = ExcelError("#N/A")
NAME: Final = ExcelError("#NAME?")
NULL: Final = ExcelError("#NULL!")
NUM: Final = ExcelError("#NUM!")
REF: Final = ExcelError("#REF!")
VALUE: Final = ExcelError("#VALUE!")
SPILL: Final = ExcelError("#SPILL!")
CALC: Final = ExcelError("#CALC!")

#: The errors a formula may name as a literal, and CVErr's numbers.
ERRORS: Final[dict[str, ExcelError]] = {
    error.name: error for error in (DIV0, NA, NAME, NULL, NUM, REF, VALUE, SPILL, CALC)
}

#: What Excel's ERROR.TYPE answers for each.
ERROR_NUMBERS: Final[dict[str, int]] = {
    "#NULL!": 1,
    "#DIV/0!": 2,
    "#VALUE!": 3,
    "#REF!": 4,
    "#NAME?": 5,
    "#NUM!": 6,
    "#N/A": 7,
    "#GETTING_DATA": 8,
    "#SPILL!": 9,
    "#CALC!": 14,
}


class _Blank:
    """An empty cell: zero in arithmetic, "" in text, and its own type."""

    __slots__ = ()

    def __repr__(self) -> str:
        return "<blank>"

    def __bool__(self) -> bool:
        return False


BLANK: Final = _Blank()

#: What a formula may compute to.
Value = "float | str | bool | ExcelError | _Blank | Matrix"


@dataclass(slots=True)
class Matrix:
    """A block of values: a range's contents, or an array constant.

    Rows of columns, one-based nowhere: this is the engine's own shape,
    and the worksheet's coordinates are applied by whoever built it.
    """

    rows: list[list[object]]

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.rows[0]) if self.rows else 0

    @property
    def single(self) -> bool:
        return self.height == 1 and self.width == 1

    def at(self, row: int, column: int) -> object:
        """One element, repeating a single row or column as Excel does."""
        if not self.rows:
            return NA
        line = self.rows[0] if self.height == 1 else self.rows[row] if row < self.height else None
        if line is None:
            return NA
        if self.width == 1:
            return line[0]
        return line[column] if column < self.width else NA

    def flat(self) -> Iterator[object]:
        for line in self.rows:
            yield from line

    def first(self) -> object:
        return self.rows[0][0] if self.rows and self.rows[0] else BLANK


# --- conversions --------------------------------------------------------------------


def as_number(value: object) -> float:
    """``value`` as a number, or the error Excel raises instead."""
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if value is BLANK:
        return 0.0
    if isinstance(value, ExcelError):
        raise value
    if isinstance(value, Matrix):
        return as_number(value.first())
    if isinstance(value, str):
        number = text_as_number(value)
        if number is None:
            raise VALUE
        return number
    raise VALUE


def text_as_number(text: str) -> float | None:
    """A string read as a number, the way a formula reads one.

    Excel takes the ordinary spellings and the two decorations that go
    with them: a percent sign divides by a hundred, and a currency sign
    is ignored.  Anything else is not a number.
    """
    body = text.strip()
    if not body:
        return None
    scale = 1.0
    if body.endswith("%"):
        scale = 0.01
        body = body[:-1].strip()
    negative = False
    if body.startswith("(") and body.endswith(")"):
        negative = True
        body = body[1:-1].strip()
    for sign in ("$", "£", "€", "¥"):
        if body.startswith(sign):
            body = body[len(sign) :].strip()
        elif body.startswith("-" + sign):
            body = "-" + body[1 + len(sign) :].strip()
    body = body.replace(",", "") if _grouped(body) else body
    try:
        number = float(body)
    except ValueError:
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return -number * scale if negative else number * scale


def _grouped(body: str) -> bool:
    """Whether the commas in a number are digit grouping."""
    if "," not in body:
        return False
    head = body.lstrip("+-")
    whole, _, rest = head.partition(".")
    if "," in rest:
        return False
    pieces = whole.split(",")
    if len(pieces) < 2 or not pieces[0] or not pieces[0].isdigit():
        return False
    return all(len(piece) == 3 and piece.isdigit() for piece in pieces[1:])


def as_text(value: object) -> str:
    """``value`` as text, the way ``&`` and TEXT-less conversions make it."""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if value is BLANK:
        return ""
    if isinstance(value, ExcelError):
        raise value
    if isinstance(value, Matrix):
        return as_text(value.first())
    if isinstance(value, (int, float)):
        return number_text(float(value))
    return str(value)


def number_text(number: float, *, formula: bool = False) -> str:
    """A number as Excel spells it in text: fifteen significant digits, in as few characters as allowed.

    Leaving the sign aside, ``=A1&""`` spells a number in at most twenty
    characters and a cell's Formula, ``formula``, in twenty-one: written
    out in full while that fits -- 12345678901234500000, 0.000001 -- and
    otherwise with an exponent of two digits at least, the mantissa losing
    digits until it fits, as 1.2345678901235E+100 does. Measured across
    exponents from -25 to 300.
    """
    if number != number:
        return "#NUM!"
    if number == 0:
        return "0"
    if number in (float("inf"), float("-inf")):
        return "#NUM!"
    widest = 21 if formula else 20
    sign = "-" if number < 0 else ""
    for precision in range(15, 0, -1):
        with localcontext() as context:
            # An exact tie goes toward zero: 4503599627370495 is 4503599627370490 (tests/fixtures/formula).
            context.prec, context.rounding = precision, ROUND_HALF_DOWN
            rounded = (+Decimal(abs(number))).normalize()
        if precision == 15 and len(plain := format(rounded, "f")) <= widest:
            return sign + plain
        _, digits, exponent = rounded.as_tuple()
        assert isinstance(exponent, int)
        highest = exponent + len(digits) - 1
        mantissa = "".join(str(digit) for digit in digits)
        body = mantissa[0] + ("." + mantissa[1:] if len(mantissa) > 1 else "")
        text = f"{body}E{'+' if highest >= 0 else '-'}{abs(highest):02d}"
        if len(text) <= widest:
            return sign + text
    raise AssertionError("a one-digit mantissa always fits")


def as_bool(value: object) -> bool:
    """``value`` as a condition."""
    if isinstance(value, bool):
        return value
    if isinstance(value, ExcelError):
        raise value
    if isinstance(value, str):
        body = value.strip().upper()
        if body == "TRUE":
            return True
        if body == "FALSE":
            return False
        raise VALUE
    if isinstance(value, Matrix):
        return as_bool(value.first())
    return as_number(value) != 0


@dataclass(frozen=True, slots=True)
class Areas:
    """Several blocks named at once, as (A1:A2,A4) names them.

    SUM, COUNT, AVERAGE, MAX, LARGE and the like read every cell of
    every block; where one value is wanted it is #VALUE!, as =(A1:A2,A4)
    is in a cell (tests/fixtures/reference_forms.json).
    """

    blocks: tuple[Matrix, ...]

    def joined(self) -> Matrix:
        """Every cell of every block, in order, as one column."""
        return Matrix([[item] for block in self.blocks for item in block.flat()])


def single(value: object) -> object:
    """A block reduced to the one value an operator can use."""
    if isinstance(value, Areas):
        return VALUE
    return value.first() if isinstance(value, Matrix) else value


def is_error(value: object) -> bool:
    return isinstance(value, ExcelError)


#: How Excel orders the types when two of them are compared: a number
#: comes before any text, and text before FALSE, which comes before TRUE.
_ORDER: Final = {"number": 0, "text": 1, "bool": 2}


def _kind(value: object) -> str:
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, str):
        return "text"
    return "number"


def _numeric_order(one: float, two: float) -> int:
    """Two numbers compared as a cell compares them.

    Excel settles a comparison on each number rounded to fifteen
    significant digits rather than on the bits, which is why =0.1+0.2=0.3
    is TRUE in a sheet and False almost everywhere else. An exact tie
    rounds away from zero, unlike the number's text: 4503599627370495
    equals 4503599627370500 but not 4503599627370490. Measured over 1,400
    pairs a few bits apart (tests/fixtures/zero_snap.json).
    """
    if one != one or two != two:
        return 0
    first = _to_fifteen(one)
    second = _to_fifteen(two)
    return (first > second) - (first < second)


def _to_fifteen(number: float) -> Decimal:
    with localcontext() as context:
        context.prec, context.rounding = 15, ROUND_HALF_UP
        return +Decimal(number)


def snapped(left: float, total: float) -> float:
    """A sum or difference that all but cancels, set to zero as Excel sets it.

    Measured bit for bit (scripts/measure_zero_snap.py): the last + or -
    of a formula, and the last addition SUM and AVERAGE make, is 0 when
    the answer's binary exponent is 50 or more below the left operand's.
    That is seven steps of the last bit within one binade, and not a
    matter of decimal digits: 3.1091263510296 below the double nearest
    STDEV(1,6,7,8) keeps its 4.9E-15. The exponents are the fields as
    stored, so an answer too small to be normal keeps its bits.
    """
    if total != 0 and _exponent(total) <= _exponent(left) - 50:
        return 0.0
    return total


def _exponent(number: float) -> int:
    """A double's biased exponent field: 0 for zero and the subnormals."""
    return (struct.unpack("<Q", struct.pack("<d", number))[0] >> 52) & 0x7FF


def compare(op: str, left: object, right: object, *, exact: bool = False) -> bool:
    """A comparison, with Excel's own ordering between the types.

    Numbers are compared to fifteen digits, as the operators and the
    criteria functions compare them; ``exact`` compares their bits, as
    MATCH, VLOOKUP, HLOOKUP and XLOOKUP do.
    """
    first = single(left)
    second = single(right)
    if isinstance(first, ExcelError):
        raise first
    if isinstance(second, ExcelError):
        raise second
    # An empty cell takes the shape of whatever it is compared against,
    # which is why =A1="" and =A1=0 are both TRUE for an empty cell.
    if first is BLANK:
        first = "" if isinstance(second, str) else (False if isinstance(second, bool) else 0.0)
    if second is BLANK:
        second = "" if isinstance(first, str) else (False if isinstance(first, bool) else 0.0)
    if _kind(first) != _kind(second):
        order = _ORDER[_kind(first)] - _ORDER[_kind(second)]
    elif isinstance(first, str) and isinstance(second, str):
        # Text comparison ignores case.
        upper_first, upper_second = first.upper(), second.upper()
        order = (upper_first > upper_second) - (upper_first < upper_second)
    else:
        one, two = float(as_number(first)), float(as_number(second))
        order = ((one > two) - (one < two)) if exact else _numeric_order(one, two)
    if op == "=":
        return order == 0
    if op == "<>":
        return order != 0
    if op == "<":
        return order < 0
    if op == ">":
        return order > 0
    if op == "<=":
        return order <= 0
    return order >= 0
