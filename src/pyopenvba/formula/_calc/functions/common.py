"""What several functions share: the numbers in their arguments, and the
criteria of COUNTIF and its relatives.

**Numbers.** SUM and most of the statistics count only numbers inside a
range or an array, skipping text, logicals and blanks, but take a value
typed as an argument as a number if it can be one: ``SUM(1,"2",TRUE)`` is
4 while ``SUM(A1:A3)`` over the same three values is 1. An error anywhere
is the result. The functions ending in A count text in a range as 0 and a
logical as 1 or 0.

**Criteria.** ``">5"``, ``"<>x"``, ``"a*"``, ``10``, ``TRUE``, ``""``: an
operator and a value, the value a number when it reads as one. A number
criterion matches a number cell, or with ``=`` a text cell that reads as
the same number; a text criterion matches text, with ``*``, ``?`` and
``~`` as wildcards; ``<>`` matches everything the rest would not, blanks
included; ``""`` matches blanks.
"""

from __future__ import annotations

from dataclasses import dataclass

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.lexer import ERROR_CODES
from pyopenvba.formula._calc.values import (
    EMPTY,
    NA,
    VALUE,
    Area,
    Array,
    Empty,
    ExcelError,
    Reference,
    Scalar,
    Value,
    compare,
    text_to_number,
)
from pyopenvba.formula._calc.cells import CellError


def numbers(context: Context, args: tuple[Value, ...], *, logicals: bool = False, text: bool = False) -> list[float]:
    """The numbers in the arguments, as SUM reads them; with ``logicals``
    and ``text``, as SUMA-style functions read a range."""
    found: list[float] = []
    for arg in args:
        if isinstance(arg, (Reference, Array)):
            for value, _ in context.scalars(arg):
                if isinstance(value, bool):
                    if logicals:
                        found.append(1.0 if value else 0.0)
                elif isinstance(value, float):
                    found.append(value)
                elif isinstance(value, CellError):
                    raise ExcelError(value)
                elif isinstance(value, str) and text:
                    found.append(0.0)
        else:
            found.append(context.number(arg))
    return found


def area_of(value: Value) -> Area:
    """The one area of a range argument, or ``#VALUE!``."""
    if isinstance(value, Reference) and value.area is not None:
        return value.area
    raise ExcelError(VALUE)


def matrix(context: Context, value: Value) -> Array:
    """A range, an array or a scalar as an array of its values."""
    if isinstance(value, CellError):
        raise ExcelError(value)
    return context.array_of(value)


def flat(context: Context, value: Value) -> list[Scalar]:
    """Every value of a range or array, row by row, blanks included."""
    return list(matrix(context, value).items())


def pairs(context: Context, first: Value, second: Value) -> list[tuple[float, float]]:
    """Two ranges or arrays read position by position, keeping the
    positions where both hold numbers. Different sizes are ``#N/A``, and
    an error in either is the result."""
    left = matrix(context, first)
    right = matrix(context, second)
    if left.height * left.width != right.height * right.width:
        raise ExcelError(NA)
    found: list[tuple[float, float]] = []
    for x, y in zip(left.items(), right.items(), strict=True):
        if isinstance(x, CellError):
            raise ExcelError(x)
        if isinstance(y, CellError):
            raise ExcelError(y)
        if isinstance(x, float) and isinstance(y, float):
            found.append((x, y))
    return found


# ----------------------------------------------------------------------
# Criteria
# ----------------------------------------------------------------------

_OPERATORS = ("<=", ">=", "<>", "=", "<", ">")


@dataclass(frozen=True)
class Criterion:
    """One COUNTIF-style criterion, parsed."""

    op: str
    value: Scalar
    #: For a text criterion with ``=`` or ``<>``: whether it has wildcards.
    pattern: bool = False

    def matches(self, cell: Scalar, context: Context) -> bool:
        op, value = self.op, self.value
        if isinstance(value, Empty):
            # "" is blank or empty text; "=" only a cell with nothing in it,
            # and "<>" anything but that.
            if op == "":
                return isinstance(cell, Empty) or cell == ""
            empty = isinstance(cell, Empty)
            return empty if op == "=" else not empty
        if op == "<>":
            # Not the complement of "=": "<>10" keeps the text "10", which
            # "=10" counts. Only a cell of the criterion's own type can be
            # the value excluded.
            return not self._same(cell)
        if isinstance(value, float):
            if isinstance(cell, float):
                return _holds(op, compare(cell, value))
            if op == "=" and isinstance(cell, str):
                number = text_to_number(cell, context.today, epoch_1904=context.epoch_1904)
                return number is not None and compare(number, value) == 0
            return False
        if isinstance(value, bool):
            return isinstance(cell, bool) and _holds(op, compare(cell, value))
        if isinstance(value, CellError):
            return isinstance(cell, CellError) and op == "=" and cell.code == value.code
        assert isinstance(value, str)
        if not isinstance(cell, str):
            return False
        if op == "=":
            return self._same(cell)
        return _holds(op, _collate.compare(cell, value))

    def _same(self, cell: Scalar) -> bool:
        """Whether ``cell`` is the criterion's value, in the value's type."""
        value = self.value
        if isinstance(value, bool) or isinstance(cell, bool):
            return isinstance(value, bool) and isinstance(cell, bool) and value == cell
        if isinstance(value, float):
            return isinstance(cell, float) and compare(cell, value) == 0
        if isinstance(value, CellError):
            return isinstance(cell, CellError) and cell.code == value.code
        if isinstance(value, str) and isinstance(cell, str):
            if self.pattern:
                return _collate.wildcard_match(value, cell)
            return _collate.equal(_collate.unescape(value), cell)
        return False


def _holds(op: str, order: int) -> bool:
    if op == "=":
        return order == 0
    if op == "<":
        return order < 0
    if op == ">":
        return order > 0
    if op == "<=":
        return order <= 0
    return order >= 0


def criterion(context: Context, value: Scalar) -> Criterion:
    """Parse a criterion as COUNTIF, SUMIF and the rest take it."""
    if isinstance(value, Empty):
        # A blank cell given as the criterion means 0, as Microsoft
        # documents for COUNTIFS.
        return Criterion("=", 0.0)
    if not isinstance(value, str):
        return Criterion("=", value)
    op = "="
    rest = value
    for candidate in _OPERATORS:
        if value.startswith(candidate):
            op, rest = candidate, value[len(candidate) :]
            break
    if rest == "":
        # Nothing after the operator. "" alone is the blank-or-empty test,
        # kept apart from "=" by an empty operator.
        return Criterion("" if value == "" else op if op in ("=", "<>") else "=", EMPTY)
    number = text_to_number(rest, context.today, epoch_1904=context.epoch_1904)
    if number is not None:
        return Criterion(op, number)
    upper = rest.upper()
    if upper in ("TRUE", "FALSE"):
        return Criterion(op, upper == "TRUE")
    if upper in ERROR_CODES:
        return Criterion(op, CellError(upper))
    return Criterion(op, rest, _collate.has_wildcards(rest) if op in ("=", "<>") else False)


def bound(context: Context, areas: list[Area]) -> tuple[int, int]:
    """How many rows and columns from each area's corner hold anything in
    any of them: past that, every one of them is blank."""
    height = width = 0
    for area in areas:
        last_row, last_column = context.book.used(area.sheet)
        height = max(height, min(area.bottom, last_row) - area.top + 1)
        width = max(width, min(area.right, last_column) - area.left + 1)
    return max(height, 0), max(width, 0)


def block(context: Context, area: Area, height: int, width: int) -> list[list[Scalar]]:
    """The values in the first ``height`` rows and ``width`` columns of an
    area, blanks as EMPTY."""
    rows: list[list[Scalar]] = [[EMPTY] * width for _ in range(height)]
    if height <= 0 or width <= 0:
        return rows
    clipped = Area(area.sheet, area.top, area.left, area.top + height - 1, area.left + width - 1)
    for row, column, value in context.book.cells(clipped):
        rows[row - area.top][column - area.left] = value
    return rows


def shaped(anchor: Area, like: Area) -> Area:
    """``anchor``'s top left corner stretched to ``like``'s size, which is
    how SUMIF reads a sum range of another size."""
    return Area(anchor.sheet, anchor.top, anchor.left, anchor.top + like.height - 1, anchor.left + like.width - 1)


def matching(
    context: Context, pairs: list[tuple[Area, Criterion]], extra: list[Area]
) -> tuple[list[tuple[int, int]], int]:
    """The offsets, from each range's corner, where every criterion holds,
    and how many more such offsets lie past every used cell.

    Past the used cells of every range, ``extra`` included, everything is
    blank, so those offsets are counted rather than read: a criterion over
    a whole column costs the rows that hold something.
    """
    first = pairs[0][0]
    for area, _ in pairs:
        if area.height != first.height or area.width != first.width:
            raise ExcelError(VALUE)
    height, width = bound(context, [area for area, _ in pairs] + extra)
    height, width = min(height, first.height), min(width, first.width)
    blocks = [block(context, area, height, width) for area, _ in pairs]
    found = [
        (row, column)
        for row in range(height)
        for column in range(width)
        if all(test.matches(values[row][column], context) for values, (_, test) in zip(blocks, pairs, strict=True))
    ]
    beyond = 0
    if all(test.matches(EMPTY, context) for _, test in pairs):
        beyond = first.height * first.width - height * width
    return found, beyond


__all__ = [
    "Criterion",
    "area_of",
    "block",
    "bound",
    "criterion",
    "flat",
    "matching",
    "matrix",
    "numbers",
    "pairs",
    "shaped",
]
