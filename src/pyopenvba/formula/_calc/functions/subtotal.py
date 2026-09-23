"""SUBTOTAL and AGGREGATE: a statistic over ranges, leaving out what a
list's own totals and filters would count twice.

Both leave out a cell whose formula calls SUBTOTAL or AGGREGATE, so a
total of a column holding subtotals counts each value once; AGGREGATE does
so only for options 0 to 3. Rows a filter hides are always left out of
SUBTOTAL, and rows hidden by hand only for function numbers 101 to 111.
AGGREGATE's options decide the rest: 1, 3, 5 and 7 leave out hidden rows,
2, 3, 6 and 7 skip errors rather than stopping at one. Only rows count as
hidden: a hidden column changes nothing.
"""

from __future__ import annotations

import math
from collections.abc import Callable

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.aggregate import spread
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.statistics import (
    percentile_exclusive,
    percentile_inclusive,
    quartile_exclusive,
    quartile_inclusive,
)
from pyopenvba.formula._calc.numbers import total
from pyopenvba.formula._calc.registry import A, R, V, function
from pyopenvba.formula._calc.values import (
    DIV0,
    NA,
    NUM,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Reference,
    Scalar,
    Value,
)
from pyopenvba.formula._calc.cells import CellError

Statistic = Callable[[list[Scalar]], Value]


def _numbers(values: list[Scalar]) -> list[float]:
    """The numbers among the values, stopping at the first error."""
    found: list[float] = []
    for value in values:
        if isinstance(value, CellError):
            raise ExcelError(value)
        if isinstance(value, float):
            found.append(value)
    return found


def _average(values: list[Scalar]) -> Value:
    numbers = _numbers(values)
    return checked(total(numbers) / len(numbers)) if numbers else DIV0


def _count(values: list[Scalar]) -> Value:
    return float(sum(1 for value in values if isinstance(value, float)))


def _count_all(values: list[Scalar]) -> Value:
    return float(sum(1 for value in values if not isinstance(value, Empty)))


def _maximum(values: list[Scalar]) -> Value:
    return max(_numbers(values), default=0.0)


def _minimum(values: list[Scalar]) -> Value:
    return min(_numbers(values), default=0.0)


def _product(values: list[Scalar]) -> Value:
    numbers = _numbers(values)
    if not numbers:
        return 0.0
    result = 1.0
    for number in numbers:
        result *= number
    return checked(result)


def _sum(values: list[Scalar]) -> Value:
    return checked(total(_numbers(values)))


def _variance(*, sample: bool, root: bool) -> Statistic:
    def statistic(values: list[Scalar]) -> Value:
        found = spread(_numbers(values), sample=sample)
        if isinstance(found, CellError):
            return found
        return checked(math.sqrt(found) if root else found)

    return statistic


def _median(values: list[Scalar]) -> Value:
    numbers = sorted(_numbers(values))
    if not numbers:
        return NUM
    middle = len(numbers) // 2
    if len(numbers) % 2:
        return numbers[middle]
    return (numbers[middle - 1] + numbers[middle]) / 2


def _mode(values: list[Scalar]) -> Value:
    numbers = _numbers(values)
    counts: dict[float, int] = {}
    for number in numbers:
        counts[number] = counts.get(number, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return NA
    return next(number for number in numbers if counts[number] == best)


#: SUBTOTAL's functions, 1 to 11, and AGGREGATE's first thirteen.
_STATISTICS: dict[int, Statistic] = {
    1: _average,
    2: _count,
    3: _count_all,
    4: _maximum,
    5: _minimum,
    6: _product,
    7: _variance(sample=True, root=True),
    8: _variance(sample=False, root=True),
    9: _sum,
    10: _variance(sample=True, root=False),
    11: _variance(sample=False, root=False),
    12: _median,
    13: _mode,
}


def _visible(
    context: Context,
    reference: Reference,
    *,
    hidden: bool,
    filtered: bool,
    nested: bool,
    errors: bool,
) -> list[Scalar]:
    """The values of a reference's cells, leaving out rows hidden or
    filtered, cells that are subtotals themselves, and errors, as asked."""
    book = context.book
    found: list[Scalar] = []
    for area in reference.areas:
        for row, column, value in book.cells(area):
            if hidden and book.row_hidden(area.sheet, row):
                continue
            if filtered and book.row_filtered(area.sheet, row):
                continue
            if nested and book.subtotal(area.sheet, row, column):
                continue
            if errors and isinstance(value, CellError):
                continue
            found.append(value)
    return found


@function("SUBTOTAL", V, R, maximum=255)
def SUBTOTAL(context: Context, which: Scalar, *refs: Value) -> Value:
    number = context.integer(which)
    hidden = number > 100
    statistic = _STATISTICS.get(number - 100 if hidden else number)
    if statistic is None or not 1 <= (number - 100 if hidden else number) <= 11:
        return VALUE
    values: list[Scalar] = []
    for ref in refs:
        if not isinstance(ref, Reference):
            return VALUE
        values.extend(_visible(context, ref, hidden=hidden, filtered=True, nested=True, errors=False))
    return statistic(values)


#: AGGREGATE's functions of an array and a k, 14 to 19.
_WITH_K: dict[int, Callable[[list[float], float], float]] = {
    14: lambda values, k: _kth(values, k, largest=True),
    15: lambda values, k: _kth(values, k, largest=False),
    16: percentile_inclusive,
    17: quartile_inclusive,
    18: percentile_exclusive,
    19: quartile_exclusive,
}


def _kth(values: list[float], k: float, *, largest: bool) -> float:
    count = len(values)
    if k < 1 or k > count:
        raise ExcelError(NUM)
    position = math.floor(count - k + 1) if largest else math.floor(k)
    return values[position - 1]


@function("AGGREGATE", V, V, A, maximum=255)
def AGGREGATE(context: Context, which: Scalar, options: Scalar, *args: Value) -> Value:
    number = context.integer(which)
    option = 0 if isinstance(options, Empty) else context.integer(options)
    if not 0 <= option <= 7 or not args:
        return VALUE
    hidden = option in (1, 3, 5, 7)
    errors = option in (2, 3, 6, 7)
    nested = option <= 3
    if number in _WITH_K:
        if len(args) != 2:
            return VALUE
        array, k = args
        values = _gathered(context, array, hidden=hidden, nested=nested, errors=errors)
        wanted = context.number(context.first(k))
        return _WITH_K[number](sorted(_numbers(values)), wanted)
    statistic = _STATISTICS.get(number)
    if statistic is None:
        return VALUE
    gathered: list[Scalar] = []
    for ref in args:
        if not isinstance(ref, Reference):
            return VALUE
        gathered.extend(_visible(context, ref, hidden=hidden, filtered=False, nested=nested, errors=errors))
    return statistic(gathered)


def _gathered(context: Context, value: Value, *, hidden: bool, nested: bool, errors: bool) -> list[Scalar]:
    """The values of AGGREGATE's array argument: a reference as its cells,
    with rows and subtotals left out as asked; an array as it is."""
    if isinstance(value, Reference):
        return _visible(context, value, hidden=hidden, filtered=False, nested=nested, errors=errors)
    if isinstance(value, Array):
        return [item for item in value.items() if not (errors and isinstance(item, CellError))]
    if isinstance(value, CellError):
        if errors:
            return []
        raise ExcelError(value)
    return [value]


__all__: list[str] = []
