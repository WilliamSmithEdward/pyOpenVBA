"""Counting, averages, extremes and spread.

These read their arguments as SUM does (see :mod:`.common`): numbers in a
range, and anything that can be a number when typed as an argument. COUNT
never fails on what it cannot count, and the functions ending in A read
text in a range as 0 and a logical as 1 or 0.
"""

from __future__ import annotations

import math
from fractions import Fraction

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import matrix, numbers
from pyopenvba.formula._calc.numbers import total
from pyopenvba.formula._calc.precise import (
    add,
    binary_log,
    binary_power,
    divide,
    extended,
    multiply,
    register,
    square_root,
    subtract,
    summed,
)
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import (
    DIV0,
    NA,
    NUM,
    Array,
    Empty,
    Reference,
    Scalar,
    Value,
    compare,
    text_to_number,
)
from pyopenvba.formula._calc.cells import CellError


@function("COUNT", R, maximum=255)
def COUNT(context: Context, *args: Value) -> Value:
    count = 0
    for arg in args:
        if isinstance(arg, (Reference, Array)):
            count += sum(1 for value, _ in context.scalars(arg) if isinstance(value, float))
        elif isinstance(arg, (float, bool, Empty)) or (isinstance(arg, str) and text_to_number(arg, context.today, epoch_1904=context.epoch_1904) is not None):
            count += 1
    return float(count)


@function("COUNTA", R, maximum=255)
def COUNTA(context: Context, *args: Value) -> Value:
    count = 0
    for arg in args:
        if isinstance(arg, (Reference, Array)):
            count += sum(1 for value, _ in context.scalars(arg) if not isinstance(value, Empty))
        else:
            count += 1
    return float(count)


@function("COUNTBLANK", R)
def COUNTBLANK(context: Context, range_: Value) -> Value:
    if isinstance(range_, Reference):
        cells = sum(area.height * area.width for area in range_.areas)
        filled = sum(1 for value, _ in context.scalars(range_) if value != "")
        return float(cells - filled)
    return float(sum(1 for value in matrix(context, range_).items() if isinstance(value, Empty) or value == ""))


@function("AVERAGE", R, maximum=255)
def AVERAGE(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values:
        return DIV0
    return checked(divide(total(values), float(len(values))))


@function("AVERAGEA", R, maximum=255)
def AVERAGEA(context: Context, *args: Value) -> Value:
    values = numbers(context, args, logicals=True, text=True)
    if not values:
        return DIV0
    return checked(divide(total(values), float(len(values))))


@function("MIN", R, maximum=255)
def MIN(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    return min(values) if values else 0.0


@function("MAX", R, maximum=255)
def MAX(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    return max(values) if values else 0.0


@function("MINA", R, maximum=255)
def MINA(context: Context, *args: Value) -> Value:
    values = numbers(context, args, logicals=True, text=True)
    return min(values) if values else 0.0


@function("MAXA", R, maximum=255)
def MAXA(context: Context, *args: Value) -> Value:
    values = numbers(context, args, logicals=True, text=True)
    return max(values) if values else 0.0


@function("MEDIAN", R, maximum=255)
def MEDIAN(context: Context, *args: Value) -> Value:
    values = sorted(numbers(context, args))
    if not values:
        return NUM
    middle = len(values) // 2
    if len(values) % 2:
        return values[middle]
    return divide(add(values[middle - 1], values[middle]), 2.0)


def _mode(values: list[float]) -> float | None:
    counts: dict[float, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return None
    # The first value, in the order given, to reach the highest count.
    return next(value for value in values if counts[value] == best)


@function("MODE", R, maximum=255)
def MODE(context: Context, *args: Value) -> Value:
    found = _mode(numbers(context, args))
    return NA if found is None else found


@function("MODE.SNGL", R, maximum=255)
def MODE_SNGL(context: Context, *args: Value) -> Value:
    return MODE(context, *args)


def _kth(context: Context, array: Value, k: Scalar, *, largest: bool) -> Value:
    """The k-th smallest, or largest, value. Measured: SMALL cuts k down to
    a whole number and LARGE takes the ``n-k+1``-th smallest the same way,
    so ``SMALL(x,2.1)`` is the second and ``LARGE(x,1.5)`` of four values
    the third smallest; k below 1 is ``#NUM!``."""
    values = sorted(numbers(context, (array,)))
    wanted = context.number(k)
    count = len(values)
    if wanted < 1 or wanted > count:
        return NUM
    position = math.floor(count - wanted + 1) if largest else math.floor(wanted)
    return values[position - 1]


@function("LARGE", R, V)
def LARGE(context: Context, array: Value, k: Scalar) -> Value:
    return _kth(context, array, k, largest=True)


@function("SMALL", R, V)
def SMALL(context: Context, array: Value, k: Scalar) -> Value:
    return _kth(context, array, k, largest=False)


def mean(values: list[float]) -> float:
    """The values' sum over their count, as the statistics take it."""
    return divide(summed(values), float(len(values)))


def deviations(values: list[float]) -> float:
    """The sum of squared deviations from the mean, DEVSQ: two passes."""
    centre = mean(values)
    return summed(multiply(subtract(value, centre), subtract(value, centre)) for value in values)


#: When the one-pass formula has cancelled away this share of the sum of
#: squares, or leaves a variance under this multiple of the mean, Excel
#: goes back over the values: measured, 2100 samples on either side.
_CANCELLED = 0.01
_SMALL_SPREAD = 0.001


def spread(values: list[float], *, sample: bool) -> float | CellError:
    """The variance of the values, of a sample or of the whole population.

    Measured on 2100 samples: from the sums in one pass,
    ``(Σx² - (Σx)²/n) / (n - 1)`` for a sample and ``(nΣx² - (Σx)²) / n²``
    for the population, unless that cancelled away all but a hundredth of
    ``Σx²`` or leaves a variance under a thousandth of the mean's size;
    then the squared deviations from the mean, over n - 1 or n."""
    count = len(values)
    if count == 0 or (sample and count < 2):
        return DIV0
    n = float(count)
    plain = summed(values)
    squares = summed(multiply(value, value) for value in values)
    if sample:
        left = subtract(squares, divide(multiply(plain, plain), n))
        whole = squares
        found = divide(left, subtract(n, 1.0))
    else:
        left = subtract(multiply(n, squares), multiply(plain, plain))
        whole = multiply(n, squares)
        found = divide(left, multiply(n, n))
    centre = divide(plain, n)
    if left < multiply(_CANCELLED, whole) or (centre != 0 and divide(found, abs(centre)) < _SMALL_SPREAD):
        return divide(deviations(values), subtract(n, 1.0) if sample else n)
    return found


def _variance(context: Context, args: tuple[Value, ...], *, sample: bool, text: bool = False) -> Value:
    found = spread(numbers(context, args, logicals=text, text=text), sample=sample)
    return found if isinstance(found, CellError) else checked(found)


def _deviation(context: Context, args: tuple[Value, ...], *, sample: bool, text: bool = False) -> Value:
    found = spread(numbers(context, args, logicals=text, text=text), sample=sample)
    return found if isinstance(found, CellError) else checked(square_root(found))


@function("VAR", R, maximum=255)
def VAR(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=True)


@function("VAR.S", R, maximum=255)
def VAR_S(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=True)


@function("VARA", R, maximum=255)
def VARA(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=True, text=True)


@function("VARP", R, maximum=255)
def VARP(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=False)


@function("VAR.P", R, maximum=255)
def VAR_P(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=False)


@function("VARPA", R, maximum=255)
def VARPA(context: Context, *args: Value) -> Value:
    return _variance(context, args, sample=False, text=True)


@function("STDEV", R, maximum=255)
def STDEV(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=True)


@function("STDEV.S", R, maximum=255)
def STDEV_S(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=True)


@function("STDEVA", R, maximum=255)
def STDEVA(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=True, text=True)


@function("STDEVP", R, maximum=255)
def STDEVP(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=False)


@function("STDEV.P", R, maximum=255)
def STDEV_P(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=False)


@function("STDEVPA", R, maximum=255)
def STDEVPA(context: Context, *args: Value) -> Value:
    return _deviation(context, args, sample=False, text=True)


@function("AVEDEV", R, maximum=255)
def AVEDEV(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values:
        return NUM
    centre = mean(values)
    return checked(divide(summed(abs(subtract(value, centre)) for value in values), float(len(values))))


@function("DEVSQ", R, maximum=255)
def DEVSQ(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values:
        return NUM
    return checked(deviations(values))


@function("GEOMEAN", R, maximum=255)
def GEOMEAN(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values or any(value <= 0 for value in values):
        return NUM
    # The base-2 logarithms summed in an x87 register, their mean there,
    # and 2 to that: Excel's last bit on 239 of 276 samples, where no
    # formula on doubles reaches 110. Excel strays a unit from this on the
    # others, not at ties, some way not yet found.
    total = Fraction(0)
    for value in values:
        total = register(total + register(binary_log(value)))
    return checked(extended(binary_power(register(total / len(values)))))


@function("PERCENTOF", R, R)
def PERCENTOF(context: Context, subset: Value, whole: Value) -> Value:
    """The share a subset's sum is of a whole's."""
    part = total(numbers(context, (subset,)))
    everything = total(numbers(context, (whole,)))
    if everything == 0:
        return DIV0
    return checked(divide(part, everything))


@function("HARMEAN", R, maximum=255)
def HARMEAN(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values or any(value <= 0 for value in values):
        return NUM
    # Measured on 276 samples: one over the mean of the reciprocals, not n
    # over their sum.
    return checked(divide(1.0, divide(summed(divide(1.0, value) for value in values), float(len(values)))))


def _rank(context: Context, number: Scalar, ref: Value, order: Scalar | None, *, average: bool) -> Value:
    value = context.number(number)
    values = numbers(context, (ref,))
    ascending = order is not None and context.number(order) != 0
    if not any(compare(item, value) == 0 for item in values):
        return NA
    before = sum(1 for item in values if (compare(item, value) < 0 if ascending else compare(item, value) > 0))
    if not average:
        return float(before + 1)
    ties = sum(1 for item in values if compare(item, value) == 0)
    return before + (ties + 1) / 2


@function("RANK", V, R, V, minimum=2)
def RANK(context: Context, number: Scalar, ref: Value, order: Scalar | None = None) -> Value:
    return _rank(context, number, ref, order, average=False)


@function("RANK.EQ", V, R, V, minimum=2)
def RANK_EQ(context: Context, number: Scalar, ref: Value, order: Scalar | None = None) -> Value:
    return _rank(context, number, ref, order, average=False)


@function("RANK.AVG", V, R, V, minimum=2)
def RANK_AVG(context: Context, number: Scalar, ref: Value, order: Scalar | None = None) -> Value:
    return _rank(context, number, ref, order, average=True)


__all__ = ["spread"]
