"""Math and trigonometry functions.

The rounding functions work on a number's fifteen significant digits, not
its exact binary value: ``ROUND(1.005,2)`` is 1.01 although the double
nearest 1.005 is a little below it, and ``ROUND(0.285,2)`` is 0.29. INT
does too: ``INT(2.9999999999999996)`` is 3.

SIN, COS and TAN reduce their argument with pi to 64 bits, as the x87
floating-point unit holds it, and round the result to 64 bits and then to
a double, which is why ``SIN(PI())`` is ``1.22514845490862E-16`` in Excel
rather than the ``1.2246...E-16`` of the C library. An argument of 2^27 or
more is ``#NUM!``. See :mod:`~pyopenvba.formula._calc.precise`.
"""

from __future__ import annotations

import math
import random
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_HALF_UP, ROUND_UP, Decimal
from decimal import Context as DecimalContext
from fractions import Fraction

from pyopenvba.formula._calc import precise
from pyopenvba.formula._calc.evaluator import Context, power
from pyopenvba.formula._calc.functions.common import matrix, numbers, pairs
from pyopenvba.formula._calc.numbers import normal, total
from pyopenvba.formula._calc.registry import A, R, V, function
from pyopenvba.formula._calc.values import DIV0, NUM, VALUE, Array, ExcelError, Scalar, Value
from pyopenvba.formula._calc.cells import CellError

#: Enough digits for any double written out in full.
_WIDE = DecimalContext(prec=800)
_TRIG_LIMIT = 2.0**27


def checked(value: float) -> float:
    """A result as Excel keeps it: no infinities, no subnormals, no -0."""
    if not math.isfinite(value):
        raise ExcelError(NUM)
    return normal(value) + 0.0


def fifteen(value: float) -> Decimal:
    """A number as its fifteen significant digits."""
    if value == 0.0:
        return Decimal(0)
    exact = Decimal(value)
    return exact.quantize(Decimal(1).scaleb(exact.adjusted() - 14), rounding=ROUND_HALF_UP, context=_WIDE)


def rounded(value: float, digits: int, rounding: str) -> float:
    """``value`` rounded at ``digits`` places after the point, negative for
    places before it, reading the value to fifteen digits first.

    Measured: ROUND, ROUNDDOWN and TRUNC round those fifteen digits, so
    rounding past them gives the fifteen-digit value itself. ROUNDUP rounds
    down and then adds one step in floating point when anything was cut
    off, which is why ``ROUNDUP(2.675,2)`` is ``2.6799999999999997``: 2.67
    plus 0.01.
    """
    if value == 0.0:
        return 0.0
    decimal = fifteen(value)
    places = max(min(digits, 400), -400)
    quantum = Decimal(1).scaleb(-places)
    if rounding == ROUND_UP:
        down = decimal.quantize(quantum, rounding=ROUND_DOWN, context=_WIDE)
        if down == decimal:
            return checked(float(down))
        try:
            step = 10.0**-digits
        except OverflowError:
            raise ExcelError(NUM) from None
        return checked(float(down) + math.copysign(step, value))
    return checked(float(decimal.quantize(quantum, rounding=rounding, context=_WIDE)))


@function("ABS", V)
def ABS(context: Context, number: Scalar) -> Value:
    return abs(context.number(number))


@function("SIGN", V)
def SIGN(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    return 1.0 if value > 0 else -1.0 if value < 0 else 0.0


@function("SQRT", V)
def SQRT(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value < 0:
        return NUM
    return precise.square_root(value)


@function("SQRTPI", V)
def SQRTPI(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value < 0:
        return NUM
    return checked(precise.square_root(precise.multiply(value, math.pi)))


@function("POWER", V, V)
def POWER(context: Context, number: Scalar, exponent: Scalar) -> Value:
    return power(context.number(number), context.number(exponent))


@function("EXP", V)
def EXP(context: Context, number: Scalar) -> Value:
    return checked(precise.exp(context.number(number)))


@function("LN", V)
def LN(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value <= 0:
        return NUM
    return precise.ln(value)


@function("LOG10", V)
def LOG10(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value <= 0:
        return NUM
    return precise.log10(value)


@function("LOG", V, V, minimum=1)
def LOG(context: Context, number: Scalar, base: Scalar | None = None) -> Value:
    """LOG10 with one argument; with a base, measured on 800 pairs, the
    quotient of the two natural logarithms."""
    value = context.number(number)
    if base is None:
        if value <= 0:
            return NUM
        return precise.log10(value)
    radix = context.number(base)
    if value <= 0 or radix <= 0:
        return NUM
    if radix == 1:
        return DIV0
    return checked(precise.divide(precise.ln(value), precise.ln(radix)))


@function("PI")
def PI(context: Context) -> Value:
    return math.pi


def _digits(context: Context, value: Scalar | None) -> int:
    return 0 if value is None else context.integer(value)


@function("ROUND", V, V)
def ROUND(context: Context, number: Scalar, digits: Scalar) -> Value:
    return rounded(context.number(number), _digits(context, digits), ROUND_HALF_UP)


@function("ROUNDUP", V, V)
def ROUNDUP(context: Context, number: Scalar, digits: Scalar) -> Value:
    return rounded(context.number(number), _digits(context, digits), ROUND_UP)


@function("ROUNDDOWN", V, V)
def ROUNDDOWN(context: Context, number: Scalar, digits: Scalar) -> Value:
    return rounded(context.number(number), _digits(context, digits), ROUND_DOWN)


@function("TRUNC", V, V, minimum=1)
def TRUNC(context: Context, number: Scalar, digits: Scalar | None = None) -> Value:
    return rounded(context.number(number), _digits(context, digits), ROUND_DOWN)


@function("INT", V)
def INT(context: Context, number: Scalar) -> Value:
    """The whole number at or below, except that a number reading as the
    next whole number up to fifteen digits is that number: INT(2.9999999999999996)
    is 3, while INT(4503599627370495.5) is 4503599627370495, not the
    fifteen-digit ...500."""
    value = context.number(number)
    floor = math.floor(value)
    if floor != value and fifteen(value) == math.ceil(value):
        return float(math.ceil(value)) + 0.0
    return float(floor) + 0.0


#: The largest quotient MOD takes: measured between 2^38 and 2^47.3.
_MOD_LIMIT = 2.0**47


@function("MOD", V, V)
def MOD(context: Context, number: Scalar, divisor: Scalar) -> Value:
    value = context.number(number)
    by = context.number(divisor)
    if by == 0:
        return DIV0
    if abs(value / by) >= _MOD_LIMIT:
        return NUM
    # The exact remainder of the two doubles, with the divisor's sign:
    # MOD(10,3.3) is 0.10000000000000053, not 10-3*3.3.
    return checked(value % by)


@function("QUOTIENT", V, V)
def QUOTIENT(context: Context, numerator: Scalar, denominator: Scalar) -> Value:
    top = context.number(numerator)
    bottom = context.number(denominator)
    if bottom == 0:
        return DIV0
    return float(math.trunc(top / bottom)) + 0.0


def _multiple(value: float, significance: float, rounding: str) -> float:
    """``value`` rounded to a multiple of ``significance``."""
    if significance == 0:
        return 0.0
    ratio = fifteen(value / significance)
    return checked(float(ratio.quantize(Decimal(1), rounding=rounding, context=_WIDE)) * significance)


@function("CEILING", V, V)
def CEILING(context: Context, number: Scalar, significance: Scalar) -> Value:
    value = context.number(number)
    step = context.number(significance)
    if value > 0 and step < 0:
        return NUM
    # Up in multiples of the step: with both negative the ratio is positive,
    # so -4.3 by -1 is -5, while -4.3 by 1 is -4.
    return 0.0 if step == 0 else _multiple(value, step, ROUND_CEILING)


@function("FLOOR", V, V)
def FLOOR(context: Context, number: Scalar, significance: Scalar) -> Value:
    value = context.number(number)
    step = context.number(significance)
    if value > 0 and step < 0:
        return NUM
    if step == 0:
        return DIV0 if value != 0 else 0.0
    return _multiple(value, step, ROUND_FLOOR)


@function("CEILING.MATH", V, V, V, minimum=1)
def CEILING_MATH(context: Context, number: Scalar, significance: Scalar | None = None, mode: Scalar | None = None) -> Value:
    value = context.number(number)
    step = 1.0 if significance is None else abs(context.number(significance))
    away = mode is not None and context.number(mode) != 0
    if step == 0:
        return 0.0
    return _multiple(value, step, ROUND_UP if value < 0 and away else ROUND_CEILING)


@function("FLOOR.MATH", V, V, V, minimum=1)
def FLOOR_MATH(context: Context, number: Scalar, significance: Scalar | None = None, mode: Scalar | None = None) -> Value:
    value = context.number(number)
    step = 1.0 if significance is None else abs(context.number(significance))
    toward = mode is not None and context.number(mode) != 0
    if step == 0:
        return 0.0
    return _multiple(value, step, ROUND_DOWN if value < 0 and toward else ROUND_FLOOR)


@function("CEILING.PRECISE", V, V, minimum=1)
def CEILING_PRECISE(context: Context, number: Scalar, significance: Scalar | None = None) -> Value:
    value = context.number(number)
    step = 1.0 if significance is None else abs(context.number(significance))
    return 0.0 if step == 0 else _multiple(value, step, ROUND_CEILING)


@function("ISO.CEILING", V, V, minimum=1)
def ISO_CEILING(context: Context, number: Scalar, significance: Scalar | None = None) -> Value:
    return CEILING_PRECISE(context, number, significance)


# The CEILING of the ECMA-376 standard, which Excel's own CEILING follows.
function("ECMA.CEILING", V, V)(CEILING)


@function("FLOOR.PRECISE", V, V, minimum=1)
def FLOOR_PRECISE(context: Context, number: Scalar, significance: Scalar | None = None) -> Value:
    value = context.number(number)
    step = 1.0 if significance is None else abs(context.number(significance))
    return 0.0 if step == 0 else _multiple(value, step, ROUND_FLOOR)


@function("MROUND", V, V)
def MROUND(context: Context, number: Scalar, multiple: Scalar) -> Value:
    value = context.number(number)
    step = context.number(multiple)
    if step == 0:
        return 0.0
    if (value > 0 and step < 0) or (value < 0 and step > 0):
        return NUM
    # The quotient as a double, not to fifteen digits: MROUND(2^53,3) is
    # 2^53, the nearest multiple as a double, not 2^53-2.
    ratio = Decimal(value / step).quantize(Decimal(1), rounding=ROUND_HALF_UP, context=_WIDE)
    return checked(float(ratio) * step)


@function("EVEN", V)
def EVEN(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    magnitude = math.ceil(abs(value) / 2) * 2
    return math.copysign(float(magnitude), value) + 0.0


@function("ODD", V)
def ODD(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    magnitude = math.ceil(abs(value))
    if magnitude % 2 == 0:
        magnitude += 1
    return math.copysign(float(magnitude), value) + 0.0


@function("FACT", V)
def FACT(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value < 0:
        return NUM
    whole = int(value)
    if whole > 170:
        return NUM
    return float(math.factorial(whole))


@function("FACTDOUBLE", V)
def FACTDOUBLE(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value < -1:
        return NUM
    whole = int(value)
    result = 1.0
    while whole > 1:
        result *= whole
        whole -= 2
    return checked(result)


@function("COMBIN", V, V)
def COMBIN(context: Context, number: Scalar, chosen: Scalar) -> Value:
    n = int(context.number(number))
    k = int(context.number(chosen))
    if n < 0 or k < 0 or n < k:
        return NUM
    return checked(float(math.comb(n, k)))


@function("PERMUT", V, V)
def PERMUT(context: Context, number: Scalar, chosen: Scalar) -> Value:
    n = int(context.number(number))
    k = int(context.number(chosen))
    if n < 0 or k < 0 or n < k:
        return NUM
    return checked(float(math.perm(n, k)))


def _whole_numbers(context: Context, args: tuple[Value, ...]) -> list[int]:
    found: list[int] = []
    for value in numbers(context, args):
        if value < 0 or value >= 2.0**53:
            raise ExcelError(NUM)
        found.append(int(value))
    return found


@function("GCD", R, maximum=255)
def GCD(context: Context, *args: Value) -> Value:
    return float(math.gcd(*_whole_numbers(context, args)))


@function("LCM", R, maximum=255)
def LCM(context: Context, *args: Value) -> Value:
    values = _whole_numbers(context, args)
    result = 1
    for value in values:
        if value == 0:
            return 0.0
        result = result * value // math.gcd(result, value)
    return checked(float(result))


@function("SUM", R, maximum=255)
def SUM(context: Context, *args: Value) -> Value:
    return checked(total(numbers(context, args)))


@function("SUMSQ", R, maximum=255)
def SUMSQ(context: Context, *args: Value) -> Value:
    return checked(precise.summed(precise.multiply(value, value) for value in numbers(context, args)))


@function("PRODUCT", R, maximum=255)
def PRODUCT(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values:
        return 0.0
    result = 1.0
    for value in values:
        result = precise.multiply(result, value)
    return checked(result)


@function("SUMPRODUCT", A, maximum=255)
def SUMPRODUCT(context: Context, *args: Value) -> Value:
    arrays: list[Array] = [matrix(context, arg) for arg in args]
    first = arrays[0]
    if any(array.height != first.height or array.width != first.width for array in arrays):
        return VALUE
    total = 0.0
    for row in range(first.height):
        for column in range(first.width):
            product = 1.0
            for array in arrays:
                item = array.rows[row][column]
                if isinstance(item, CellError):
                    return item
                product = precise.multiply(product, item if isinstance(item, float) else 0.0)
            total = precise.add(total, product)
    return checked(total)


# Measured on 300 pairs of ranges each: the sum of each position's term, in
# order, not a difference or sum of two sums.


@function("SUMX2MY2", A, A)
def SUMX2MY2(context: Context, first: Value, second: Value) -> Value:
    found = pairs(context, first, second)
    return checked(precise.summed(precise.subtract(precise.multiply(x, x), precise.multiply(y, y)) for x, y in found))


@function("SUMX2PY2", A, A)
def SUMX2PY2(context: Context, first: Value, second: Value) -> Value:
    found = pairs(context, first, second)
    return checked(precise.summed(precise.add(precise.multiply(x, x), precise.multiply(y, y)) for x, y in found))


@function("SUMXMY2", A, A)
def SUMXMY2(context: Context, first: Value, second: Value) -> Value:
    found = pairs(context, first, second)
    return checked(precise.summed(precise.multiply(precise.subtract(x, y), precise.subtract(x, y)) for x, y in found))


# ----------------------------------------------------------------------
# Trigonometry
# ----------------------------------------------------------------------


def _angle(context: Context, number: Scalar) -> float:
    value = context.number(number)
    if abs(value) >= _TRIG_LIMIT:
        raise ExcelError(NUM)
    return value


@function("SIN", V)
def SIN(context: Context, number: Scalar) -> Value:
    return checked(precise.sin(_angle(context, number)))


@function("COS", V)
def COS(context: Context, number: Scalar) -> Value:
    return checked(precise.cos(_angle(context, number)))


@function("TAN", V)
def TAN(context: Context, number: Scalar) -> Value:
    sine, cosine = precise.sine_cosine(_angle(context, number))
    if cosine == 0:
        return DIV0
    return checked(precise.extended(sine / cosine))


def _reciprocal(value: Value) -> Value:
    """One over a result, as COT, CSC, SEC and their hyperbolic kin take
    it: measured, of TAN's, SIN's or COS's double, not of the exact value."""
    if not isinstance(value, float):
        return value
    if value == 0.0:
        return DIV0
    return checked(precise.divide(1.0, value))


@function("COT", V)
def COT(context: Context, number: Scalar) -> Value:
    if context.number(number) == 0:
        return DIV0
    return _reciprocal(TAN(context, number))


@function("CSC", V)
def CSC(context: Context, number: Scalar) -> Value:
    if context.number(number) == 0:
        return DIV0
    return _reciprocal(SIN(context, number))


@function("SEC", V)
def SEC(context: Context, number: Scalar) -> Value:
    return _reciprocal(COS(context, number))


@function("COTH", V)
def COTH(context: Context, number: Scalar) -> Value:
    if context.number(number) == 0:
        return DIV0
    return _reciprocal(TANH(context, number))


@function("CSCH", V)
def CSCH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value == 0:
        return DIV0
    if abs(value) > 710:
        # Measured: 0 where SINH overflows, as SECH.
        return 0.0
    return _reciprocal(SINH(context, number))


@function("SECH", V)
def SECH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if abs(value) > 710:
        return 0.0
    return _reciprocal(COSH(context, number))


def arctangent(value: float) -> float:
    """ATAN as Excel computes it: the exact arctangent, rounded to 64 bits
    and then to a double."""
    return precise.extended(Fraction(precise.atan(value)))


@function("ATAN", V)
def ATAN(context: Context, number: Scalar) -> Value:
    return arctangent(context.number(number))


def _arcsine(value: float) -> float:
    """``atan(x / sqrt((1 - x)(1 + x)))``, each step as the x87 rounds it.
    Measured: ACOS is pi/2 less this on 1000 of 1000 arguments. ASIN agrees
    on 1953 of 2000; Excel's own is not odd below 0.35, ``ASIN(-x)`` being
    ``-ASIN(x)`` a unit off in 45 of 1000, which no odd formula follows."""
    if abs(value) == 1:
        return math.copysign(math.pi / 2, value)
    root = precise.square_root(precise.multiply(precise.subtract(1.0, value), precise.add(1.0, value)))
    return arctangent(precise.divide(value, root))


@function("ASIN", V)
def ASIN(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if not -1 <= value <= 1:
        return NUM
    return _arcsine(value)


@function("ACOS", V)
def ACOS(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if not -1 <= value <= 1:
        return NUM
    return precise.subtract(math.pi / 2, _arcsine(value))


@function("ATAN2", V, V)
def ATAN2(context: Context, x: Scalar, y: Scalar) -> Value:
    """The arctangent of y/x, a half turn added or taken away when x is
    negative: measured, 800 of 800."""
    across = context.number(x)
    up = context.number(y)
    if across == 0:
        if up == 0:
            return DIV0
        return math.copysign(math.pi / 2, up)
    angle = arctangent(precise.divide(up, across))
    if across < 0:
        angle = precise.add(angle, math.pi) if up >= 0 else precise.subtract(angle, math.pi)
    return angle


# SINH, COSH and TANH are EXP's (e^x - e^-x)/2, (e^x + e^-x)/2 and their
# ratio, each operation as the x87 rounds it: measured, to the bit on 1221
# arguments for COSH and for the others from 1 up. Below 1 Excel avoids the
# cancellation some way these differences of e^x - 1 approach, a unit in
# the last place off in about one argument in four.
_CANCELS = 1.0


def _hyperbolic_sine(value: float) -> float:
    if abs(value) < _CANCELS:
        return precise.divide(precise.subtract(precise.exp_less_one(value), precise.exp_less_one(-value)), 2.0)
    return precise.divide(precise.subtract(precise.exp(value), precise.exp(-value)), 2.0)


def _hyperbolic_cosine(value: float) -> float:
    return precise.divide(precise.add(precise.exp(value), precise.exp(-value)), 2.0)


@function("SINH", V)
def SINH(context: Context, number: Scalar) -> Value:
    return checked(_hyperbolic_sine(context.number(number)))


@function("COSH", V)
def COSH(context: Context, number: Scalar) -> Value:
    return checked(_hyperbolic_cosine(context.number(number)))


@function("TANH", V)
def TANH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if abs(value) >= 20:
        # tanh(20) is within a unit in the last place of 1.
        return math.copysign(1.0, value)
    if abs(value) < _CANCELS:
        # Measured: the sine over the cosine, which is the sine itself
        # while the cosine rounds to 1.
        return checked(precise.divide(_hyperbolic_sine(value), _hyperbolic_cosine(value)))
    grow, shrink = precise.exp(value), precise.exp(-value)
    return checked(precise.divide(precise.subtract(grow, shrink), precise.add(grow, shrink)))


# The inverse hyperbolic functions are their logarithms, each step as the
# x87 rounds it: ASINH ln(|x| + sqrt(x^2 + 1)) with x's sign, ACOSH
# ln(x + sqrt(x^2 - 1)), ATANH ln((1 + x) / (1 - x)) / 2, measured on 500
# arguments each. ASINH and ACOSH are #NUM! once x^2 overflows. Near zero
# ATANH, and ACOTH for large x, sum their series instead, each from its own
# measured switch.
#: ATANH sums x + x^3/3 + x^5/5 + ... up to this, the logarithm above.
_ATANH_SERIES = float.fromhex("0x1.af82b729c1d83p-14")
#: ACOTH sums 1/x + 1/(3x^3) + 1/(5x^5) + ... from this up.
_ACOTH_SERIES = 3.69662725


def _arctanh_series(value: float) -> float:
    """``x + x^3/3 + x^5/5 + ...``, until a term no longer changes the sum."""
    square = precise.multiply(value, value)
    total = power = value
    order = 1.0
    while True:
        order += 2.0
        power = precise.multiply(power, square)
        grown = precise.add(total, precise.divide(power, order))
        if grown == total:
            return total
        total = grown


def _arccoth_series(size: float) -> float:
    """``1/x + 1/(3x^3) + 1/(5x^5) + ...``, until a term no longer changes
    the sum."""
    square = precise.multiply(size, size)
    total = precise.divide(1.0, size)
    power = size
    order = 1.0
    while True:
        order += 2.0
        power = precise.multiply(power, square)
        if math.isinf(power):
            return total
        grown = precise.add(total, precise.divide(1.0, precise.multiply(order, power)))
        if grown == total:
            return total
        total = grown


@function("ASINH", V)
def ASINH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    size = abs(value)
    square = precise.multiply(size, size)
    if math.isinf(square):
        return NUM
    found = precise.ln(precise.add(size, precise.square_root(precise.add(square, 1.0))))
    return checked(math.copysign(found, value))


@function("ACOSH", V)
def ACOSH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if value < 1:
        return NUM
    square = precise.multiply(value, value)
    if math.isinf(square):
        return NUM
    return checked(precise.ln(precise.add(value, precise.square_root(precise.subtract(square, 1.0)))))


@function("ATANH", V)
def ATANH(context: Context, number: Scalar) -> Value:
    value = context.number(number)
    if not -1 < value < 1:
        return NUM
    if abs(value) <= _ATANH_SERIES:
        return checked(_arctanh_series(value))
    return checked(precise.multiply(0.5, precise.ln(precise.divide(precise.add(1.0, value), precise.subtract(1.0, value)))))


@function("ACOT", V)
def ACOT(context: Context, number: Scalar) -> Value:
    """Measured: the arctangent of 1/x, a half turn added below zero."""
    value = context.number(number)
    if value == 0:
        return math.pi / 2
    angle = arctangent(precise.divide(1.0, value))
    return checked(angle if value > 0 else precise.add(angle, math.pi))


@function("ACOTH", V)
def ACOTH(context: Context, number: Scalar) -> Value:
    """Half the logarithm of ``(|x| + 1) / (|x| - 1)`` below 3.69662725 and
    the series above, with x's sign: measured, 604 of 604 below and 521 of
    521 above, the switch found to the last double."""
    value = context.number(number)
    if abs(value) <= 1:
        return NUM
    size = abs(value)
    if size >= _ACOTH_SERIES:
        found = _arccoth_series(size)
    else:
        found = precise.multiply(0.5, precise.ln(precise.divide(precise.add(size, 1.0), precise.subtract(size, 1.0))))
    return checked(math.copysign(found, value))


@function("DEGREES", V)
def DEGREES(context: Context, angle: Scalar) -> Value:
    # One multiplication by the ratio, which is what matches Excel.
    return checked(precise.multiply(context.number(angle), 180 / math.pi))


@function("RADIANS", V)
def RADIANS(context: Context, angle: Scalar) -> Value:
    return checked(precise.multiply(context.number(angle), math.pi / 180))


@function("RAND", volatile=True)
def RAND(context: Context) -> Value:
    return random.random()


@function("RANDBETWEEN", V, V, volatile=True)
def RANDBETWEEN(context: Context, bottom: Scalar, top: Scalar) -> Value:
    low = math.ceil(context.number(bottom))
    high = math.floor(context.number(top))
    if low > high:
        return NUM
    return float(random.randint(low, high))


__all__ = ["checked", "fifteen", "rounded"]
