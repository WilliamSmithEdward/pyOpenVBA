"""Engineering functions: bases, bits, complex numbers, Bessel functions,
the error function, unit conversion, Roman numerals.

Base conversions work in ten digits of two's complement, as Excel's do: a
negative number in binary is ten bits, ``DEC2BIN(-100)`` is
``1110011100``, and one past the range is ``#NUM!``.

A complex number is text, ``"3+4i"``. Its parts print with the fifteen
digits every number has as text, so ``IMPOWER("2+3i",3)``, which Excel
works out in polar form, is ``-46+9.00000000000001i``.

The Bessel functions are the approximations of Abramowitz and Stegun and
of Hart's *Computer Approximations* in the form *Numerical Recipes* gives
them, with three of Excel's coefficients differing from the book's and
every operation rounded as the x87 rounds it. Excel's results are those to
the bit, all 1632 measured, which is how ``BESSELJ(1.9,2)`` comes to be
``0.3299258286697852`` where the exact value is ``0.32992572769238722``.
ERF and ERFC are the doubles nearest the exact values.
"""

from __future__ import annotations

import cmath
import math
import re
from collections.abc import Callable
from decimal import Decimal
from fractions import Fraction

from pyopenvba.formula._calc import precise, special
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import numbers
from pyopenvba.formula._calc.numbers import number_text
from pyopenvba.formula._calc.precise import add, divide, extended, multiply, square_root, subtract
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import DIV0, NA, NUM, VALUE, Array, Empty, ExcelError, Reference, Scalar, Value
from pyopenvba.formula._calc.cells import CellError

# ----------------------------------------------------------------------
# Bases
# ----------------------------------------------------------------------

#: The bits of each base's ten digits, the top one a sign.
_BITS = {2: 10, 8: 30, 16: 40}
_DIGITS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _digits_of(context: Context, value: Scalar, base: int) -> int:
    """A number written in a base, as a signed integer."""
    if isinstance(value, bool):
        raise ExcelError(VALUE)
    text = context.text(value).upper()
    if text == "":
        return 0
    if len(text) > 10 or any(_DIGITS.find(char) not in range(base) for char in text):
        raise ExcelError(NUM)
    number = int(text, base)
    bits = _BITS[base]
    if number >= 1 << (bits - 1):
        number -= 1 << bits
    return number


def _write(context: Context, number: int, base: int, places: Scalar | None) -> str:
    """A signed integer in a base, padded to ``places`` when given. A
    negative number is always ten digits and ignores ``places``."""
    bits = _BITS[base]
    if not -(1 << (bits - 1)) <= number < 1 << (bits - 1):
        raise ExcelError(NUM)
    if number < 0:
        return _in_base(number + (1 << bits), base)
    text = _in_base(number, base)
    if places is not None and not isinstance(places, Empty):
        width = context.integer(places)
        if width < len(text) or width > 10:
            raise ExcelError(NUM)
        text = text.rjust(width, "0")
    return text


def _in_base(number: int, base: int) -> str:
    out = ""
    while number:
        number, rest = divmod(number, base)
        out = _DIGITS[rest] + out
    return out or "0"


def _register_bases() -> None:
    names = {2: "BIN", 8: "OCT", 16: "HEX"}
    for base, name in names.items():

        def to_decimal(context: Context, value: Scalar, base: int = base) -> Value:
            return float(_digits_of(context, value, base))

        def from_decimal(context: Context, value: Scalar, places: Scalar | None = None, base: int = base) -> Value:
            return _write(context, math.trunc(context.number(value)), base, places)

        function(f"{name}2DEC", V)(to_decimal)
        function(f"DEC2{name}", V, V, minimum=1)(from_decimal)
        for other, other_name in names.items():
            if other != base:

                def across(
                    context: Context, value: Scalar, places: Scalar | None = None, base: int = base, other: int = other
                ) -> Value:
                    return _write(context, _digits_of(context, value, base), other, places)

                function(f"{name}2{other_name}", V, V, minimum=1)(across)


_register_bases()


@function("BASE", V, V, V, minimum=2)
def BASE(context: Context, number: Scalar, radix: Scalar, length: Scalar | None = None) -> Value:
    value = context.number(number)
    base = context.integer(radix)
    if value < 0 or value >= 2**53 or not 2 <= base <= 36:
        return NUM
    whole = int(value)
    out = ""
    while whole:
        whole, rest = divmod(whole, base)
        out = _DIGITS[rest] + out
    out = out or "0"
    if length is not None and not isinstance(length, Empty):
        width = context.integer(length)
        if not 0 <= width <= 255:
            return NUM
        out = out.rjust(width, "0")
    return out


@function("DECIMAL", V, V)
def DECIMAL(context: Context, text: Scalar, radix: Scalar) -> Value:
    source = context.text(text).upper()
    base = context.integer(radix)
    if not 2 <= base <= 36 or len(source) > 255:
        return NUM
    if any(_DIGITS.find(char) not in range(base) for char in source):
        return NUM
    return checked(float(int(source, base))) if source else 0.0


def _bits(context: Context, value: Scalar) -> int:
    number = context.number(value)
    if number < 0 or number >= 2**48 or not number.is_integer():
        raise ExcelError(NUM)
    return int(number)


@function("BITAND", V, V)
def BITAND(context: Context, first: Scalar, second: Scalar) -> Value:
    return float(_bits(context, first) & _bits(context, second))


@function("BITOR", V, V)
def BITOR(context: Context, first: Scalar, second: Scalar) -> Value:
    return float(_bits(context, first) | _bits(context, second))


@function("BITXOR", V, V)
def BITXOR(context: Context, first: Scalar, second: Scalar) -> Value:
    return float(_bits(context, first) ^ _bits(context, second))


def _shift(context: Context, number: Scalar, amount: Scalar, sign: int) -> Value:
    value = _bits(context, number)
    by = context.integer(amount) * sign
    if abs(by) > 53:
        return NUM
    result = value << by if by >= 0 else value >> -by
    if result >= 2**48:
        return NUM
    return float(result)


@function("BITLSHIFT", V, V)
def BITLSHIFT(context: Context, number: Scalar, amount: Scalar) -> Value:
    return _shift(context, number, amount, 1)


@function("BITRSHIFT", V, V)
def BITRSHIFT(context: Context, number: Scalar, amount: Scalar) -> Value:
    return _shift(context, number, amount, -1)


@function("DELTA", V, V, minimum=1)
def DELTA(context: Context, first: Scalar, second: Scalar | None = None) -> Value:
    other = 0.0 if second is None else context.number(second)
    return 1.0 if context.number(first) == other else 0.0


@function("GESTEP", V, V, minimum=1)
def GESTEP(context: Context, number: Scalar, step: Scalar | None = None) -> Value:
    threshold = 0.0 if step is None else context.number(step)
    return 1.0 if context.number(number) >= threshold else 0.0


# ----------------------------------------------------------------------
# Roman numerals
# ----------------------------------------------------------------------

_CLASSIC = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]  # fmt: skip

#: The subtractions each form of ROMAN adds to the classic ones, from the
#: classic 0 to the most concise 4. ``ROMAN(499,1)`` is LDVLIV, 450 + 45 + 4.
_CONCISE = {
    1: [(950, "LM"), (450, "LD"), (95, "VC"), (45, "VL")],
    2: [(990, "XM"), (490, "XD"), (99, "IC"), (49, "IL")],
    3: [(995, "VM"), (495, "VD")],
    4: [(999, "IM"), (499, "ID")],
}


def _roman_form(style: int) -> list[tuple[int, str]]:
    pairs = list(_CLASSIC)
    for level in range(1, style + 1):
        pairs.extend(_CONCISE[level])
    return sorted(pairs, key=lambda pair: -pair[0])


@function("ROMAN", V, V, minimum=1)
def ROMAN(context: Context, number: Scalar, form: Scalar | None = None) -> Value:
    value = context.number(number)
    if form is None or isinstance(form, Empty):
        style = 0
    elif isinstance(form, bool):
        style = 0 if form else 4
    else:
        style = context.integer(form)
    if not 0 <= value < 4000 or not 0 <= style <= 4:
        return VALUE
    whole = int(value)
    out = ""
    for amount, letters in _roman_form(style):
        while whole >= amount:
            out += letters
            whole -= amount
    return out


_ROMAN_VALUES = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}


@function("ARABIC", V)
def ARABIC(context: Context, text: Scalar) -> Value:
    source = context.text(text).strip().upper()
    negative = source.startswith("-")
    if negative:
        source = source[1:]
    if len(source) > 255 or any(char not in _ROMAN_VALUES for char in source):
        return VALUE
    total = 0
    for index, char in enumerate(source):
        value = _ROMAN_VALUES[char]
        if index + 1 < len(source) and _ROMAN_VALUES[source[index + 1]] > value:
            total -= value
        else:
            total += value
    return float(-total if negative else total)


# ----------------------------------------------------------------------
# Complex numbers
# ----------------------------------------------------------------------

_PART = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?")


def _parse_complex(context: Context, value: Scalar) -> tuple[complex, str | None]:
    """A complex number from its text or a plain number, and the unit its
    text used, ``None`` for a number with no imaginary part written."""
    if isinstance(value, bool):
        raise ExcelError(VALUE)
    if isinstance(value, float):
        return complex(value, 0), None
    text = context.text(value)
    if text == "":
        return 0j, None
    if text[-1] not in "ij":
        if not _PART.fullmatch(text):
            raise ExcelError(NUM)
        return complex(float(text), 0), None
    unit = text[-1]
    body = text[:-1]
    split = max(
        (index for index in range(1, len(body)) if body[index] in "+-" and body[index - 1] not in "eE"),
        default=0,
    )
    real_text, imaginary_text = body[:split], body[split:]
    if real_text and not _PART.fullmatch(real_text):
        raise ExcelError(NUM)
    if imaginary_text in ("", "+"):
        imaginary = 1.0
    elif imaginary_text == "-":
        imaginary = -1.0
    elif _PART.fullmatch(imaginary_text):
        imaginary = float(imaginary_text)
    else:
        raise ExcelError(NUM)
    return complex(float(real_text) if real_text else 0.0, imaginary), unit


def _complex_text(value: complex, unit: str) -> str:
    real, imaginary = value.real + 0.0, value.imag + 0.0
    if not (math.isfinite(real) and math.isfinite(imaginary)):
        raise ExcelError(NUM)
    real_text = number_text(real) if real != 0 else ""
    if imaginary == 0:
        return real_text or "0"
    if imaginary == 1:
        imaginary_text = unit
    elif imaginary == -1:
        imaginary_text = "-" + unit
    else:
        imaginary_text = number_text(imaginary) + unit
    if real_text and not imaginary_text.startswith("-"):
        imaginary_text = "+" + imaginary_text
    return real_text + imaginary_text


@function("COMPLEX", V, V, V, minimum=2)
def COMPLEX(context: Context, real: Scalar, imaginary: Scalar, suffix: Scalar | None = None) -> Value:
    unit = "i" if suffix is None or isinstance(suffix, Empty) else context.text(suffix)
    if unit not in ("i", "j"):
        return VALUE
    return _complex_text(complex(context.number(real), context.number(imaginary)), unit)


def _complex_to_text(name: str, operation: Callable[[complex], complex]) -> None:
    def implementation(context: Context, value: Scalar) -> Value:
        number, unit = _parse_complex(context, value)
        try:
            result = operation(number)
        except (ValueError, ZeroDivisionError, OverflowError):
            return NUM
        return _complex_text(result, unit or "i")

    function(name, V)(implementation)


def _complex_to_number(name: str, operation: Callable[[complex], float]) -> None:
    def implementation(context: Context, value: Scalar) -> Value:
        number, _ = _parse_complex(context, value)
        return checked(operation(number))

    function(name, V)(implementation)


def _polar_power(number: complex, exponent: float) -> complex:
    """Excel's IMPOWER and IMSQRT: the modulus to the power, the angle
    times it."""
    if number == 0:
        if exponent <= 0:
            raise ValueError("zero to a power not above zero")
        return 0j
    modulus = abs(number) ** exponent
    angle = math.atan2(number.imag, number.real) * exponent
    return complex(modulus * math.cos(angle), modulus * math.sin(angle))


def _nonzero(number: complex) -> complex:
    if number == 0:
        raise ValueError("the logarithm of zero")
    return number


_complex_to_number("IMABS", abs)
_complex_to_number("IMREAL", lambda number: number.real)
_complex_to_number("IMAGINARY", lambda number: number.imag)
_complex_to_text("IMCONJUGATE", lambda number: number.conjugate())
_complex_to_text("IMSQRT", lambda number: _polar_power(number, 0.5))
_complex_to_text("IMEXP", cmath.exp)
_complex_to_text("IMLN", lambda number: cmath.log(_nonzero(number)))
_complex_to_text("IMLOG10", lambda number: cmath.log10(_nonzero(number)))
_complex_to_text("IMLOG2", lambda number: cmath.log(_nonzero(number)) / math.log(2))
_complex_to_text("IMSIN", cmath.sin)
_complex_to_text("IMCOS", cmath.cos)
_complex_to_text("IMTAN", cmath.tan)
_complex_to_text("IMSINH", cmath.sinh)
_complex_to_text("IMCOSH", cmath.cosh)
_complex_to_text("IMSEC", lambda number: 1 / cmath.cos(number))
_complex_to_text("IMSECH", lambda number: 1 / cmath.cosh(number))
_complex_to_text("IMCSC", lambda number: 1 / cmath.sin(number))
_complex_to_text("IMCSCH", lambda number: 1 / cmath.sinh(number))
_complex_to_text("IMCOT", lambda number: 1 / cmath.tan(number))


@function("IMARGUMENT", V)
def IMARGUMENT(context: Context, value: Scalar) -> Value:
    number, _ = _parse_complex(context, value)
    if number == 0:
        return DIV0
    return math.atan2(number.imag, number.real)


@function("IMPOWER", V, V)
def IMPOWER(context: Context, value: Scalar, exponent: Scalar) -> Value:
    number, unit = _parse_complex(context, value)
    try:
        return _complex_text(_polar_power(number, context.number(exponent)), unit or "i")
    except ValueError:
        return NUM


def _complex_arguments(context: Context, args: tuple[Value, ...]) -> tuple[list[complex], str]:
    """The complex numbers in the arguments, and the one unit they share:
    ``i`` and ``j`` in one call are ``#VALUE!``."""
    found: list[complex] = []
    units: set[str] = set()
    for arg in args:
        items = [item for item, _ in context.scalars(arg)] if isinstance(arg, (Reference, Array)) else [arg]
        for item in items:
            if isinstance(item, CellError):
                raise ExcelError(item)
            number, unit = _parse_complex(context, item)
            found.append(number)
            if unit is not None:
                units.add(unit)
    if len(units) > 1:
        raise ExcelError(VALUE)
    return found, units.pop() if units else "i"


@function("IMSUM", R, maximum=255)
def IMSUM(context: Context, *args: Value) -> Value:
    values, unit = _complex_arguments(context, args)
    # In order, as the x87 adds: Python's own sum compensates.
    real = imaginary = 0.0
    for value in values:
        real = add(real, value.real)
        imaginary = add(imaginary, value.imag)
    return _complex_text(complex(real, imaginary), unit)


@function("IMPRODUCT", R, maximum=255)
def IMPRODUCT(context: Context, *args: Value) -> Value:
    values, unit = _complex_arguments(context, args)
    result = 1 + 0j
    for value in values:
        result *= value
    return _complex_text(result, unit)


@function("IMSUB", V, V)
def IMSUB(context: Context, first: Scalar, second: Scalar) -> Value:
    values, unit = _complex_arguments(context, (first, second))
    return _complex_text(values[0] - values[1], unit)


@function("IMDIV", V, V)
def IMDIV(context: Context, first: Scalar, second: Scalar) -> Value:
    values, unit = _complex_arguments(context, (first, second))
    if values[1] == 0:
        return NUM
    return _complex_text(values[0] / values[1], unit)


# ----------------------------------------------------------------------
# Bessel functions
# ----------------------------------------------------------------------


def _polynomial(y: float, coefficients: tuple[float, ...]) -> float:
    """``c0 + y*(c1 + y*(c2 + ...))``, in that order of operations."""
    result = coefficients[-1]
    for coefficient in reversed(coefficients[:-1]):
        result = add(coefficient, multiply(y, result))
    return result


# J0 and J1 below 8: rational approximations; above: the asymptotic form.
# Excel's tables differ from the book's in three places, each measured on
# 288 arguments: J0's Q4 is 267.8532712, not 267.8530712; J0's asymptotic
# P1 is -0.1098628267e-2, where Y0 keeps -0.1098628627e-2; and J1's
# asymptotic P repeats its third coefficient.
_J0_P = (57568490574.0, -13362590354.0, 651619640.7, -11214424.18, 77392.33017, -184.9052456)
_J0_Q = (57568490411.0, 1029532985.0, 9494680.718, 59272.64853, 267.8532712, 1.0)
_J1_P = (72362614232.0, -7895059235.0, 242396853.1, -2972611.439, 15704.48260, -30.16036606)
_J1_Q = (144725228442.0, 2300535178.0, 18583304.74, 99447.43394, 376.9991397, 1.0)
_Y0_P = (-2957821389.0, 7062834065.0, -512359803.6, 10879881.29, -86327.92757, 228.4622733)
_Y0_Q = (40076544269.0, 745249964.8, 7189466.438, 47447.26470, 226.1030244, 1.0)
_Y1_P = (-0.4900604943e13, 0.1275274390e13, -0.5153438139e11, 0.7349264551e9, -0.4237922726e7, 0.8511937935e4)
_Y1_Q = (0.2499580570e14, 0.4244419664e12, 0.3733650367e10, 0.2245904002e8, 0.1020426050e6, 0.3549632885e3, 1.0)
_ASYMPTOTIC_0_J = (1.0, -0.1098628267e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6)
_ASYMPTOTIC_0_JQ = (-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934935152e-7)
_ASYMPTOTIC_0_Y = (1.0, -0.1098628627e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6)
_ASYMPTOTIC_0_YQ = (-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, -0.934945152e-7)
_ASYMPTOTIC_1_J = (1.0, 0.183105e-2, -0.3516396496e-4, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6)
_ASYMPTOTIC_1_Y = (1.0, 0.183105e-2, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6)
_ASYMPTOTIC_1_Q = (0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6)
_TWO_OVER_PI = 0.636619772


def _asymptotic(x: float, shift: float) -> tuple[float, float, float, float, float]:
    """What the forms above 8 share: ``z = 8/x``, ``y = z*z``, the sine and
    cosine of ``x - shift`` as the x87 gives them, and ``sqrt(2/(pi x))``."""
    z = divide(8.0, x)
    sine, cosine = precise.sine_cosine(subtract(x, shift))
    return z, multiply(z, z), extended(sine), extended(cosine), square_root(divide(_TWO_OVER_PI, x))
# I0, I1, K0 and K1: the polynomials of Abramowitz and Stegun 9.8.1 to 9.8.8.
_I0_SMALL = (1.0, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.360768e-1, 0.45813e-2)
_I0_LARGE = (0.39894228, 0.1328592e-1, 0.225319e-2, -0.157565e-2, 0.916281e-2, -0.2057706e-1, 0.2635537e-1,
             -0.1647633e-1, 0.392377e-2)  # fmt: skip
_I1_SMALL = (0.5, 0.87890594, 0.51498869, 0.15084934, 0.2658733e-1, 0.301532e-2, 0.32411e-3)
_I1_LARGE = (0.39894228, -0.3988024e-1, -0.362018e-2, 0.163801e-2, -0.1031555e-1, 0.2282967e-1, -0.2895312e-1,
             0.1787654e-1, -0.420059e-2)  # fmt: skip
_K0_SMALL = (-0.57721566, 0.42278420, 0.23069756, 0.3488590e-1, 0.262698e-2, 0.10750e-3, 0.74e-5)
_K0_LARGE = (1.25331414, -0.7832358e-1, 0.2189568e-1, -0.1062446e-1, 0.587872e-2, -0.251540e-2, 0.53208e-3)
_K1_SMALL = (1.0, 0.15443144, -0.67278579, -0.18156897, -0.1919402e-1, -0.110404e-2, -0.4686e-4)
_K1_LARGE = (1.25331414, 0.23498619, -0.3655620e-1, 0.1504268e-1, -0.780353e-2, 0.325614e-2, -0.68245e-3)


def _j0(x: float) -> float:
    ax = abs(x)
    if ax < 8.0:
        y = multiply(x, x)
        return divide(_polynomial(y, _J0_P), _polynomial(y, _J0_Q))
    z, y, sine, cosine, scale = _asymptotic(ax, 0.785398164)
    p, q = _polynomial(y, _ASYMPTOTIC_0_J), _polynomial(y, _ASYMPTOTIC_0_JQ)
    return multiply(scale, subtract(multiply(cosine, p), multiply(multiply(z, sine), q)))


def _j1(x: float) -> float:
    ax = abs(x)
    if ax < 8.0:
        y = multiply(x, x)
        return divide(multiply(x, _polynomial(y, _J1_P)), _polynomial(y, _J1_Q))
    z, y, sine, cosine, scale = _asymptotic(ax, 2.356194491)
    p, q = _polynomial(y, _ASYMPTOTIC_1_J), _polynomial(y, _ASYMPTOTIC_1_Q)
    # z times (sine times Q), unlike J0's (z times sine) times Q: measured.
    result = multiply(scale, subtract(multiply(cosine, p), multiply(z, multiply(sine, q))))
    return -result if x < 0 else result


#: Miller's downward recurrence: how far above the order to start, for J
#: and for I alike, and when to rescale.
_START = 40.0
_BIG = 1.0e10
_SMALL = 1.0e-10


def _step(j: int, twice: float, current: float) -> float:
    """``j * (2/x) * current``, in that order."""
    return multiply(multiply(float(j), twice), current)


def _jn(n: int, x: float) -> float:
    if n == 0:
        return _j0(x)
    if n == 1:
        return _j1(x)
    ax = abs(x)
    if ax == 0.0:
        return 0.0
    twice = divide(2.0, ax)
    if ax > n:
        # Upward recurrence is stable once x passes the order.
        below, current = _j0(ax), _j1(ax)
        for j in range(1, n):
            below, current = current, subtract(_step(j, twice, current), below)
        result = current
    else:
        start = 2 * ((n + int(math.sqrt(_START * n))) // 2)
        even = False
        above = result = total = 0.0
        current = 1.0
        for j in range(start, 0, -1):
            below = subtract(_step(j, twice, current), above)
            above, current = current, below
            if abs(current) > _BIG:
                current = multiply(current, _SMALL)
                above = multiply(above, _SMALL)
                result = multiply(result, _SMALL)
                total = multiply(total, _SMALL)
            if even:
                total = add(total, current)
            even = not even
            if j == n:
                result = above
        result = divide(result, subtract(multiply(2.0, total), current))
    return -result if x < 0 and n & 1 else result


def _y0(x: float) -> float:
    if x < 8.0:
        y = multiply(x, x)
        ratio = divide(_polynomial(y, _Y0_P), _polynomial(y, _Y0_Q))
        return add(ratio, multiply(_TWO_OVER_PI, multiply(_j0(x), math.log(x))))
    z, y, sine, cosine, scale = _asymptotic(x, 0.785398164)
    p, q = _polynomial(y, _ASYMPTOTIC_0_Y), _polynomial(y, _ASYMPTOTIC_0_YQ)
    return multiply(scale, add(multiply(sine, p), multiply(z, multiply(cosine, q))))


def _y1(x: float) -> float:
    if x < 8.0:
        y = multiply(x, x)
        ratio = divide(multiply(x, _polynomial(y, _Y1_P)), _polynomial(y, _Y1_Q))
        return add(ratio, multiply(_TWO_OVER_PI, subtract(multiply(_j1(x), math.log(x)), divide(1.0, x))))
    z, y, sine, cosine, scale = _asymptotic(x, 2.356194491)
    p, q = _polynomial(y, _ASYMPTOTIC_1_Y), _polynomial(y, _ASYMPTOTIC_1_Q)
    return multiply(scale, add(multiply(sine, p), multiply(z, multiply(cosine, q))))


def _yn(n: int, x: float) -> float:
    if n == 0:
        return _y0(x)
    if n == 1:
        return _y1(x)
    twice = divide(2.0, x)
    below, current = _y0(x), _y1(x)
    for j in range(1, n):
        below, current = current, subtract(_step(j, twice, current), below)
    return current


def _growth(ax: float) -> float:
    """``exp(x) / sqrt(x)``, the large-argument factor of I0 and I1."""
    return divide(math.exp(ax), square_root(ax))


def _i0(x: float) -> float:
    ax = abs(x)
    if ax < 3.75:
        y = divide(x, 3.75)
        return _polynomial(multiply(y, y), _I0_SMALL)
    return multiply(_growth(ax), _polynomial(divide(3.75, ax), _I0_LARGE))


def _i1(x: float) -> float:
    ax = abs(x)
    if ax < 3.75:
        y = divide(x, 3.75)
        result = multiply(ax, _polynomial(multiply(y, y), _I1_SMALL))
    else:
        result = multiply(_polynomial(divide(3.75, ax), _I1_LARGE), _growth(ax))
    return -result if x < 0 else result


def _in(n: int, x: float) -> float:
    if n == 0:
        return _i0(x)
    if n == 1:
        return _i1(x)
    if x == 0.0:
        return 0.0
    twice = divide(2.0, abs(x))
    above = result = 0.0
    current = 1.0
    for j in range(2 * (n + int(math.sqrt(_START * n))), 0, -1):
        below = add(above, _step(j, twice, current))
        above, current = current, below
        if abs(current) > _BIG:
            result = multiply(result, _SMALL)
            current = multiply(current, _SMALL)
            above = multiply(above, _SMALL)
        if j == n:
            result = above
    result = multiply(result, divide(_i0(x), current))
    return -result if x < 0 and n & 1 else result


def _k0(x: float) -> float:
    if x <= 2.0:
        y = divide(multiply(x, x), 4.0)
        return add(multiply(-math.log(divide(x, 2.0)), _i0(x)), _polynomial(y, _K0_SMALL))
    return multiply(divide(math.exp(-x), square_root(x)), _polynomial(divide(2.0, x), _K0_LARGE))


def _k1(x: float) -> float:
    if x <= 2.0:
        y = divide(multiply(x, x), 4.0)
        return add(multiply(math.log(divide(x, 2.0)), _i1(x)), multiply(divide(1.0, x), _polynomial(y, _K1_SMALL)))
    return multiply(divide(math.exp(-x), square_root(x)), _polynomial(divide(2.0, x), _K1_LARGE))


def _kn(n: int, x: float) -> float:
    if n == 0:
        return _k0(x)
    if n == 1:
        return _k1(x)
    twice = divide(2.0, x)
    below, current = _k0(x), _k1(x)
    for j in range(1, n):
        below, current = current, add(below, _step(j, twice, current))
    return current


def _bessel(name: str, evaluate: Callable[[int, float], float], *, positive: bool) -> None:
    def implementation(context: Context, x: Scalar, order: Scalar) -> Value:
        value = context.number(x)
        n = context.integer(order)
        if n < 0 or (positive and value <= 0):
            return NUM
        try:
            return checked(evaluate(n, value))
        except (OverflowError, ZeroDivisionError):
            return NUM

    function(name, V, V)(implementation)


_bessel("BESSELJ", _jn, positive=False)
_bessel("BESSELY", _yn, positive=True)
_bessel("BESSELI", _in, positive=False)
_bessel("BESSELK", _kn, positive=True)


# ----------------------------------------------------------------------
# The error function
# ----------------------------------------------------------------------


def _erf(context: Context, value: Scalar) -> float:
    return special.excel_erf(context.number(value))


@function("ERF", V, V, minimum=1)
def ERF(context: Context, lower: Scalar, upper: Scalar | None = None) -> Value:
    if upper is None or isinstance(upper, Empty):
        return checked(_erf(context, lower))
    # Measured: ERF(1,2) is erf(2) - erf(1) of the two doubles, not the
    # difference rounded once.
    return checked(subtract(_erf(context, upper), _erf(context, lower)))


@function("ERF.PRECISE", V)
def ERF_PRECISE(context: Context, value: Scalar) -> Value:
    return checked(_erf(context, value))


def _erfc(context: Context, value: Scalar) -> Value:
    return checked(special.excel_erfc(context.number(value)))


@function("ERFC", V)
def ERFC(context: Context, value: Scalar) -> Value:
    return _erfc(context, value)


@function("ERFC.PRECISE", V)
def ERFC_PRECISE(context: Context, value: Scalar) -> Value:
    return _erfc(context, value)


# ----------------------------------------------------------------------
# CONVERT
# ----------------------------------------------------------------------

#: Each unit: its quantity and exactly how much of the SI unit it is. Excel
#: keeps the atomic mass unit and the electronvolt at their 2006 CODATA
#: values, and the slug and the psi a unit in the last place off the
#: exact quotients: all measured.
_DEFINITIONS: dict[str, tuple[str, Fraction]] = {
    # Weight and mass, in grams.
    **{unit: ("mass", Fraction(Decimal(grams))) for unit, grams in (
        ("g", "1"), ("lbm", "453.59237"), ("u", "1.660538782e-24"), ("ozm", "28.349523125"),
        ("grain", "0.06479891"), ("cwt", "45359.237"), ("shweight", "45359.237"), ("uk_cwt", "50802.34544"),
        ("lcwt", "50802.34544"), ("hweight", "50802.34544"), ("stone", "6350.29318"), ("ton", "907184.74"),
        ("uk_ton", "1016046.9088"), ("LTON", "1016046.9088"), ("brton", "1016046.9088"))},
    "sg": ("mass", Fraction(14593.902937206363)),
    # Distance, in metres.
    **{unit: ("distance", Fraction(Decimal(metres))) for unit, metres in (
        ("m", "1"), ("mi", "1609.344"), ("Nmi", "1852"), ("in", "0.0254"), ("ft", "0.3048"), ("yd", "0.9144"),
        ("ang", "1e-10"), ("ell", "1.143"), ("ly", "9460730472580800"), ("parsec", "30856775812815532"),
        ("pc", "30856775812815532"))},
    "Pica": ("distance", Fraction(254, 720000)), "Picapt": ("distance", Fraction(254, 720000)),
    "pica": ("distance", Fraction(254, 60000)), "survey_mi": ("distance", Fraction(6336000, 3937)),
    # Time, in seconds.
    **{unit: ("time", Fraction(seconds)) for unit, seconds in (
        ("yr", 31557600), ("day", 86400), ("d", 86400), ("hr", 3600), ("mn", 60), ("min", 60), ("sec", 1), ("s", 1))},
    # Pressure, in pascals.
    **{unit: ("pressure", Fraction(Decimal(pascals))) for unit, pascals in (
        ("Pa", "1"), ("p", "1"), ("atm", "101325"), ("at", "101325"), ("mmHg", "133.322"))},
    "psi": ("pressure", Fraction(6894.757293168362)), "Torr": ("pressure", Fraction(101325, 760)),
    # Force, in newtons.
    **{unit: ("force", Fraction(Decimal(newtons))) for unit, newtons in (
        ("N", "1"), ("dyn", "1e-5"), ("dy", "1e-5"), ("lbf", "4.4482216152605"), ("pond", "0.00980665"))},
    # Energy, in joules.
    **{unit: ("energy", Fraction(Decimal(joules))) for unit, joules in (
        ("J", "1"), ("e", "1e-7"), ("c", "4.184"), ("cal", "4.1868"), ("eV", "1.602176487e-19"),
        ("ev", "1.602176487e-19"), ("HPh", "2684519.537696173"), ("hh", "2684519.537696173"), ("Wh", "3600"),
        ("wh", "3600"), ("flb", "1.3558179483314004"), ("BTU", "1055.05585262"), ("btu", "1055.05585262"))},
    # Power, in watts.
    **{unit: ("power", Fraction(Decimal(watts))) for unit, watts in (
        ("HP", "745.69987158227022"), ("h", "745.69987158227022"), ("PS", "735.49875"), ("W", "1"), ("w", "1"))},
    # Magnetism, in teslas.
    "T": ("magnetism", Fraction(1)), "ga": ("magnetism", Fraction(1, 10000)),
    # Volume, in cubic metres.
    **{unit: ("volume", Fraction(Decimal(cubic))) for unit, cubic in (
        ("tsp", "4.92892159375e-6"), ("tspm", "5e-6"), ("tbs", "1.478676478125e-5"), ("oz", "2.95735295625e-5"),
        ("cup", "0.0002365882365"), ("pt", "0.000473176473"), ("us_pt", "0.000473176473"),
        ("uk_pt", "0.00056826125"), ("qt", "0.000946352946"), ("uk_qt", "0.0011365225"), ("gal", "0.003785411784"),
        ("uk_gal", "0.00454609"), ("l", "0.001"), ("L", "0.001"), ("lt", "0.001"), ("m3", "1"),
        ("mi3", "4168181825.440579584"), ("yd3", "0.764554857984"), ("ft3", "0.028316846592"),
        ("in3", "1.6387064e-5"), ("ang3", "1e-30"), ("barrel", "0.158987294928"), ("bushel", "0.03523907016688"),
        ("regton", "2.8316846592"), ("GRT", "2.8316846592"), ("MTON", "1.13267386368"))},
    # Area, in square metres.
    **{unit: ("area", Fraction(Decimal(square))) for unit, square in (
        ("m2", "1"), ("mi2", "2589988.110336"), ("Nmi2", "3429904"), ("in2", "0.00064516"), ("ft2", "0.09290304"),
        ("yd2", "0.83612736"), ("ang2", "1e-20"), ("ar", "100"), ("ha", "10000"), ("uk_acre", "4046.8564224"),
        ("Morgen", "2500"))},
    "us_acre": ("area", 43560 * Fraction(1200, 3937) ** 2),
    # Information, in bits.
    "bit": ("information", Fraction(1)), "byte": ("information", Fraction(8)),
    # Speed, in metres a second.
    "m/s": ("speed", Fraction(1)), "m/sec": ("speed", Fraction(1)), "m/h": ("speed", Fraction(1, 3600)),
    "m/hr": ("speed", Fraction(1, 3600)), "mph": ("speed", Fraction(Decimal("0.44704"))),
    "kn": ("speed", Fraction(1852, 3600)), "admkn": ("speed", Fraction(6080, 3600) * Fraction(Decimal("0.3048"))),
}  # fmt: skip

#: The unit each quantity converts through. Measured: a unit converted to
#: itself goes through it, which is not always the identity, and only these
#: reproduce every such conversion.
_REFERENCE = {
    "mass": "g", "distance": "ang", "time": "sec", "pressure": "Pa", "force": "dyn", "energy": "e", "power": "W",
    "magnetism": "ga", "volume": "l", "area": "ang2", "information": "bit", "speed": "m/s",
}  # fmt: skip

#: Each unit: its quantity and how many of the quantity's reference unit it is.
_UNITS: dict[str, tuple[str, float]] = {
    unit: (quantity, float(size / _DEFINITIONS[_REFERENCE[quantity]][1]))
    for unit, (quantity, size) in _DEFINITIONS.items()
}

#: Units that take a metric prefix.
_PREFIXABLE = frozenset({
    "g", "u", "m", "ang", "ly", "sec", "s", "Pa", "p", "atm", "at", "mmHg", "N", "dyn", "dy", "pond", "J", "e", "c",
    "cal", "eV", "ev", "Wh", "wh", "W", "w", "T", "ga", "l", "L", "lt", "m2", "m3", "bit", "byte", "m/s", "m/sec",
    "m/h", "m/hr", "ar", "ha", "Torr", "pc", "parsec",
})  # fmt: skip

#: The prefixes, as powers of ten and of two.
_PREFIXES: dict[str, tuple[int, int]] = {
    "Y": (24, 0), "Z": (21, 0), "E": (18, 0), "P": (15, 0), "T": (12, 0), "G": (9, 0), "M": (6, 0), "k": (3, 0),
    "h": (2, 0), "da": (1, 0), "e": (1, 0), "d": (-1, 0), "c": (-2, 0), "m": (-3, 0), "u": (-6, 0), "n": (-9, 0),
    "p": (-12, 0), "f": (-15, 0), "a": (-18, 0), "z": (-21, 0), "y": (-24, 0), "Yi": (0, 80), "Zi": (0, 70),
    "Ei": (0, 60), "Pi": (0, 50), "Ti": (0, 40), "Gi": (0, 30), "Mi": (0, 20), "ki": (0, 10),
}  # fmt: skip

_PREFIX_ORDER = sorted(_PREFIXES, key=len, reverse=True)

_SCALES = {"C": "C", "cel": "C", "F": "F", "fah": "F", "K": "K", "kel": "K", "Rank": "Rank", "Reau": "Reau"}

#: Measured, 2333 conversions: each pair of temperature scales has its own
#: formula, not one through Celsius or Kelvin, and a scale to itself is the
#: number unchanged. Réaumur to Fahrenheit alone is not settled, 44 of 58.
_TEMPERATURES: dict[tuple[str, str], Callable[[float], float]] = {
    ("C", "F"): lambda v: add(multiply(v, 1.8), 32.0),
    ("C", "K"): lambda v: add(v, 273.15),
    ("C", "Rank"): lambda v: multiply(add(v, 273.15), 1.8),
    ("C", "Reau"): lambda v: multiply(v, 0.8),
    ("F", "C"): lambda v: divide(subtract(v, 32.0), 1.8),
    ("F", "K"): lambda v: add(divide(subtract(v, 32.0), 1.8), 273.15),
    ("F", "Rank"): lambda v: add(v, 459.67),
    ("F", "Reau"): lambda v: divide(multiply(subtract(v, 32.0), 0.8), 1.8),
    ("K", "C"): lambda v: subtract(v, 273.15),
    ("K", "F"): lambda v: add(multiply(subtract(v, 273.15), 1.8), 32.0),
    ("K", "Rank"): lambda v: multiply(v, 1.8),
    ("K", "Reau"): lambda v: multiply(subtract(v, 273.15), 0.8),
    ("Rank", "C"): lambda v: subtract(divide(v, 1.8), 273.15),
    ("Rank", "F"): lambda v: subtract(v, 459.67),
    ("Rank", "K"): lambda v: divide(v, 1.8),
    ("Rank", "Reau"): lambda v: multiply(subtract(divide(v, 1.8), 273.15), 0.8),
    ("Reau", "C"): lambda v: divide(v, 0.8),
    ("Reau", "F"): lambda v: add(divide(multiply(divide(v, 0.8), 9.0), 5.0), 32.0),
    ("Reau", "K"): lambda v: add(divide(v, 0.8), 273.15),
    ("Reau", "Rank"): lambda v: multiply(add(divide(v, 0.8), 273.15), 1.8),
}


def _unit(name: str) -> tuple[str, float, int, int] | None:
    """A unit's quantity, its size in the reference unit, and its prefix as
    powers of ten and of two, each counted once per dimension."""
    if name in _UNITS:
        return (*_UNITS[name], 0, 0)
    for prefix in _PREFIX_ORDER:
        base = name[len(prefix) :]
        if name.startswith(prefix) and base in _PREFIXABLE:
            quantity, factor = _UNITS[base]
            tens, twos = _PREFIXES[prefix]
            if base[-1] in "23" and quantity in ("area", "volume"):
                tens, twos = tens * int(base[-1]), twos * int(base[-1])
            return quantity, factor, tens, twos
    return None


@function("CONVERT", V, V, V)
def CONVERT(context: Context, number: Scalar, source: Scalar, target: Scalar) -> Value:
    """Measured: the number times the source's size in the reference unit,
    over the target's, then scaled once by the difference of their
    prefixes' powers of ten, to the bit on 2333 conversions."""
    value = context.number(number)
    first = context.text(source)
    second = context.text(target)
    if first in _SCALES or second in _SCALES:
        if first not in _SCALES or second not in _SCALES:
            return NA
        if _SCALES[first] == _SCALES[second]:
            return value
        return checked(_TEMPERATURES[_SCALES[first], _SCALES[second]](value))
    one = _unit(first)
    other = _unit(second)
    if one is None or other is None or one[0] != other[0]:
        return NA
    result = divide(multiply(value, one[1]), other[1])
    if tens := one[2] - other[2]:
        result = multiply(result, float(Fraction(10) ** tens))
    if twos := one[3] - other[3]:
        result = multiply(result, float(Fraction(2) ** twos))
    return checked(result)


@function("SERIESSUM", V, V, V, R)
def SERIESSUM(context: Context, x: Scalar, start: Scalar, step: Scalar, coefficients: Value) -> Value:
    base = context.number(x)
    power = context.number(start)
    increment = context.number(step)
    total = 0.0
    for index, coefficient in enumerate(numbers(context, (coefficients,))):
        total += coefficient * base ** (power + index * increment)
    return checked(total)


__all__: list[str] = []
