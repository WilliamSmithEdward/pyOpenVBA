"""Arithmetic as Excel's x87 floating-point unit does it, in 80-bit
registers with a 64-bit mantissa.

Every operation Excel performs is rounded twice: to the register's 64 bits,
then to a double when the result is stored. The second rounding usually
changes nothing, but when the exact result lies within 2^-12 of a unit of
the midpoint between two doubles, the first rounding puts it exactly on the
midpoint and the second goes to the even side, which is sometimes the other
side. Measured on crafted cases: ``+``, ``*``, ``/`` and SQRT, SUM,
AVERAGE, PRODUCT and SUMSQ all round twice, 84 of 84, each step stored
before the next. :func:`add`, :func:`multiply`, :func:`divide` and
:func:`square_root` do the same, fast unless a result is near such a tie.

SIN, COS and TAN reduce their argument with pi to 64 bits, the x87's own
constant. ATAN, and the reciprocal a negative power is taken through, show
the double rounding too: ``8^(-1/3)`` is exactly 0.5 because
``1/1.9999999999999998`` rounds to 64 bits as a tie.

The exact paths work in :class:`~fractions.Fraction` and
:class:`~decimal.Decimal` at forty-odd digits, far more than a double
needs, and round once to 64 bits and once to a double at the end.
"""

from __future__ import annotations

import math
from collections.abc import Iterable
from decimal import Decimal, localcontext
from fractions import Fraction

#: Pi to the 64 bits of an x87 register, 0xC90FDAA22168C235 x 2^-62: pi
#: rounded up in its last bit, 5.0e-20 above pi itself.
PI_64 = Fraction(0xC90FDAA22168C235, 2**62)
#: ln 2, log10 2 and log2 e to 64 bits, as FLDLN2, FLDLG2 and FLDL2E load
#: them. LOG10 is only right to the bit with log10 2 rounded so.
_LN2_64 = Fraction(0xB17217F7D1CF79AC, 2**64)
_LG2_64 = Fraction(0x9A209A84FBCFF799, 2**65)
_L2E_64 = Fraction(0xB8AA3B295C17F0BC, 2**63)
#: Pi to more digits than any result here needs.
_PI = Decimal("3.14159265358979323846264338327950288419716939937510582097494459")
_DIGITS = 50
#: The digits the logarithms and exponentials work to: far past the 20
#: that decide a 64-bit rounding.
_WIDE = 45
with localcontext() as _context:
    _context.prec = _WIDE
    _LN2 = Decimal(2).ln()


#: Veltkamp's splitter for a double, 2^27 + 1.
_SPLITTER = 134217729.0
#: Within these magnitudes the error-free transformations below can
#: neither overflow nor lose bits to underflow.
_SMALL = 2.0**-900
_LARGE = 2.0**995
#: How near half a unit a result's error must come before the rounding to
#: 64 bits could land on the midpoint, 1/2 - 2^-12, less a margin for the
#: estimate of a quotient's or a root's error.
_NEAR_TIE = 0.4997


def _mantissa(magnitude: Fraction) -> tuple[int, int]:
    """A positive ``magnitude`` rounded to 64 bits, half to even, as
    ``(mantissa, shift)`` with the value ``mantissa * 2^-shift``."""
    numerator, denominator = magnitude.numerator, magnitude.denominator
    # A shift that puts the quotient's leading bit at 2^63.
    shift = 63 - (numerator.bit_length() - denominator.bit_length())
    while True:
        scaled, divisor = (numerator << shift, denominator) if shift >= 0 else (numerator, denominator << -shift)
        quotient, remainder = divmod(scaled, divisor)
        if quotient.bit_length() > 64:
            shift -= 1
        elif quotient.bit_length() < 64:
            shift += 1
        else:
            break
    twice = 2 * remainder
    if twice > divisor or (twice == divisor and quotient & 1):
        quotient += 1
    return quotient, shift


def register(value: Fraction) -> Fraction:
    """``value`` rounded to a 64-bit mantissa, as an x87 register holds it."""
    if value == 0:
        return Fraction(0)
    mantissa, shift = _mantissa(abs(value))
    rounded = Fraction(mantissa, 2**shift) if shift >= 0 else Fraction(mantissa << -shift)
    return -rounded if value < 0 else rounded


def extended(value: Fraction) -> float:
    """``value`` rounded to a 64-bit mantissa, as an x87 register holds it,
    then to a double, each time half to even."""
    if value == 0:
        return 0.0
    negative = value < 0
    quotient, shift = _mantissa(-value if negative else value)
    # int to float rounds half to even: the second rounding. A result
    # outside the normal range goes through Fraction, which rounds once
    # to the subnormal grid or overflows.
    if -960 <= shift <= 1000:
        result = float(quotient) * 2.0**-shift
    else:
        try:
            result = float(Fraction(quotient, 2**shift) if shift >= 0 else Fraction(quotient << -shift))
        except OverflowError:
            result = math.inf
    return -result if negative else result


def _settled(result: float, error: float) -> bool:
    """Whether ``result + error``, rounded to 64 bits and then to a double,
    is still ``result``: so unless the exact value lies within a 64-bit half
    step of the midpoint to the neighbouring double."""
    if error == 0.0:
        return True
    spacing = math.ulp(result)
    if (error < 0.0) == (result > 0.0) and abs(math.frexp(result)[0]) == 0.5:
        # Toward zero from a power of two the doubles are twice as dense.
        spacing *= 0.5
    return abs(error) < _NEAR_TIE * spacing


def _product_error(a: float, b: float, product: float) -> float:
    """``a * b - product`` exactly, ``product`` being ``a * b`` rounded,
    by Dekker's splitting; all three within the safe magnitudes."""
    t = _SPLITTER * a
    high_a = t - (t - a)
    low_a = a - high_a
    t = _SPLITTER * b
    high_b = t - (t - b)
    low_b = b - high_b
    return ((high_a * high_b - product) + high_a * low_b + low_a * high_b) + low_a * low_b


def _safe(*values: float) -> bool:
    return all(_SMALL < abs(value) < _LARGE for value in values)


def add(a: float, b: float) -> float:
    """``a + b`` as the x87 gives it."""
    total = a + b
    if not math.isfinite(total):
        return total
    # Knuth's two-sum: the rounding error, exactly.
    back = total - a
    error = (a - (total - back)) + (b - back)
    if _settled(total, error):
        return total
    return extended(Fraction(a) + Fraction(b))


def subtract(a: float, b: float) -> float:
    """``a - b`` as the x87 gives it."""
    return add(a, -b)


def summed(values: Iterable[float]) -> float:
    """The values added in order, each sum rounded as the x87 rounds it:
    how Excel's statistics accumulate, with none of SUM's treatment of a
    last addition near zero. Python's own sum is compensated since 3.12
    and gives other bits."""
    total = 0.0
    for value in values:
        total = add(total, value)
    return total


def multiply(a: float, b: float) -> float:
    """``a * b`` as the x87 gives it."""
    product = a * b
    if product == 0.0 or not math.isfinite(product):
        return product
    if not _safe(a, b, product):
        return extended(Fraction(a) * Fraction(b))
    if _settled(product, _product_error(a, b, product)):
        return product
    return extended(Fraction(a) * Fraction(b))


def divide(a: float, b: float) -> float:
    """``a / b`` as the x87 gives it; ``b`` is not zero."""
    quotient = a / b
    if quotient == 0.0 or not math.isfinite(quotient):
        return quotient
    if not _safe(a, b, quotient):
        return extended(Fraction(a) / Fraction(b))
    product = quotient * b
    # The remainder a - quotient*b is a double and comes out exactly.
    remainder = (a - product) - _product_error(quotient, b, product)
    if _settled(quotient, remainder / b):
        return quotient
    return extended(Fraction(a) / Fraction(b))


def square_root(a: float) -> float:
    """The square root of ``a``, not negative, as the x87 gives it."""
    root = math.sqrt(a)
    if root == 0.0 or not math.isfinite(root):
        return root
    if _safe(a, root):
        product = root * root
        remainder = (a - product) - _product_error(root, root, product)
        if _settled(root, remainder / (2.0 * root)):
            return root
    # The root to enough bits that the 64-bit rounding sees no difference:
    # floor(sqrt(a) * 2^k), and half a step more when it is inexact.
    numerator, denominator = a.as_integer_ratio()
    power = denominator.bit_length() - 1
    k = max((power + 1) // 2, 72 - (numerator.bit_length() - power) // 2)
    scaled = numerator * 4**k // denominator
    floor = math.isqrt(scaled)
    inexact = floor * floor != scaled
    return extended(Fraction(2 * floor + inexact, 2 ** (k + 1)))


def _decimal(value: Fraction) -> Decimal:
    return Decimal(value.numerator) / Decimal(value.denominator)


def reduced(value: float) -> tuple[Decimal, int]:
    """``value`` less the nearest multiple of pi/2, with pi to 64 bits, to
    fifty digits, and which quarter turn that multiple is."""
    exact = Fraction(value)
    half = PI_64 / 2
    turns = round(exact / half)
    with localcontext() as context:
        context.prec = _DIGITS
        return _decimal(exact - turns * half), turns % 4


def sin_cos(angle: Decimal) -> tuple[Decimal, Decimal]:
    """The sine and cosine of a small angle, by their series."""
    with localcontext() as context:
        context.prec = _DIGITS
        square = angle * angle
        sine = term = angle
        index = 1
        while True:
            term = -term * square / ((index + 1) * (index + 2))
            index += 2
            if term == 0 or abs(term) < Decimal(10) ** -(_DIGITS + 2):
                break
            sine += term
        cosine = term = Decimal(1)
        index = 0
        while True:
            term = -term * square / ((index + 1) * (index + 2))
            index += 2
            if term == 0 or abs(term) < Decimal(10) ** -(_DIGITS + 2):
                break
            cosine += term
        return +sine, +cosine


def _log2(value: float) -> Fraction:
    """The base-2 logarithm of a positive double, to seventy digits."""
    with localcontext() as context:
        context.prec = _WIDE
        return Fraction(Decimal(value).ln() / _LN2)


def ln(value: float) -> float:
    """LN of a positive double as FYL2X gives it: log2 of the value times
    ln 2 to 64 bits, rounded to 64 bits and then to a double. Measured on
    1760 arguments."""
    return extended(_LN2_64 * _log2(value))


def log10(value: float) -> float:
    """LOG10 of a positive double as FYL2X gives it, with log10 2 to 64
    bits. Measured on 1760 arguments, and the exact value differs."""
    return extended(_LG2_64 * _log2(value))


#: Past these EXP overflows to infinity or underflows to zero.
_EXP_OVER = 710.0
_EXP_UNDER = -746.0


def exp(value: float) -> float:
    """EXP as the x87 computes it: ``t = x * log2 e`` in a register, split
    into its nearest whole number and the rest, ``2^rest - 1`` by F2XM1,
    plus one, and scaled by FSCALE, each step rounded to 64 bits, the
    result then to a double. Measured on 2442 arguments, where the exact
    value rounded twice misses one."""
    if value > _EXP_OVER:
        return math.inf
    if value < _EXP_UNDER:
        return 0.0
    return extended(binary_power(register(Fraction(value) * _L2E_64)))


def binary_log(value: float) -> Fraction:
    """log2 of a positive double, to forty-odd digits, as FYL2X computes it
    before rounding."""
    return _log2(value)


def binary_power(t: Fraction) -> Fraction:
    """2^t as the x87 computes it in a register: FRNDINT to the nearest
    whole number, ``2^rest - 1`` by F2XM1, plus one, then FSCALE, each step
    rounded to 64 bits."""
    whole = round(t)
    with localcontext() as context:
        context.prec = _WIDE
        rest = t - whole
        less_one = (Decimal(rest.numerator) / Decimal(rest.denominator) * _LN2).exp() - 1
    scaled = register(register(Fraction(less_one)) + 1)
    return scaled * Fraction(2) ** whole


def exp_less_one(value: float) -> float:
    """``e^x - 1`` as Excel computes it. Below 1 in size it uses Kahan's
    trick: with ``u`` EXP's double, ``(u - 1) * x / ln u``, where the
    logarithm undoes the error ``u`` was rounded with. From 1 up it is
    ``u - 1``. Each step is rounded as the x87 rounds it. Measured through
    SINH, TANH, COTH and CSCH below 1, 1500 of 1500 each, and through
    WEIBULL.DIST's distribution, 900 of 900."""
    u = exp(value)
    if abs(value) >= 1.0:
        return subtract(u, 1.0)
    if u == 1.0:
        return value
    return divide(multiply(subtract(u, 1.0), value), ln(u))


#: ``ln(1 + x)`` switches from its series to LN at this size of x.
_LN_ONE_PLUS_SERIES = 0.375


def ln_one_plus(value: float) -> float:
    """``ln(1 + x)`` as Excel computes it. Below 0.375 in size it is
    ``2 atanh(s)`` with ``s = x / (2 + x)`` and ``z = s^2``: the terms
    ``s (z^k (1/(2k+1)))`` added in order, largest first, until one no
    longer changes their sum, then ``2 (s + sum)``. From 0.375 up it is LN
    of ``1 + x`` rounded. Each step is rounded as the x87 rounds it.
    Measured through PMT, which is built on it: 1128 rates pinned to one
    double, 745 of them chosen where the sum lands within 0.01 of a unit
    of a midpoint, and the switch at 0.375 one unit either side."""
    if abs(value) >= _LN_ONE_PLUS_SERIES:
        return ln(add(1.0, value))
    s = divide(value, add(2.0, value))
    z = multiply(s, s)
    power = 1.0
    total = 0.0
    order = 1
    while True:
        order += 2
        power = multiply(power, z)
        grown = add(total, multiply(s, multiply(power, 1.0 / order)))
        if grown == total:
            return multiply(2.0, add(s, total))
        total = grown


def sine_cosine(value: float) -> tuple[Fraction, Fraction]:
    """The sine and cosine of a double to fifty digits, its argument reduced
    as the x87 reduces it, with pi to 64 bits."""
    rest, quarter = reduced(value)
    sine, cosine = (Fraction(part) for part in sin_cos(rest))
    return ((sine, cosine), (cosine, -sine), (-sine, -cosine), (-cosine, sine))[quarter]


def sin(value: float) -> float:
    """SIN as the x87 gives it: the reduced sine, rounded to 64 bits and
    then to a double."""
    return extended(sine_cosine(value)[0])


def cos(value: float) -> float:
    """COS as the x87 gives it."""
    return extended(sine_cosine(value)[1])


def atan(value: float) -> Decimal:
    """The arctangent of a double, to fifty digits."""
    with localcontext() as context:
        context.prec = _DIGITS + 5
        x = _decimal(Fraction(value))
        negative = x < 0
        x = -x if negative else x
        inverted = x > 1
        if inverted:
            x = 1 / x
        halvings = 0
        while x > Decimal("0.1"):
            # atan(x) = 2 atan(x / (1 + sqrt(1 + x^2)))
            x = x / (1 + (1 + x * x).sqrt())
            halvings += 1
        square = x * x
        total = term = x
        index = 1
        while True:
            term = -term * square
            index += 2
            piece = term / index
            if piece == 0 or abs(piece) < Decimal(10) ** -(_DIGITS + 4):
                break
            total += piece
        result = total * (2**halvings)
        if inverted:
            result = _PI / 2 - result
        return -result if negative else result


__all__ = [
    "PI_64",
    "add",
    "atan",
    "binary_log",
    "binary_power",
    "cos",
    "divide",
    "exp",
    "exp_less_one",
    "extended",
    "ln",
    "ln_one_plus",
    "log10",
    "multiply",
    "reduced",
    "register",
    "sin",
    "sin_cos",
    "sine_cosine",
    "square_root",
    "subtract",
    "summed",
]
