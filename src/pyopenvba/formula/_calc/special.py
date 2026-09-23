"""Special functions to the last bit: the error function, gamma, and the
incomplete gamma and beta functions that the distributions build on.

Each is computed with :mod:`decimal` at fifty significant digits and rounded
once, so a result is the double nearest the exact value. Measured, that is
what Excel gives for ERF, ERFC and GAMMALN, and for most of its
distributions; where Excel strays by a few units in the last place, as in
the far tails of the normal distribution, the formula corpus records how
far.

The functions here take and give :class:`~decimal.Decimal` values, so the
distributions can combine them without rounding in between, and a double
passed in is exact: ``Decimal(0.1)`` is the double nearest 0.1, not 0.1.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from decimal import Context, Decimal, localcontext
from fractions import Fraction
from functools import cache

from pyopenvba.formula._calc import precise

#: Working precision, in significant digits.
DIGITS = 50

CONTEXT = Context(prec=DIGITS, Emax=10**6, Emin=-(10**6))

#: Enough of pi for fifty digits and a margin.
PI = Decimal("3.14159265358979323846264338327950288419716939937510582097494459230781640628620899863")

_ONE = Decimal(1)
_TWO = Decimal(2)
_HALF = Decimal("0.5")
#: A step below this, relative to the sum, cannot change fifty digits.
_TINY = Decimal(10) ** -(DIGITS + 5)
_LIMIT = 100_000


class ConvergenceError(ArithmeticError):
    """A series or continued fraction did not settle; the caller answers
    ``#NUM!``, as Excel does when its own iteration gives up."""


def d(value: float | int | Decimal) -> Decimal:
    """A number as an exact Decimal."""
    return value if isinstance(value, Decimal) else Decimal(value)


def nearest(value: Decimal) -> float:
    """The double nearest a Decimal, or an overflow to infinity."""
    return float(value)


def exp(value: Decimal) -> Decimal:
    return value.exp(CONTEXT)


def ln(value: Decimal) -> Decimal:
    return value.ln(CONTEXT)


def sqrt(value: Decimal) -> Decimal:
    return value.sqrt(CONTEXT)


_SQRT_PI = PI.sqrt(CONTEXT)
_SQRT_2 = Decimal(2).sqrt(CONTEXT)
_LN_2PI_HALF = (2 * PI).ln(CONTEXT) / 2


def _converged(step: Decimal, total: Decimal) -> bool:
    return abs(step) <= abs(total) * _TINY


# ----------------------------------------------------------------------
# The error function
# ----------------------------------------------------------------------


def _erf_series(x: Decimal) -> Decimal:
    """erf(x) for x >= 0 by the series of positive terms
    ``2/sqrt(pi) exp(-x^2) sum (2x^2)^n x / (1*3*...*(2n+1))``."""
    with localcontext(CONTEXT) as context:
        # The factor exp(-x^2) cancels nothing, so fifty digits hold.
        context.prec = DIGITS + 10
        square = x * x
        term = x
        total = term
        n = 0
        while True:
            n += 1
            term = term * 2 * square / (2 * n + 1)
            total += term
            if _converged(term, total):
                break
            if n > _LIMIT:
                raise ConvergenceError("erf")
        return 2 / _SQRT_PI * (-square).exp() * total


def _erfc_fraction(x: Decimal) -> Decimal:
    """erfc(x) for x >= 3 by Laplace's continued fraction
    ``exp(-x^2)/sqrt(pi) / (x + (1/2)/(x + 1/(x + (3/2)/(x + ...))))``."""
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        tiny = Decimal(10) ** -300
        value = x
        c = x
        dd = Decimal(0)
        n = 0
        while True:
            n += 1
            a = Decimal(n) / 2
            dd = x + a * dd
            if dd == 0:
                dd = tiny
            c = x + a / c
            if c == 0:
                c = tiny
            dd = 1 / dd
            delta = c * dd
            value *= delta
            if abs(delta - 1) <= _TINY:
                break
            if n > _LIMIT:
                raise ConvergenceError("erfc")
        return (-(x * x)).exp() / _SQRT_PI / value


def erf(x: Decimal) -> Decimal:
    if x == 0:
        return Decimal(0)
    if x < 0:
        return -erf(-x)
    if x >= 3:
        return 1 - _erfc_fraction(x)
    return _erf_series(x)


def erfc(x: Decimal) -> Decimal:
    if x < 0:
        return 2 - erfc(-x)
    if x >= 3:
        return _erfc_fraction(x)
    # 1 - erf(x) loses as many digits as erfc(3) = 2e-5 has zeros: few.
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        return +(1 - _erf_series(x))


def normal_cdf(x: Decimal) -> Decimal:
    """The standard normal distribution below ``x``."""
    return erfc(-x / _SQRT_2) / 2


def normal_pdf(x: Decimal) -> Decimal:
    return exp(-(x * x) / 2) / sqrt(2 * PI)


# ----------------------------------------------------------------------
# ERF and ERFC the way Excel assembles them
# ----------------------------------------------------------------------

#: Excel's ERF and ERFC go through the incomplete gamma function at a = 1/2
#: of the argument squared, and turn from its lower to its upper part here.
_ERF_TURN = 0.25


def _erf_of_square(square: float) -> float:
    """erf(sqrt(square)), the double nearest the exact value."""
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        return float(erf(Decimal(square).sqrt()))


def _erfc_of_square(square: float) -> float:
    """erfc(sqrt(square)) as Excel assembles it: EXP(-square) as the x87
    gives it, times ``erfc(sqrt(square)) * exp(square)`` rounded once."""
    decay = precise.exp(-square)
    if decay == 0.0:
        return 0.0
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        exact = Decimal(square)
        scaled = float(erfc(exact.sqrt()) * exact.exp())
    return precise.multiply(decay, scaled)


def excel_erf(x: float) -> float:
    """ERF as Excel computes it, measured.

    The argument is squared to a double first, as GAMMA.DIST(x*x, 0.5, 1,
    TRUE) would take it; below 0.25 the result is erf of the root, from
    there ``1 - ERFC``, 726 of 744. Excel's own approximations for the two
    parts are not the exact functions used here, a few units in the last
    place apart at most, but squaring first is what moves ERFC hundreds of
    units from the exact value in the tail, and that part is Excel's."""
    square = precise.multiply(x, x)
    found = _erf_of_square(square) if square < _ERF_TURN else precise.subtract(1.0, _erfc_of_square(square))
    return -found if x < 0 else found


def excel_erfc(x: float) -> float:
    """ERFC as Excel computes it, measured: from the squared argument, as
    :func:`excel_erf`; below the turn ``1 - erf`` or ``1 + erf``, above it
    the upper part, or 2 less it for a negative argument, 213 of 220."""
    square = precise.multiply(x, x)
    if square < _ERF_TURN:
        found = _erf_of_square(square)
        return precise.add(1.0, found) if x < 0 else precise.subtract(1.0, found)
    found = _erfc_of_square(square)
    return precise.subtract(2.0, found) if x < 0 else found


# ----------------------------------------------------------------------
# Gamma
# ----------------------------------------------------------------------


@cache
def _bernoulli(count: int) -> tuple[Fraction, ...]:
    """B_2, B_4, ..., B_2count."""
    numbers = [Fraction(1)]
    for m in range(1, 2 * count + 1):
        numbers.append(-sum((Fraction(math.comb(m + 1, k)) * numbers[k] for k in range(m)), Fraction(0)) / (m + 1))
    return tuple(numbers[2 * k] for k in range(1, count + 1))


_STIRLING_FROM = 40
_STIRLING_TERMS = 20


def lgamma(x: Decimal) -> Decimal:
    """ln(gamma(x)) for x > 0."""
    if x <= 0:
        raise ValueError("lgamma of a number not above zero")
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        shift = Decimal(1)
        z = x
        while z < _STIRLING_FROM:
            shift *= z
            z += 1
        total = (z - _HALF) * z.ln() - z + _LN_2PI_HALF
        power = z
        square = z * z
        for index, number in enumerate(_bernoulli(_STIRLING_TERMS), start=1):
            total += Decimal(number.numerator) / Decimal(number.denominator) / (2 * index * (2 * index - 1) * power)
            power *= square
        return +(total - shift.ln())


def sin_pi(x: Decimal) -> Decimal:
    """sin(pi*x), reducing x exactly first."""
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        reduced = x % 2
        if reduced > 1:
            reduced -= 2
        elif reduced < -1:
            reduced += 2
        # sin(pi*r) = sin(pi*(1-r)) keeps the angle within [-pi/2, pi/2].
        if reduced > _HALF:
            reduced = 1 - reduced
        elif reduced < -_HALF:
            reduced = -1 - reduced
        angle = PI * reduced
        square = angle * angle
        term = angle
        total = term
        n = 1
        while True:
            term = -term * square / ((n + 1) * (n + 2))
            n += 2
            total += term
            if _converged(term, total) or term == 0:
                return total


def gamma(x: Decimal) -> Decimal:
    """gamma(x) for any x but zero and the negative whole numbers."""
    if x > 0:
        return exp(lgamma(x))
    if x == x.to_integral_value():
        raise ValueError("gamma at a pole")
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        return +(PI / (sin_pi(x) * exp(lgamma(1 - x))))


def lbeta(a: Decimal, b: Decimal) -> Decimal:
    return lgamma(a) + lgamma(b) - lgamma(a + b)


# ----------------------------------------------------------------------
# The incomplete gamma function
# ----------------------------------------------------------------------


def _gamma_factor(a: Decimal, x: Decimal) -> Decimal:
    """``x^a exp(-x) / gamma(a)``."""
    return exp(a * ln(x) - x - lgamma(a))


def gamma_lower(a: Decimal, x: Decimal) -> Decimal:
    """P(a, x), the regularized lower incomplete gamma function."""
    if x <= 0:
        return Decimal(0)
    if x < a + 1:
        return _gamma_series(a, x)
    return 1 - _gamma_fraction(a, x)


def gamma_upper(a: Decimal, x: Decimal) -> Decimal:
    """Q(a, x) = 1 - P(a, x)."""
    if x <= 0:
        return Decimal(1)
    if x < a + 1:
        return 1 - _gamma_series(a, x)
    return _gamma_fraction(a, x)


def _gamma_series(a: Decimal, x: Decimal) -> Decimal:
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        term = 1 / a
        total = term
        n = 0
        while True:
            n += 1
            term = term * x / (a + n)
            total += term
            if _converged(term, total):
                break
            if n > _LIMIT:
                raise ConvergenceError("incomplete gamma")
        return +(total * _gamma_factor(a, x))


def _gamma_fraction(a: Decimal, x: Decimal) -> Decimal:
    """Q(a, x) by Legendre's continued fraction, evaluated by Lentz."""
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        tiny = Decimal(10) ** -300
        b = x + 1 - a
        c = 1 / tiny
        dd = 1 / b
        value = dd
        n = 0
        while True:
            n += 1
            an = -n * (n - a)
            b += 2
            dd = an * dd + b
            if abs(dd) < tiny:
                dd = tiny
            c = b + an / c
            if abs(c) < tiny:
                c = tiny
            dd = 1 / dd
            delta = dd * c
            value *= delta
            if abs(delta - 1) <= _TINY:
                break
            if n > _LIMIT:
                raise ConvergenceError("incomplete gamma")
        return +(value * _gamma_factor(a, x))


# ----------------------------------------------------------------------
# The incomplete beta function
# ----------------------------------------------------------------------


def beta_lower(a: Decimal, b: Decimal, x: Decimal) -> Decimal:
    """I_x(a, b), the regularized incomplete beta function."""
    if x <= 0:
        return Decimal(0)
    if x >= 1:
        return Decimal(1)
    if x * (a + b + 2) < a + 1:
        return _beta_fraction(a, b, x)
    return 1 - _beta_fraction(b, a, 1 - x)


def beta_upper(a: Decimal, b: Decimal, x: Decimal) -> Decimal:
    """1 - I_x(a, b), computed without the subtraction where it is small."""
    if x <= 0:
        return Decimal(1)
    if x >= 1:
        return Decimal(0)
    if x * (a + b + 2) < a + 1:
        return 1 - _beta_fraction(a, b, x)
    return _beta_fraction(b, a, 1 - x)


def _beta_fraction(a: Decimal, b: Decimal, x: Decimal) -> Decimal:
    """I_x(a, b) by its continued fraction, for x below the mean."""
    with localcontext(CONTEXT) as context:
        context.prec = DIGITS + 10
        tiny = Decimal(10) ** -300
        both = a + b
        above = a + 1
        below = a - 1
        c = Decimal(1)
        dd = 1 - both * x / above
        if abs(dd) < tiny:
            dd = tiny
        dd = 1 / dd
        value = dd
        m = 0
        while True:
            m += 1
            twice = 2 * m
            step = m * (b - m) * x / ((below + twice) * (a + twice))
            dd = 1 + step * dd
            if abs(dd) < tiny:
                dd = tiny
            c = 1 + step / c
            if abs(c) < tiny:
                c = tiny
            dd = 1 / dd
            value *= dd * c
            step = -(a + m) * (both + m) * x / ((a + twice) * (above + twice))
            dd = 1 + step * dd
            if abs(dd) < tiny:
                dd = tiny
            c = 1 + step / c
            if abs(c) < tiny:
                c = tiny
            dd = 1 / dd
            delta = dd * c
            value *= delta
            if abs(delta - 1) <= _TINY:
                break
            if m > _LIMIT:
                raise ConvergenceError("incomplete beta")
        front = exp(a * ln(x) + b * ln(1 - x) - lbeta(a, b)) / a
        return +(front * value)


# ----------------------------------------------------------------------
# Inverting a distribution
# ----------------------------------------------------------------------


def solve(
    function: Callable[[Decimal], Decimal],
    derivative: Callable[[Decimal], Decimal],
    target: Decimal,
    guess: Decimal,
    low: Decimal | None,
    high: Decimal | None,
    *,
    decreasing: bool = False,
) -> Decimal:
    """The x where ``function(x) == target``, by Newton's method kept inside
    a shrinking bracket, to far more digits than a double holds.

    ``low`` and ``high`` bound the answer where known; ``function`` must be
    monotonic, increasing unless ``decreasing``.
    """
    x = guess
    for _ in range(400):
        value = function(x)
        error = value - target
        if error == 0:
            return x
        if (error > 0) != decreasing:
            high = x
        else:
            low = x
        slope = derivative(x)
        step = error / slope if slope != 0 else None
        candidate = x - step if step is not None else None
        if candidate is None or (low is not None and candidate <= low) or (high is not None and candidate >= high):
            if low is not None and high is not None:
                candidate = (low + high) / 2
            elif low is not None:
                candidate = low * 2 if low > 0 else low + 1
            elif high is not None:
                candidate = high / 2 if high > 0 else high - 1
            else:  # pragma: no cover - some bound always exists once the sign of the error is known
                raise ConvergenceError("no bracket")
        if x != 0 and abs(candidate - x) <= abs(x) * _TINY * 10**10:
            return candidate
        if low is not None and high is not None and high - low <= abs(high) * _TINY * 10**10:
            return candidate
        x = candidate
    raise ConvergenceError("inverse")


__all__ = [
    "CONTEXT",
    "DIGITS",
    "PI",
    "ConvergenceError",
    "beta_lower",
    "beta_upper",
    "d",
    "erf",
    "erfc",
    "excel_erf",
    "excel_erfc",
    "exp",
    "gamma",
    "gamma_lower",
    "gamma_upper",
    "lbeta",
    "lgamma",
    "ln",
    "nearest",
    "normal_cdf",
    "normal_pdf",
    "sin_pi",
    "solve",
    "sqrt",
]
