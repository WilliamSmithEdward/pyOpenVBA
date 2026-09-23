"""Probability distributions and the tests built on them.

The continuous distributions, normal, t, chi-square, F, gamma, beta,
lognormal, are computed from :mod:`~pyopenvba.formula._calc.special` to
fifty digits and rounded once, and their inverses solved to the same
precision. Measured, Excel gives the nearest double for most of them and
strays by a few units in the last place for the rest, by up to a few
dozen in the far tails of the normal distribution.

Where Excel is a formula in doubles the engine uses the same one, found by
measuring: BINOM.DIST's density is ``exp(ln(COMBIN(n,k)) + k*ln(p) +
(n-k)*ln(1-p))``, POISSON's terms are ``exp(-mean) * mean^k / k!`` summed
from 0, and MULTINOMIAL is ``exp`` of the log-gamma differences.

Excel cuts the degrees of freedom of T.DIST, CHISQ.DIST and F.DIST down
to whole numbers: ``T.DIST(0.5,1.5,TRUE)`` is ``T.DIST(0.5,1,TRUE)``.
"""

from __future__ import annotations

import math
import statistics
from collections.abc import Callable
from decimal import Decimal, localcontext

from pyopenvba.formula._calc import precise, special
from pyopenvba.formula._calc.evaluator import Context, power
from pyopenvba.formula._calc.functions.aggregate import spread
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import matrix, numbers, pairs
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import DIV0, NA, NUM, Empty, ExcelError, Scalar, Value
from pyopenvba.formula._calc.cells import CellError

D = Decimal
Exact = Callable[[Decimal], Decimal]
#: Where WEIBULL.DIST's distribution turns from ``-expm1(-t)`` to ``1 - exp(-t)``.
_HALF_LIFE = math.log(2.0)
#: 1/sqrt(2 pi), the double the normal density multiplies by.
_INVERSE_ROOT_TWO_PI = 0.3989422804014327


def _exact(value: Decimal) -> float:
    return checked(float(value))


def _flag(context: Context, value: Scalar) -> bool:
    return context.logical(value)


def _optional(context: Context, value: Scalar | None, default: float) -> float:
    return default if value is None or isinstance(value, Empty) else context.number(value)


def _invert(
    cdf: Exact,
    pdf: Exact,
    target: float,
    guess: float,
    low: float | None,
    high: float | None,
    *,
    decreasing: bool = False,
) -> float:
    """The x a distribution function gives ``target`` at, to the bit."""
    with localcontext(special.CONTEXT):
        try:
            found = special.solve(
                cdf,
                pdf,
                D(target),
                D(guess),
                None if low is None else D(low),
                None if high is None else D(high),
                decreasing=decreasing,
            )
        except special.ConvergenceError:
            raise ExcelError(NUM) from None
    return _exact(found)


# ----------------------------------------------------------------------
# Normal
# ----------------------------------------------------------------------


#: SQRT(0.5), what NORM.S.DIST scales its argument by for ERFC.
_ROOT_HALF = precise.square_root(0.5)


def normal_cdf(z: float) -> float:
    """NORM.S.DIST, measured: ``0.5*ERFC(-z*SQRT(0.5))`` to the bit, the
    ERFC as :func:`~.special.excel_erfc` has it."""
    return checked(0.5 * special.excel_erfc(precise.multiply(-z, _ROOT_HALF)))


def normal_pdf(z: float, deviation: float = 1.0) -> float:
    """The normal density, measured on 1358 standard and 400 general
    arguments: ``EXP(-z*z/2)`` as the x87 gives it, divided by the deviation,
    times 1/sqrt(2 pi) as a double."""
    decay = precise.exp(-precise.divide(precise.multiply(z, z), 2.0))
    return checked(precise.multiply(precise.divide(decay, deviation), _INVERSE_ROOT_TWO_PI))


def normal_inverse(p: float) -> float:
    if not 0 < p < 1:
        raise ExcelError(NUM)
    if p == 0.5:
        return 0.0
    guess = statistics.NormalDist().inv_cdf(p)
    return _invert(special.normal_cdf, special.normal_pdf, p, guess, None, None)


def _normal(context: Context, x: Scalar, mean: Scalar, deviation: Scalar, cumulative: Scalar) -> Value:
    spread_ = context.number(deviation)
    if spread_ <= 0:
        return NUM
    z = precise.divide(precise.subtract(context.number(x), context.number(mean)), spread_)
    return normal_cdf(z) if _flag(context, cumulative) else normal_pdf(z, spread_)


@function("NORM.DIST", V, V, V, V)
def NORM_DIST(context: Context, x: Scalar, mean: Scalar, deviation: Scalar, cumulative: Scalar) -> Value:
    return _normal(context, x, mean, deviation, cumulative)


@function("NORMDIST", V, V, V, V)
def NORMDIST(context: Context, x: Scalar, mean: Scalar, deviation: Scalar, cumulative: Scalar) -> Value:
    return _normal(context, x, mean, deviation, cumulative)


@function("NORM.S.DIST", V, V)
def NORM_S_DIST(context: Context, z: Scalar, cumulative: Scalar) -> Value:
    value = context.number(z)
    return normal_cdf(value) if _flag(context, cumulative) else normal_pdf(value)


@function("NORMSDIST", V)
def NORMSDIST(context: Context, z: Scalar) -> Value:
    return normal_cdf(context.number(z))


def _normal_inverse(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    spread_ = context.number(deviation)
    if spread_ <= 0:
        return NUM
    return checked(context.number(mean) + spread_ * normal_inverse(context.number(p)))


@function("NORM.INV", V, V, V)
def NORM_INV(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    return _normal_inverse(context, p, mean, deviation)


@function("NORMINV", V, V, V)
def NORMINV(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    return _normal_inverse(context, p, mean, deviation)


@function("NORM.S.INV", V)
def NORM_S_INV(context: Context, p: Scalar) -> Value:
    return normal_inverse(context.number(p))


@function("NORMSINV", V)
def NORMSINV(context: Context, p: Scalar) -> Value:
    return normal_inverse(context.number(p))


@function("PHI", V)
def PHI(context: Context, x: Scalar) -> Value:
    return normal_pdf(context.number(x))


@function("GAUSS", V)
def GAUSS(context: Context, x: Scalar) -> Value:
    with localcontext(special.CONTEXT):
        return _exact(special.normal_cdf(D(context.number(x))) - D("0.5"))


def _confidence(context: Context, alpha: Scalar, deviation: Scalar, size: Scalar) -> tuple[float, float, int]:
    level = context.number(alpha)
    spread_ = context.number(deviation)
    count = math.trunc(context.number(size))
    if not 0 < level < 1 or spread_ <= 0 or count < 1:
        raise ExcelError(NUM)
    return level, spread_, count


@function("CONFIDENCE", V, V, V)
def CONFIDENCE(context: Context, alpha: Scalar, deviation: Scalar, size: Scalar) -> Value:
    level, spread_, count = _confidence(context, alpha, deviation, size)
    return checked(-normal_inverse(level / 2) * spread_ / math.sqrt(count))


@function("CONFIDENCE.NORM", V, V, V)
def CONFIDENCE_NORM(context: Context, alpha: Scalar, deviation: Scalar, size: Scalar) -> Value:
    return CONFIDENCE(context, alpha, deviation, size)


@function("CONFIDENCE.T", V, V, V)
def CONFIDENCE_T(context: Context, alpha: Scalar, deviation: Scalar, size: Scalar) -> Value:
    level, spread_, count = _confidence(context, alpha, deviation, size)
    if count == 1:
        return DIV0
    return checked(t_inverse(1 - level / 2, count - 1) * spread_ / math.sqrt(count))


# ----------------------------------------------------------------------
# t
# ----------------------------------------------------------------------


def _freedom(context: Context, value: Scalar, *, least: float = 1.0) -> float:
    found = math.trunc(context.number(value))
    if found < least or found > 1e10:
        raise ExcelError(NUM)
    return float(found)


def _t_upper(x: Decimal, n: Decimal) -> Decimal:
    """P(T > x) for x >= 0."""
    return special.beta_lower(n / 2, D("0.5"), n / (n + x * x)) / 2


def _t_cdf(x: Decimal, n: Decimal) -> Decimal:
    if x >= 0:
        return 1 - _t_upper(x, n)
    return _t_upper(-x, n)


def _t_pdf(x: Decimal, n: Decimal) -> Decimal:
    log = special.lgamma((n + 1) / 2) - special.lgamma(n / 2) - special.ln(n * special.PI) / 2
    return special.exp(log - (n + 1) / 2 * special.ln(1 + x * x / n))


def t_cdf(x: float, n: float) -> float:
    with localcontext(special.CONTEXT):
        return _exact(_t_cdf(D(x), D(n)))


def t_upper(x: float, n: float) -> float:
    with localcontext(special.CONTEXT):
        return _exact(_t_upper(D(x), D(n)))


def t_inverse(p: float, n: float) -> float:
    """The x with P(T <= x) = p."""
    if not 0 < p < 1:
        raise ExcelError(NUM)
    if p == 0.5:
        return 0.0
    exact_n = D(n)
    guess = statistics.NormalDist().inv_cdf(p)
    return _invert(lambda x: _t_cdf(x, exact_n), lambda x: _t_pdf(x, exact_n), p, guess, None, None)


@function("T.DIST", V, V, V)
def T_DIST(context: Context, x: Scalar, freedom: Scalar, cumulative: Scalar) -> Value:
    value = context.number(x)
    n = _freedom(context, freedom)
    if _flag(context, cumulative):
        return t_cdf(value, n)
    with localcontext(special.CONTEXT):
        return _exact(_t_pdf(D(value), D(n)))


@function("T.DIST.RT", V, V)
def T_DIST_RT(context: Context, x: Scalar, freedom: Scalar) -> Value:
    value = context.number(x)
    n = _freedom(context, freedom)
    if value >= 0:
        return t_upper(value, n)
    with localcontext(special.CONTEXT):
        return _exact(1 - _t_upper(D(-value), D(n)))


@function("T.DIST.2T", V, V)
def T_DIST_2T(context: Context, x: Scalar, freedom: Scalar) -> Value:
    value = context.number(x)
    n = _freedom(context, freedom)
    if value < 0:
        return NUM
    with localcontext(special.CONTEXT):
        return _exact(2 * _t_upper(D(value), D(n)))


@function("TDIST", V, V, V)
def TDIST(context: Context, x: Scalar, freedom: Scalar, tails: Scalar) -> Value:
    value = context.number(x)
    n = _freedom(context, freedom)
    sides = math.trunc(context.number(tails))
    if value < 0 or sides not in (1, 2):
        return NUM
    with localcontext(special.CONTEXT):
        return _exact(sides * _t_upper(D(value), D(n)))


@function("T.INV", V, V)
def T_INV(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return t_inverse(context.number(p), _freedom(context, freedom))


def _two_tailed_inverse(context: Context, p: Scalar, freedom: Scalar) -> Value:
    chance = context.number(p)
    if not 0 < chance <= 1:
        return NUM
    if chance == 1:
        return 0.0
    return t_inverse(1 - chance / 2, _freedom(context, freedom))


@function("T.INV.2T", V, V)
def T_INV_2T(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return _two_tailed_inverse(context, p, freedom)


@function("TINV", V, V)
def TINV(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return _two_tailed_inverse(context, p, freedom)


def _t_test(context: Context, first: Value, second: Value, tails: Scalar, kind: Scalar) -> Value:
    sides = math.trunc(context.number(tails))
    which = math.trunc(context.number(kind))
    if sides not in (1, 2) or which not in (1, 2, 3):
        return NUM
    if which == 1:
        found = pairs(context, first, second)
        if len(found) < 2:
            return DIV0
        differences = [x - y for x, y in found]
        count = len(differences)
        variance = spread(differences, sample=True)
        if isinstance(variance, CellError):
            return variance
        mean = precise.summed(differences) / count
        if variance == 0:
            return DIV0
        statistic = mean / math.sqrt(variance / count)
        freedom = float(count - 1)
    else:
        a = numbers(context, (first,))
        b = numbers(context, (second,))
        if len(a) < 2 or len(b) < 2:
            return DIV0
        var_a = spread(a, sample=True)
        var_b = spread(b, sample=True)
        if isinstance(var_a, CellError) or isinstance(var_b, CellError):
            return DIV0
        mean_a = precise.summed(a) / len(a)
        mean_b = precise.summed(b) / len(b)
        if which == 2:
            freedom = float(len(a) + len(b) - 2)
            pooled = ((len(a) - 1) * var_a + (len(b) - 1) * var_b) / freedom
            error = math.sqrt(pooled * (1 / len(a) + 1 / len(b)))
        else:
            part_a = var_a / len(a)
            part_b = var_b / len(b)
            error = math.sqrt(part_a + part_b)
            freedom = (part_a + part_b) ** 2 / (part_a**2 / (len(a) - 1) + part_b**2 / (len(b) - 1))
        if error == 0:
            return DIV0
        statistic = (mean_a - mean_b) / error
    with localcontext(special.CONTEXT):
        return _exact(sides * _t_upper(D(abs(statistic)), D(freedom)))


@function("T.TEST", R, R, V, V)
def T_TEST(context: Context, first: Value, second: Value, tails: Scalar, kind: Scalar) -> Value:
    return _t_test(context, first, second, tails, kind)


@function("TTEST", R, R, V, V)
def TTEST(context: Context, first: Value, second: Value, tails: Scalar, kind: Scalar) -> Value:
    return _t_test(context, first, second, tails, kind)


# ----------------------------------------------------------------------
# Chi-square and gamma
# ----------------------------------------------------------------------


def _gamma_pdf(x: Decimal, shape: Decimal, scale: Decimal) -> Decimal:
    if x <= 0:
        return D(0)
    return special.exp((shape - 1) * special.ln(x) - x / scale - special.lgamma(shape) - shape * special.ln(scale))


def _gamma_guess(p: float, shape: float, scale: float) -> float:
    """Wilson and Hilferty's approximation, for a start."""
    z = statistics.NormalDist().inv_cdf(min(max(p, 1e-300), 1 - 1e-16))
    k = 2 * shape
    guess = k * (1 - 2 / (9 * k) + z * math.sqrt(2 / (9 * k))) ** 3 / 2 * scale
    return guess if guess > 0 else shape * scale * 0.1


def gamma_inverse(p: float, shape: float, scale: float, *, upper: bool = False) -> float:
    exact_shape, exact_scale = D(shape), D(scale)
    if upper:
        return _invert(
            lambda x: special.gamma_upper(exact_shape, x / exact_scale),
            lambda x: -_gamma_pdf(x, exact_shape, exact_scale),
            p,
            _gamma_guess(1 - p, shape, scale),
            0.0,
            None,
            decreasing=True,
        )
    return _invert(
        lambda x: special.gamma_lower(exact_shape, x / exact_scale),
        lambda x: _gamma_pdf(x, exact_shape, exact_scale),
        p,
        _gamma_guess(p, shape, scale),
        0.0,
        None,
    )


def _chi_square(context: Context, x: Scalar, freedom: Scalar) -> tuple[float, float]:
    value = context.number(x)
    n = _freedom(context, freedom)
    if value < 0:
        raise ExcelError(NUM)
    return value, n


@function("CHISQ.DIST", V, V, V)
def CHISQ_DIST(context: Context, x: Scalar, freedom: Scalar, cumulative: Scalar) -> Value:
    value, n = _chi_square(context, x, freedom)
    with localcontext(special.CONTEXT):
        if _flag(context, cumulative):
            return _exact(special.gamma_lower(D(n) / 2, D(value) / 2))
        return _exact(_gamma_pdf(D(value), D(n) / 2, D(2)))


@function("CHISQ.DIST.RT", V, V)
def CHISQ_DIST_RT(context: Context, x: Scalar, freedom: Scalar) -> Value:
    value, n = _chi_square(context, x, freedom)
    with localcontext(special.CONTEXT):
        return _exact(special.gamma_upper(D(n) / 2, D(value) / 2))


@function("CHIDIST", V, V)
def CHIDIST(context: Context, x: Scalar, freedom: Scalar) -> Value:
    return CHISQ_DIST_RT(context, x, freedom)


def _chi_inverse(context: Context, p: Scalar, freedom: Scalar, *, upper: bool) -> Value:
    chance = context.number(p)
    n = _freedom(context, freedom)
    if upper:
        if not 0 < chance <= 1:
            return NUM
        if chance == 1:
            return 0.0
    elif not 0 <= chance < 1:
        return NUM
    if chance == 0:
        return 0.0
    return gamma_inverse(chance, n / 2, 2.0, upper=upper)


@function("CHISQ.INV", V, V)
def CHISQ_INV(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return _chi_inverse(context, p, freedom, upper=False)


@function("CHISQ.INV.RT", V, V)
def CHISQ_INV_RT(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return _chi_inverse(context, p, freedom, upper=True)


@function("CHIINV", V, V)
def CHIINV(context: Context, p: Scalar, freedom: Scalar) -> Value:
    return _chi_inverse(context, p, freedom, upper=True)


def _chi_test(context: Context, actual: Value, expected: Value) -> Value:
    seen = matrix(context, actual)
    wanted = matrix(context, expected)
    if seen.height != wanted.height or seen.width != wanted.width:
        return NA
    statistic = 0.0
    for observed, expect in zip(seen.items(), wanted.items(), strict=True):
        if isinstance(observed, CellError):
            return observed
        if isinstance(expect, CellError):
            return expect
        if not isinstance(observed, float) or not isinstance(expect, float):
            continue
        if expect == 0.0:
            return DIV0
        statistic += (observed - expect) ** 2 / expect
    rows, columns = seen.height, seen.width
    freedom = (rows - 1) * (columns - 1) if rows > 1 and columns > 1 else rows * columns - 1
    if freedom < 1:
        return NA
    with localcontext(special.CONTEXT):
        return _exact(special.gamma_upper(D(freedom) / 2, D(statistic) / 2))


@function("CHISQ.TEST", R, R)
def CHISQ_TEST(context: Context, actual: Value, expected: Value) -> Value:
    return _chi_test(context, actual, expected)


@function("CHITEST", R, R)
def CHITEST(context: Context, actual: Value, expected: Value) -> Value:
    return _chi_test(context, actual, expected)


def _gamma_arguments(context: Context, x: Scalar, alpha: Scalar, beta: Scalar) -> tuple[float, float, float]:
    value = context.number(x)
    shape = context.number(alpha)
    scale = context.number(beta)
    if value < 0 or shape <= 0 or scale <= 0:
        raise ExcelError(NUM)
    return value, shape, scale


def _gamma_dist(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    value, shape, scale = _gamma_arguments(context, x, alpha, beta)
    with localcontext(special.CONTEXT):
        if _flag(context, cumulative):
            return _exact(special.gamma_lower(D(shape), D(value) / D(scale)))
        return _exact(_gamma_pdf(D(value), D(shape), D(scale)))


@function("GAMMA.DIST", V, V, V, V)
def GAMMA_DIST(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    return _gamma_dist(context, x, alpha, beta, cumulative)


@function("GAMMADIST", V, V, V, V)
def GAMMADIST(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    return _gamma_dist(context, x, alpha, beta, cumulative)


def _gamma_inv(context: Context, p: Scalar, alpha: Scalar, beta: Scalar) -> Value:
    chance = context.number(p)
    shape = context.number(alpha)
    scale = context.number(beta)
    if not 0 <= chance < 1 or shape <= 0 or scale <= 0:
        return NUM
    if chance == 0:
        return 0.0
    return gamma_inverse(chance, shape, scale)


@function("GAMMA.INV", V, V, V)
def GAMMA_INV(context: Context, p: Scalar, alpha: Scalar, beta: Scalar) -> Value:
    return _gamma_inv(context, p, alpha, beta)


@function("GAMMAINV", V, V, V)
def GAMMAINV(context: Context, p: Scalar, alpha: Scalar, beta: Scalar) -> Value:
    return _gamma_inv(context, p, alpha, beta)


#: Stirling's series for ln(gamma), B_2k / (2k (2k-1)) for k = 1 to 10.
_STIRLING = (
    1 / 12, -1 / 360, 1 / 1260, -1 / 1680, 1 / 1188, -691 / 360360, 1 / 156, -3617 / 122400, 43867 / 244188,
    -174611 / 125400,
)  # fmt: skip
#: ln(sqrt(2 pi)) as a double.
_LN_ROOT_TWO_PI = 0.9189385332046728


def _nearest_log_gamma(x: float) -> float:
    with localcontext(special.CONTEXT):
        return float(special.lgamma(D(x)))


def _nearest_gamma(x: float) -> float:
    with localcontext(special.CONTEXT):
        return float(special.gamma(D(x)))


def log_gamma(x: float) -> float:
    """GAMMALN of a positive number, as Excel computes it, each step as the
    x87 gives it.

    Measured: from 8 up, Stirling's series to the tenth Bernoulli term,
    ``(x-0.5)*LN(x) - x + ln(sqrt(2 pi)) + series/x`` with the series
    summed in ``1/(x*x)``, to the bit on 1025 arguments from 8 to 2.5e305;
    past that ``(x-0.5)*LN(x)`` overflows and Excel answers ``#NUM!``. Above
    3 and below 8, GAMMALN of x less a whole number, in [2, 3), plus the log
    of the numbers between, multiplied from the bottom up: 772 of 772, a
    whole number reduced to 2, so GAMMALN(4) is LN(2*3) though GAMMALN(3),
    in Excel's own approximation, is a unit above LN(2). Below
    0.7, ``GAMMALN(x+1) - LN(x)``, 377 of 377. From 0.7 to 3 Excel has
    approximations of its own, none of the published ones tried (Cody and
    Hillstrom 1967, Cody 1988, Cephes, fdlibm); there this is the double
    nearest the exact value, a few units in the last place from Excel's.
    """
    if x >= 8.0:
        inverse_square = precise.divide(1.0, precise.multiply(x, x))
        series = _STIRLING[-1]
        for coefficient in reversed(_STIRLING[:-1]):
            series = precise.add(precise.multiply(series, inverse_square), coefficient)
        main = precise.subtract(precise.multiply(precise.subtract(x, 0.5), precise.ln(x)), x)
        return precise.add(precise.add(main, _LN_ROOT_TWO_PI), precise.divide(series, x))
    if x < 0.7:
        return precise.subtract(log_gamma(precise.add(x, 1.0)), precise.ln(x))
    if x > 3.0:
        steps = math.floor(x) - 2
        base = x - steps
        product = 1.0
        for step in range(steps):
            product = precise.multiply(product, base + step)
        return precise.add(_nearest_log_gamma(base), precise.ln(product))
    return _nearest_log_gamma(x)


def _gamma_small(x: float) -> float:
    """GAMMA for 0 < x < 10, measured: from 1 to 2 Excel's own approximation,
    here the nearest double; below 1, ``GAMMA(x+1) * (1/x)``, 100 of 100;
    from 2, the numbers from x-1 down to the base in (1, 2) multiplied
    from the top, then GAMMA of the base, 60 of 60 at each step to 10."""
    if x < 1.0:
        return precise.multiply(_gamma_small(precise.add(x, 1.0)), precise.divide(1.0, x))
    if x < 2.0:
        return _nearest_gamma(x)
    steps = math.floor(x) - 1
    base = x - steps
    product = 1.0
    for step in range(steps, 0, -1):
        product = precise.multiply(product, base + (step - 1))
    return precise.multiply(product, _nearest_gamma(base))


@function("GAMMA", V)
def GAMMA(context: Context, x: Scalar) -> Value:
    value = context.number(x)
    if value == 0 or (value < 0 and value.is_integer()):
        return NUM
    if 0 < value < 10:
        return checked(_gamma_small(value))
    with localcontext(special.CONTEXT):
        try:
            return _exact(special.gamma(D(value)))
        except OverflowError:
            return NUM


def _gamma_ln(context: Context, x: Scalar) -> Value:
    value = context.number(x)
    if value <= 0:
        return NUM
    return checked(log_gamma(value))


@function("GAMMALN", V)
def GAMMALN(context: Context, x: Scalar) -> Value:
    return _gamma_ln(context, x)


@function("GAMMALN.PRECISE", V)
def GAMMALN_PRECISE(context: Context, x: Scalar) -> Value:
    return _gamma_ln(context, x)


# ----------------------------------------------------------------------
# F and beta
# ----------------------------------------------------------------------


def _f_arguments(context: Context, x: Scalar, first: Scalar, second: Scalar) -> tuple[float, float, float]:
    value = context.number(x)
    a = _freedom(context, first)
    b = _freedom(context, second)
    if value < 0:
        raise ExcelError(NUM)
    return value, a, b


def _f_lower(x: Decimal, a: Decimal, b: Decimal) -> Decimal:
    return special.beta_lower(a / 2, b / 2, a * x / (a * x + b))


def _f_upper(x: Decimal, a: Decimal, b: Decimal) -> Decimal:
    return special.beta_lower(b / 2, a / 2, b / (b + a * x))


def _f_pdf(x: Decimal, a: Decimal, b: Decimal) -> Decimal:
    if x <= 0:
        return D(0)
    log = (
        a / 2 * special.ln(a)
        + b / 2 * special.ln(b)
        + (a / 2 - 1) * special.ln(x)
        - (a + b) / 2 * special.ln(a * x + b)
        - special.lbeta(a / 2, b / 2)
    )
    return special.exp(log)


@function("F.DIST", V, V, V, V)
def F_DIST(context: Context, x: Scalar, first: Scalar, second: Scalar, cumulative: Scalar) -> Value:
    value, a, b = _f_arguments(context, x, first, second)
    with localcontext(special.CONTEXT):
        if _flag(context, cumulative):
            return _exact(_f_lower(D(value), D(a), D(b)))
        return _exact(_f_pdf(D(value), D(a), D(b)))


@function("F.DIST.RT", V, V, V)
def F_DIST_RT(context: Context, x: Scalar, first: Scalar, second: Scalar) -> Value:
    value, a, b = _f_arguments(context, x, first, second)
    with localcontext(special.CONTEXT):
        return _exact(_f_upper(D(value), D(a), D(b)))


@function("FDIST", V, V, V)
def FDIST(context: Context, x: Scalar, first: Scalar, second: Scalar) -> Value:
    return F_DIST_RT(context, x, first, second)


def f_inverse(p: float, a: float, b: float, *, upper: bool) -> float:
    exact_a, exact_b = D(a), D(b)
    guess = max(b / (b - 2), 1.0) if b > 2 else 1.0
    if upper:
        return _invert(
            lambda x: _f_upper(x, exact_a, exact_b),
            lambda x: -_f_pdf(x, exact_a, exact_b),
            p,
            guess,
            0.0,
            None,
            decreasing=True,
        )
    return _invert(
        lambda x: _f_lower(x, exact_a, exact_b), lambda x: _f_pdf(x, exact_a, exact_b), p, guess, 0.0, None
    )


def _f_inv(context: Context, p: Scalar, first: Scalar, second: Scalar, *, upper: bool) -> Value:
    chance = context.number(p)
    a = _freedom(context, first)
    b = _freedom(context, second)
    if upper:
        if not 0 < chance <= 1:
            return NUM
        if chance == 1:
            return 0.0
    else:
        if not 0 <= chance < 1:
            return NUM
        if chance == 0:
            return 0.0
    return f_inverse(chance, a, b, upper=upper)


@function("F.INV", V, V, V)
def F_INV(context: Context, p: Scalar, first: Scalar, second: Scalar) -> Value:
    return _f_inv(context, p, first, second, upper=False)


@function("F.INV.RT", V, V, V)
def F_INV_RT(context: Context, p: Scalar, first: Scalar, second: Scalar) -> Value:
    return _f_inv(context, p, first, second, upper=True)


@function("FINV", V, V, V)
def FINV(context: Context, p: Scalar, first: Scalar, second: Scalar) -> Value:
    return _f_inv(context, p, first, second, upper=True)


def _f_test(context: Context, first: Value, second: Value) -> Value:
    a = numbers(context, (first,))
    b = numbers(context, (second,))
    if len(a) < 2 or len(b) < 2:
        return DIV0
    var_a = spread(a, sample=True)
    var_b = spread(b, sample=True)
    if isinstance(var_a, CellError) or isinstance(var_b, CellError) or var_a == 0 or var_b == 0:
        return DIV0
    ratio = var_a / var_b
    with localcontext(special.CONTEXT):
        lower = _f_lower(D(ratio), D(len(a) - 1), D(len(b) - 1))
        return _exact(2 * min(lower, 1 - lower))


@function("F.TEST", R, R)
def F_TEST(context: Context, first: Value, second: Value) -> Value:
    return _f_test(context, first, second)


@function("FTEST", R, R)
def FTEST(context: Context, first: Value, second: Value) -> Value:
    return _f_test(context, first, second)


def _beta_pdf(z: Decimal, a: Decimal, b: Decimal) -> Decimal:
    if z <= 0 or z >= 1:
        return D(0)
    return special.exp((a - 1) * special.ln(z) + (b - 1) * special.ln(1 - z) - special.lbeta(a, b))


def _beta_arguments(
    context: Context, alpha: Scalar, beta: Scalar, low: Scalar | None, high: Scalar | None
) -> tuple[float, float, float, float]:
    a = context.number(alpha)
    b = context.number(beta)
    bottom = _optional(context, low, 0.0)
    top = _optional(context, high, 1.0)
    if a <= 0 or b <= 0 or bottom >= top:
        raise ExcelError(NUM)
    return a, b, bottom, top


def _beta_dist(
    context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: bool, low: Scalar | None, high: Scalar | None
) -> Value:
    value = context.number(x)
    a, b, bottom, top = _beta_arguments(context, alpha, beta, low, high)
    if not bottom <= value <= top:
        return NUM
    with localcontext(special.CONTEXT):
        z = (D(value) - D(bottom)) / (D(top) - D(bottom))
        if cumulative:
            return _exact(special.beta_lower(D(a), D(b), z))
        return _exact(_beta_pdf(z, D(a), D(b)) / (D(top) - D(bottom)))


@function("BETA.DIST", V, V, V, V, V, V, minimum=4)
def BETA_DIST(
    context: Context,
    x: Scalar,
    alpha: Scalar,
    beta: Scalar,
    cumulative: Scalar,
    low: Scalar | None = None,
    high: Scalar | None = None,
) -> Value:
    return _beta_dist(context, x, alpha, beta, _flag(context, cumulative), low, high)


@function("BETADIST", V, V, V, V, V, minimum=3)
def BETADIST(
    context: Context, x: Scalar, alpha: Scalar, beta: Scalar, low: Scalar | None = None, high: Scalar | None = None
) -> Value:
    return _beta_dist(context, x, alpha, beta, True, low, high)


def _beta_inv(
    context: Context, p: Scalar, alpha: Scalar, beta: Scalar, low: Scalar | None, high: Scalar | None
) -> Value:
    chance = context.number(p)
    a, b, bottom, top = _beta_arguments(context, alpha, beta, low, high)
    if not 0 < chance < 1:
        return NUM
    exact_a, exact_b = D(a), D(b)
    z = _invert(
        lambda value: special.beta_lower(exact_a, exact_b, value),
        lambda value: _beta_pdf(value, exact_a, exact_b),
        chance,
        a / (a + b),
        0.0,
        1.0,
    )
    return checked(bottom + z * (top - bottom))


@function("BETA.INV", V, V, V, V, V, minimum=3)
def BETA_INV(
    context: Context, p: Scalar, alpha: Scalar, beta: Scalar, low: Scalar | None = None, high: Scalar | None = None
) -> Value:
    return _beta_inv(context, p, alpha, beta, low, high)


@function("BETAINV", V, V, V, V, V, minimum=3)
def BETAINV(
    context: Context, p: Scalar, alpha: Scalar, beta: Scalar, low: Scalar | None = None, high: Scalar | None = None
) -> Value:
    return _beta_inv(context, p, alpha, beta, low, high)


# ----------------------------------------------------------------------
# Lognormal, exponential, Weibull
# ----------------------------------------------------------------------


def _lognormal(context: Context, x: Scalar, mean: Scalar, deviation: Scalar, cumulative: bool) -> Value:
    value = context.number(x)
    centre = context.number(mean)
    spread_ = context.number(deviation)
    if value <= 0 or spread_ <= 0:
        return NUM
    if cumulative:
        return normal_cdf(precise.divide(precise.subtract(precise.ln(value), centre), spread_))
    z = (math.log(value) - centre) / spread_
    return checked(math.exp(-z * z / 2) / (value * spread_ * math.sqrt(2 * math.pi)))


@function("LOGNORM.DIST", V, V, V, V)
def LOGNORM_DIST(context: Context, x: Scalar, mean: Scalar, deviation: Scalar, cumulative: Scalar) -> Value:
    return _lognormal(context, x, mean, deviation, _flag(context, cumulative))


@function("LOGNORMDIST", V, V, V)
def LOGNORMDIST(context: Context, x: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    return _lognormal(context, x, mean, deviation, True)


def _lognormal_inverse(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    spread_ = context.number(deviation)
    if spread_ <= 0:
        return NUM
    return checked(math.exp(context.number(mean) + spread_ * normal_inverse(context.number(p))))


@function("LOGNORM.INV", V, V, V)
def LOGNORM_INV(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    return _lognormal_inverse(context, p, mean, deviation)


@function("LOGINV", V, V, V)
def LOGINV(context: Context, p: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    return _lognormal_inverse(context, p, mean, deviation)


def _exponential(context: Context, x: Scalar, rate: Scalar, cumulative: Scalar) -> Value:
    value = context.number(x)
    lam = context.number(rate)
    if value < 0 or lam <= 0:
        return NUM
    if _flag(context, cumulative):
        return checked(1 - math.exp(-lam * value))
    return checked(lam * math.exp(-lam * value))


@function("EXPON.DIST", V, V, V)
def EXPON_DIST(context: Context, x: Scalar, rate: Scalar, cumulative: Scalar) -> Value:
    return _exponential(context, x, rate, cumulative)


@function("EXPONDIST", V, V, V)
def EXPONDIST(context: Context, x: Scalar, rate: Scalar, cumulative: Scalar) -> Value:
    return _exponential(context, x, rate, cumulative)


def _power(base: float, exponent: float) -> float:
    found = power(base, exponent)
    if isinstance(found, CellError):
        raise ExcelError(found)
    assert isinstance(found, float)
    return found


def _raised(base: float, exponent: float) -> float:
    """``base ^ exponent`` as ``exp(exponent * ln(base))`` even for a whole
    exponent, each step as the x87 rounds it."""
    return precise.exp(precise.multiply(exponent, precise.ln(base)))


def _weibull(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    """Measured on 500 random arguments: the density is ``a / b^a``, times
    ``x^(a-1)``, times ``exp(-(x/b)^a)``, every power an exponential of a
    logarithm, 500 of 500. The distribution is ``1 - exp(-t)`` from
    ``t = ln 2`` up, all of them; below, ``-expm1(-t)`` is Excel's value 111
    times in 159 and a unit in the last place off the rest."""
    value = context.number(x)
    shape = context.number(alpha)
    scale = context.number(beta)
    if value < 0 or shape <= 0 or scale <= 0:
        return NUM
    if value == 0:
        if _flag(context, cumulative):
            return 0.0
        return checked(shape / scale) if shape == 1 else (0.0 if shape > 1 else NUM)
    stretched = _raised(precise.divide(value, scale), shape)
    if _flag(context, cumulative):
        if stretched >= _HALF_LIFE:
            return checked(precise.subtract(1.0, precise.exp(-stretched)))
        return checked(-precise.exp_less_one(-stretched))
    decay = precise.exp(-stretched)
    if decay == 0.0:
        return 0.0
    front = precise.divide(shape, _raised(scale, shape))
    return checked(precise.multiply(precise.multiply(front, _raised(value, precise.subtract(shape, 1.0))), decay))


@function("WEIBULL.DIST", V, V, V, V)
def WEIBULL_DIST(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    return _weibull(context, x, alpha, beta, cumulative)


@function("WEIBULL", V, V, V, V)
def WEIBULL(context: Context, x: Scalar, alpha: Scalar, beta: Scalar, cumulative: Scalar) -> Value:
    return _weibull(context, x, alpha, beta, cumulative)


# ----------------------------------------------------------------------
# Discrete distributions
# ----------------------------------------------------------------------


def _whole(context: Context, value: Scalar) -> int:
    return math.trunc(context.number(value))


def _binomial_arguments(context: Context, successes: Scalar, trials: Scalar, chance: Scalar) -> tuple[int, int, float]:
    k = _whole(context, successes)
    n = _whole(context, trials)
    p = context.number(chance)
    if n < 0 or k < 0 or k > n or not 0 <= p <= 1:
        raise ExcelError(NUM)
    return k, n, p


def _exact_power(base: Decimal, exponent: int) -> Decimal:
    """``base ** exponent``, with 0 to the 0 being 1, which Decimal will not work out: BINOM.DIST(1,1,1,TRUE) is
    answered in Excel (tests/fixtures/formula/cells_functions.json). pyOpenVBA's own (docs/formula_engine.md)."""
    return D(1) if exponent == 0 else base**exponent


def _binomial_exact(k: int, n: int, p: float) -> Decimal:
    exact_p = D(p)
    return D(math.comb(n, k)) * _exact_power(exact_p, k) * _exact_power(1 - exact_p, n - k)


def _binomial_density(k: int, n: int, p: float) -> float:
    """Measured: ``exp(ln(COMBIN(n,k)) + k*ln(p) + (n-k)*ln(1-p))``."""
    if p == 0:
        return 1.0 if k == 0 else 0.0
    if p == 1:
        return 1.0 if k == n else 0.0
    return checked(math.exp(math.log(math.comb(n, k)) + k * math.log(p) + (n - k) * math.log(1 - p)))


def _binomial(context: Context, successes: Scalar, trials: Scalar, chance: Scalar, cumulative: Scalar) -> Value:
    k, n, p = _binomial_arguments(context, successes, trials, chance)
    if not _flag(context, cumulative):
        return _binomial_density(k, n, p)
    with localcontext(special.CONTEXT):
        return _exact(sum((_binomial_exact(i, n, p) for i in range(k + 1)), D(0)))


@function("BINOM.DIST", V, V, V, V)
def BINOM_DIST(context: Context, successes: Scalar, trials: Scalar, chance: Scalar, cumulative: Scalar) -> Value:
    return _binomial(context, successes, trials, chance, cumulative)


@function("BINOMDIST", V, V, V, V)
def BINOMDIST(context: Context, successes: Scalar, trials: Scalar, chance: Scalar, cumulative: Scalar) -> Value:
    return _binomial(context, successes, trials, chance, cumulative)


@function("BINOM.DIST.RANGE", V, V, V, V, minimum=3)
def BINOM_DIST_RANGE(
    context: Context, trials: Scalar, chance: Scalar, first: Scalar, last: Scalar | None = None
) -> Value:
    n = _whole(context, trials)
    p = context.number(chance)
    low = _whole(context, first)
    high = low if last is None or isinstance(last, Empty) else _whole(context, last)
    if n < 0 or not 0 <= p <= 1 or low < 0 or low > n or high < low or high > n:
        return NUM
    with localcontext(special.CONTEXT):
        return _exact(sum((_binomial_exact(i, n, p) for i in range(low, high + 1)), D(0)))


def _binomial_inverse(context: Context, trials: Scalar, chance: Scalar, alpha: Scalar) -> Value:
    n = _whole(context, trials)
    p = context.number(chance)
    target = context.number(alpha)
    if n < 0 or not 0 <= p <= 1 or not 0 <= target <= 1:
        return NUM
    with localcontext(special.CONTEXT):
        total = D(0)
        wanted = D(target)
        for k in range(n + 1):
            total += _binomial_exact(k, n, p)
            if total >= wanted:
                return float(k)
    return float(n)


@function("BINOM.INV", V, V, V)
def BINOM_INV(context: Context, trials: Scalar, chance: Scalar, alpha: Scalar) -> Value:
    return _binomial_inverse(context, trials, chance, alpha)


@function("CRITBINOM", V, V, V)
def CRITBINOM(context: Context, trials: Scalar, chance: Scalar, alpha: Scalar) -> Value:
    return _binomial_inverse(context, trials, chance, alpha)


def _poisson_term(k: int, mean: float) -> float:
    """``exp(-mean) * mean^k / k!``, in that order: measured."""
    try:
        return math.exp(-mean) * _power(mean, k) / math.factorial(k)
    except (OverflowError, ExcelError):
        return math.exp(k * math.log(mean) - mean - math.lgamma(k + 1))


def _poisson(context: Context, x: Scalar, mean: Scalar, cumulative: Scalar) -> Value:
    k = _whole(context, x)
    lam = context.number(mean)
    if k < 0 or lam < 0:
        return NUM
    if not _flag(context, cumulative):
        return checked(_poisson_term(k, lam))
    total = 0.0
    for i in range(k + 1):
        total += _poisson_term(i, lam)
    return checked(total)


@function("POISSON.DIST", V, V, V)
def POISSON_DIST(context: Context, x: Scalar, mean: Scalar, cumulative: Scalar) -> Value:
    return _poisson(context, x, mean, cumulative)


@function("POISSON", V, V, V)
def POISSON(context: Context, x: Scalar, mean: Scalar, cumulative: Scalar) -> Value:
    return _poisson(context, x, mean, cumulative)


def _hypergeometric_exact(k: int, n: int, successes: int, population: int) -> Decimal:
    return D(math.comb(successes, k) * math.comb(population - successes, n - k)) / D(math.comb(population, n))


def _hypergeometric(
    context: Context, sample_s: Scalar, sample: Scalar, population_s: Scalar, population: Scalar, cumulative: bool
) -> Value:
    k = _whole(context, sample_s)
    n = _whole(context, sample)
    successes = _whole(context, population_s)
    size = _whole(context, population)
    if (
        k < 0
        or n <= 0
        or successes <= 0
        or size <= 0
        or k > n
        or k > successes
        or n > size
        or successes > size
        or n - k > size - successes
    ):
        return NUM
    with localcontext(special.CONTEXT):
        if not cumulative:
            return _exact(_hypergeometric_exact(k, n, successes, size))
        low = max(0, n - (size - successes))
        return _exact(sum((_hypergeometric_exact(i, n, successes, size) for i in range(low, k + 1)), D(0)))


@function("HYPGEOM.DIST", V, V, V, V, V)
def HYPGEOM_DIST(
    context: Context, sample_s: Scalar, sample: Scalar, population_s: Scalar, population: Scalar, cumulative: Scalar
) -> Value:
    return _hypergeometric(context, sample_s, sample, population_s, population, _flag(context, cumulative))


@function("HYPGEOMDIST", V, V, V, V)
def HYPGEOMDIST(context: Context, sample_s: Scalar, sample: Scalar, population_s: Scalar, population: Scalar) -> Value:
    return _hypergeometric(context, sample_s, sample, population_s, population, False)


def _negative_binomial(context: Context, failures: Scalar, successes: Scalar, chance: Scalar, cumulative: bool) -> Value:
    f = _whole(context, failures)
    s = _whole(context, successes)
    p = context.number(chance)
    if f < 0 or s < 1 or not 0 < p < 1:
        return NUM
    exact_p = D(p)
    with localcontext(special.CONTEXT):

        def term(i: int) -> Decimal:
            return D(math.comb(i + s - 1, s - 1)) * exact_p**s * (1 - exact_p) ** i

        if not cumulative:
            return _exact(term(f))
        return _exact(sum((term(i) for i in range(f + 1)), D(0)))


@function("NEGBINOM.DIST", V, V, V, V)
def NEGBINOM_DIST(context: Context, failures: Scalar, successes: Scalar, chance: Scalar, cumulative: Scalar) -> Value:
    return _negative_binomial(context, failures, successes, chance, _flag(context, cumulative))


@function("NEGBINOMDIST", V, V, V)
def NEGBINOMDIST(context: Context, failures: Scalar, successes: Scalar, chance: Scalar) -> Value:
    return _negative_binomial(context, failures, successes, chance, False)


# ----------------------------------------------------------------------
# Z.TEST, FISHER, counting
# ----------------------------------------------------------------------


def _z_test(context: Context, array: Value, mean: Scalar, deviation: Scalar | None) -> Value:
    values = numbers(context, (array,))
    centre = context.number(mean)
    count = len(values)
    if count == 0:
        return NA
    if deviation is None or isinstance(deviation, Empty):
        variance = spread(values, sample=True)
        if isinstance(variance, CellError):
            return variance
        sigma = math.sqrt(variance)
    else:
        sigma = context.number(deviation)
    if sigma == 0:
        return DIV0
    z = (precise.summed(values) / count - centre) / (sigma / math.sqrt(count))
    with localcontext(special.CONTEXT):
        return _exact(1 - special.normal_cdf(D(z)))


@function("Z.TEST", R, V, V, minimum=2)
def Z_TEST(context: Context, array: Value, mean: Scalar, deviation: Scalar | None = None) -> Value:
    return _z_test(context, array, mean, deviation)


@function("ZTEST", R, V, V, minimum=2)
def ZTEST(context: Context, array: Value, mean: Scalar, deviation: Scalar | None = None) -> Value:
    return _z_test(context, array, mean, deviation)


@function("FISHER", V)
def FISHER(context: Context, x: Scalar) -> Value:
    value = context.number(x)
    if not -1 < value < 1:
        return NUM
    return checked(0.5 * math.log((1 + value) / (1 - value)))


@function("FISHERINV", V)
def FISHERINV(context: Context, y: Scalar) -> Value:
    value = context.number(y)
    try:
        doubled = math.exp(2 * value)
    except OverflowError:
        return 1.0
    return checked((doubled - 1) / (doubled + 1))


@function("PERMUTATIONA", V, V)
def PERMUTATIONA(context: Context, number: Scalar, chosen: Scalar) -> Value:
    n = _whole(context, number)
    k = _whole(context, chosen)
    if n < 0 or k < 0:
        return NUM
    return checked(_power(float(n), float(k)))


@function("COMBINA", V, V)
def COMBINA(context: Context, number: Scalar, chosen: Scalar) -> Value:
    n = _whole(context, number)
    k = _whole(context, chosen)
    if n < 0 or k < 0 or (n < 1 and k > 0):
        return NUM
    if k == 0:
        return 1.0
    return checked(float(math.comb(n + k - 1, k)))


@function("MULTINOMIAL", R, maximum=255)
def MULTINOMIAL(context: Context, *args: Value) -> Value:
    """Measured: ``exp(GAMMALN(sum+1) - (GAMMALN(a+1) + GAMMALN(b+1) + ...))``,
    the denominators' logarithms summed first."""
    values = [math.trunc(value) for value in numbers(context, args)]
    if any(value < 0 for value in values):
        return NUM
    with localcontext(special.CONTEXT):
        top = float(special.lgamma(D(sum(values) + 1)))
        below = 0.0
        for value in values:
            below += float(special.lgamma(D(value + 1)))
    return checked(math.exp(top - below))


__all__ = ["normal_cdf", "normal_inverse", "t_cdf", "t_inverse"]
