"""Statistics: regression and correlation, percentiles and ranks, the shape
of a distribution, and FREQUENCY, MODE.MULT and PROB.

The paired functions, CORREL, SLOPE and the rest, read their two ranges
position by position and keep the positions where both hold numbers;
ranges of different sizes are ``#N/A``. They work in two passes: the means
first, then sums of deviations from them, each operation rounded as the
x87 rounds it. So do SKEW and KURT; VAR does not (see
:func:`.aggregate.spread`). Each formula here was measured on 300 random
samples.

PERCENTILE and QUARTILE interpolate between the two values around the rank
as ``low + fraction * (high - low)``, which is the form that gives Excel's
bits. LINEST fits by a Householder QR factorization of the centred data,
so its slope can differ from SLOPE's in the last place, as Excel's does.
"""

from __future__ import annotations

import math
from collections.abc import Iterable
from decimal import ROUND_DOWN

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.aggregate import deviations, mean
from pyopenvba.formula._calc.functions.arithmetic import checked, rounded
from pyopenvba.formula._calc.functions.common import matrix, numbers, pairs
from pyopenvba.formula._calc.numbers import total
from pyopenvba.formula._calc.precise import add, divide, multiply, square_root, subtract, summed
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import (
    DIV0,
    NA,
    NUM,
    REF,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Scalar,
    Value,
)
from pyopenvba.formula._calc.cells import CellError

# ----------------------------------------------------------------------
# Regression on one variable
# ----------------------------------------------------------------------


class _Moments:
    """The sums a simple regression needs, from (y, x) pairs: the means,
    then sums of products of deviations from them."""

    def __init__(self, found: list[tuple[float, float]]) -> None:
        self.count = count = len(found)
        if count == 0:
            raise ExcelError(DIV0)
        n = float(count)
        self.mean_y = divide(summed(y for y, _ in found), n)
        self.mean_x = divide(summed(x for _, x in found), n)
        across = [(subtract(y, self.mean_y), subtract(x, self.mean_x)) for y, x in found]
        self.yy = summed(multiply(dy, dy) for dy, _ in across)
        self.xx = summed(multiply(dx, dx) for _, dx in across)
        self.xy = summed(multiply(dy, dx) for dy, dx in across)

    @property
    def slope(self) -> float:
        if self.xx == 0:
            raise ExcelError(DIV0)
        return divide(self.xy, self.xx)

    def correlation(self, degrees: float) -> float:
        """The covariance over the product of the standard deviations, each
        of the sums over ``degrees``: n for CORREL, n - 1 for RSQ."""
        if self.xx == 0 or self.yy == 0 or degrees == 0:
            raise ExcelError(DIV0)
        spread = multiply(square_root(divide(self.yy, degrees)), square_root(divide(self.xx, degrees)))
        return divide(divide(self.xy, degrees), spread)


def _moments(context: Context, known_y: Value, known_x: Value) -> _Moments:
    return _Moments(pairs(context, known_y, known_x))


@function("CORREL", R, R)
def CORREL(context: Context, first: Value, second: Value) -> Value:
    moments = _moments(context, first, second)
    return checked(moments.correlation(float(moments.count)))


@function("PEARSON", R, R)
def PEARSON(context: Context, first: Value, second: Value) -> Value:
    moments = _moments(context, first, second)
    return checked(moments.correlation(float(moments.count)))


@function("RSQ", R, R)
def RSQ(context: Context, known_y: Value, known_x: Value) -> Value:
    """Measured: the square of a correlation taken over n - 1, where
    CORREL's is over n; the two differ in the last place."""
    moments = _moments(context, known_y, known_x)
    correlation = moments.correlation(float(moments.count - 1))
    return checked(multiply(correlation, correlation))


@function("SLOPE", R, R)
def SLOPE(context: Context, known_y: Value, known_x: Value) -> Value:
    return checked(_moments(context, known_y, known_x).slope)


@function("INTERCEPT", R, R)
def INTERCEPT(context: Context, known_y: Value, known_x: Value) -> Value:
    moments = _moments(context, known_y, known_x)
    return checked(subtract(moments.mean_y, multiply(moments.slope, moments.mean_x)))


@function("STEYX", R, R)
def STEYX(context: Context, known_y: Value, known_x: Value) -> Value:
    """Measured: ``sqrt((Σdy² - (Σdxdy)²/Σdx²) / (n - 2))``."""
    moments = _moments(context, known_y, known_x)
    if moments.count < 3:
        return DIV0
    if moments.xx == 0:
        return DIV0
    explained = divide(multiply(moments.xy, moments.xy), moments.xx)
    left = divide(subtract(moments.yy, explained), float(moments.count - 2))
    return checked(square_root(max(left, 0.0)))


def _forecast(context: Context, x: Scalar, known_y: Value, known_x: Value) -> Value:
    at = context.number(x)
    moments = _moments(context, known_y, known_x)
    # Measured: the mean plus the slope times the distance from the mean,
    # not the intercept plus the slope times x.
    return checked(add(moments.mean_y, multiply(moments.slope, subtract(at, moments.mean_x))))


@function("FORECAST", V, R, R)
def FORECAST(context: Context, x: Scalar, known_y: Value, known_x: Value) -> Value:
    return _forecast(context, x, known_y, known_x)


@function("FORECAST.LINEAR", V, R, R)
def FORECAST_LINEAR(context: Context, x: Scalar, known_y: Value, known_x: Value) -> Value:
    return _forecast(context, x, known_y, known_x)


def _covariance(context: Context, first: Value, second: Value, *, sample: bool) -> Value:
    moments = _moments(context, first, second)
    count = moments.count
    if sample and count < 2:
        return DIV0
    # Measured on 300 samples: the sum of products of deviations over n - 1
    # or over n, in two passes as CORREL takes it.
    return checked(divide(moments.xy, float(count - 1 if sample else count)))


@function("COVAR", R, R)
def COVAR(context: Context, first: Value, second: Value) -> Value:
    return _covariance(context, first, second, sample=False)


@function("COVARIANCE.P", R, R)
def COVARIANCE_P(context: Context, first: Value, second: Value) -> Value:
    return _covariance(context, first, second, sample=False)


@function("COVARIANCE.S", R, R)
def COVARIANCE_S(context: Context, first: Value, second: Value) -> Value:
    return _covariance(context, first, second, sample=True)


@function("STANDARDIZE", V, V, V)
def STANDARDIZE(context: Context, x: Scalar, mean: Scalar, deviation: Scalar) -> Value:
    spread = context.number(deviation)
    value = context.number(x)
    centre = context.number(mean)
    if spread <= 0:
        return NUM
    return checked((value - centre) / spread)


# ----------------------------------------------------------------------
# Percentiles and ranks
# ----------------------------------------------------------------------


def sorted_numbers(context: Context, array: Value) -> list[float]:
    return sorted(numbers(context, (array,)))


def interpolate(values: list[float], rank: float) -> float:
    """The value at a fractional position, counted from 0, in sorted values."""
    whole = math.floor(rank)
    fraction = rank - whole
    low = values[whole]
    if fraction == 0 or whole + 1 >= len(values):
        return low
    return low + fraction * (values[whole + 1] - low)


def percentile_inclusive(values: list[float], k: float) -> float:
    if not values or not 0 <= k <= 1:
        raise ExcelError(NUM)
    return interpolate(values, k * (len(values) - 1))


def percentile_exclusive(values: list[float], k: float) -> float:
    if not values or not 0 < k < 1:
        raise ExcelError(NUM)
    rank = k * (len(values) + 1) - 1
    if rank < 0 or rank > len(values) - 1:
        raise ExcelError(NUM)
    return interpolate(values, rank)


def quartile_inclusive(values: list[float], quart: float) -> float:
    which = math.trunc(quart)
    if not 0 <= which <= 4:
        raise ExcelError(NUM)
    return percentile_inclusive(values, which / 4)


def quartile_exclusive(values: list[float], quart: float) -> float:
    which = math.trunc(quart)
    if not 1 <= which <= 3:
        raise ExcelError(NUM)
    return percentile_exclusive(values, which / 4)


@function("PERCENTILE", R, V)
def PERCENTILE(context: Context, array: Value, k: Scalar) -> Value:
    return percentile_inclusive(sorted_numbers(context, array), context.number(k))


@function("PERCENTILE.INC", R, V)
def PERCENTILE_INC(context: Context, array: Value, k: Scalar) -> Value:
    return percentile_inclusive(sorted_numbers(context, array), context.number(k))


@function("PERCENTILE.EXC", R, V)
def PERCENTILE_EXC(context: Context, array: Value, k: Scalar) -> Value:
    return percentile_exclusive(sorted_numbers(context, array), context.number(k))


@function("QUARTILE", R, V)
def QUARTILE(context: Context, array: Value, quart: Scalar) -> Value:
    return quartile_inclusive(sorted_numbers(context, array), context.number(quart))


@function("QUARTILE.INC", R, V)
def QUARTILE_INC(context: Context, array: Value, quart: Scalar) -> Value:
    return quartile_inclusive(sorted_numbers(context, array), context.number(quart))


@function("QUARTILE.EXC", R, V)
def QUARTILE_EXC(context: Context, array: Value, quart: Scalar) -> Value:
    return quartile_exclusive(sorted_numbers(context, array), context.number(quart))


def _percent_rank(
    context: Context, array: Value, x: Scalar, significance: Scalar | None, *, exclusive: bool
) -> Value:
    """Where x falls among the values, cut down, not rounded, to the
    digits asked for: three unless said otherwise."""
    values = sorted_numbers(context, array)
    at = context.number(x)
    digits = 3 if significance is None or isinstance(significance, Empty) else context.integer(significance)
    if digits < 1:
        return NUM
    if not values:
        return NUM
    count = len(values)
    if at < values[0] or at > values[-1]:
        return NA
    below = sum(1 for value in values if value < at)
    if at in values:
        position = float(below + 1 if exclusive else below)
    else:
        low, high = values[below - 1], values[below]
        position = (below if exclusive else below - 1) + (at - low) / (high - low)
    scale = count + 1 if exclusive else count - 1
    if scale == 0:
        return 1.0
    return rounded(position / scale, digits, ROUND_DOWN)


@function("PERCENTRANK", R, V, V, minimum=2)
def PERCENTRANK(context: Context, array: Value, x: Scalar, significance: Scalar | None = None) -> Value:
    return _percent_rank(context, array, x, significance, exclusive=False)


@function("PERCENTRANK.INC", R, V, V, minimum=2)
def PERCENTRANK_INC(context: Context, array: Value, x: Scalar, significance: Scalar | None = None) -> Value:
    return _percent_rank(context, array, x, significance, exclusive=False)


@function("PERCENTRANK.EXC", R, V, V, minimum=2)
def PERCENTRANK_EXC(context: Context, array: Value, x: Scalar, significance: Scalar | None = None) -> Value:
    return _percent_rank(context, array, x, significance, exclusive=True)


@function("TRIMMEAN", R, V)
def TRIMMEAN(context: Context, array: Value, percent: Scalar) -> Value:
    values = sorted_numbers(context, array)
    share = context.number(percent)
    if not values or not 0 <= share < 1:
        return NUM
    # The points left out, cut down to an even number, half from each end.
    cut = math.floor(len(values) * share / 2)
    kept = values[cut : len(values) - cut]
    return checked(divide(total(kept), float(len(kept))))


# ----------------------------------------------------------------------
# The shape of a distribution
# ----------------------------------------------------------------------


def _standardized(values: list[float], *, sample: bool) -> list[float]:
    """Each value's distance from the mean in standard deviations, the
    deviation taken in two passes: measured, where VAR's one-pass sums
    give other bits."""
    centre = mean(values)
    deviation = square_root(divide(deviations(values), float(len(values) - 1 if sample else len(values))))
    if deviation == 0:
        raise ExcelError(DIV0)
    return [divide(subtract(value, centre), deviation) for value in values]


@function("KURT", R, maximum=255)
def KURT(context: Context, *args: Value) -> Value:
    """Measured on 300 samples: ``n(n+1)/((n-1)(n-2)(n-3)) * Σz⁴ -
    3(n-1)²/((n-2)(n-3))``, z the standardized values."""
    values = numbers(context, args)
    count = len(values)
    if count < 4:
        return DIV0
    fourth = summed(multiply(multiply(z, z), multiply(z, z)) for z in _standardized(values, sample=True))
    n = float(count)
    scale = divide(n * (n + 1), (n - 1) * (n - 2) * (n - 3))
    shift = divide(3 * (n - 1) * (n - 1), (n - 2) * (n - 3))
    return checked(subtract(multiply(scale, fourth), shift))


@function("SKEW", R, maximum=255)
def SKEW(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    count = len(values)
    if count < 3:
        return DIV0
    third = summed(multiply(multiply(z, z), z) for z in _standardized(values, sample=True))
    n = float(count)
    return checked(multiply(third, divide(n, (n - 1) * (n - 2))))


@function("SKEW.P", R, maximum=255)
def SKEW_P(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    if not values:
        return DIV0
    third = summed(multiply(multiply(z, z), z) for z in _standardized(values, sample=False))
    return checked(divide(third, float(len(values))))


# ----------------------------------------------------------------------
# FREQUENCY, MODE.MULT and PROB
# ----------------------------------------------------------------------


@function("FREQUENCY", R, R)
def FREQUENCY(context: Context, data: Value, bins: Value) -> Value:
    values = numbers(context, (data,))
    edges = numbers(context, (bins,))
    counts = [0] * (len(edges) + 1)
    order = sorted(range(len(edges)), key=lambda index: edges[index])
    for value in values:
        for index in order:
            if value <= edges[index]:
                counts[index] += 1
                break
        else:
            counts[-1] += 1
    return Array([[float(count)] for count in counts])


@function("MODE.MULT", R, maximum=255)
def MODE_MULT(context: Context, *args: Value) -> Value:
    values = numbers(context, args)
    counts: dict[float, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        return NA
    modes: list[float] = []
    for value in values:
        if counts[value] == best and value not in modes:
            modes.append(value)
    return Array([[value] for value in modes])


@function("PROB", R, R, V, V, minimum=3)
def PROB(context: Context, x_range: Value, prob_range: Value, lower: Scalar, upper: Scalar | None = None) -> Value:
    found = pairs(context, x_range, prob_range)
    low = context.number(lower)
    high = low if upper is None or isinstance(upper, Empty) else context.number(upper)
    if any(not 0 <= chance <= 1 for _, chance in found):
        return NUM
    if abs(summed(chance for _, chance in found) - 1) > 1e-10:
        return NUM
    return checked(summed(chance for value, chance in found if low <= value <= high))


# ----------------------------------------------------------------------
# LINEST and its relatives
# ----------------------------------------------------------------------


class _Fit:
    """A least-squares fit of y on one or more x variables.

    As KB 828533 describes and the corpus bears out: with a constant, the
    x and y values are centred on their means and the QR is taken of a
    column of ones followed by the centred x columns, the ones reflected
    first. Measured on crafted data whose sums are exact, that gives
    Excel's slope to the bit for two and three points, where a fit of the
    centred columns alone does not; on longer random data it matches a
    third of the time and is otherwise within a few units in the last
    place. Every sum runs in order as the x87 adds, never Python's
    compensated ``sum``."""

    def __init__(self, ys: list[float], xs: list[list[float]], constant: bool) -> None:
        count = len(ys)
        self.count = count
        self.constant = constant
        self.variables = len(xs)
        means = [divide(summed(column), float(count)) for column in xs] if constant else [0.0] * len(xs)
        mean_y = divide(summed(ys), float(count)) if constant else 0.0
        columns = [[subtract(value, mean) for value in column] for column, mean in zip(xs, means, strict=True)]
        target = [subtract(value, mean_y) for value in ys]
        if constant:
            coefficients, inverse, residual, kept = _least_squares([[1.0] * count, *columns], target)
            coefficients, inverse, kept = coefficients[1:], [row[1:] for row in inverse[1:]], kept[1:]
        else:
            coefficients, inverse, residual, kept = _least_squares(columns, target)
        self.coefficients = coefficients
        self.intercept = mean_y
        if constant:
            for coefficient, mean in zip(coefficients, means, strict=True):
                self.intercept = subtract(self.intercept, multiply(coefficient, mean))
        else:
            self.intercept = 0.0
        self.residual = residual
        self.total = summed(multiply(value, value) for value in target)
        self.kept = kept
        self.inverse = inverse
        self.means = means

    @property
    def freedom(self) -> int:
        return self.count - sum(self.kept) - (1 if self.constant else 0)

    def predict(self, row: list[float]) -> float:
        found = self.intercept
        for coefficient, x in zip(self.coefficients, row, strict=True):
            found = add(found, multiply(coefficient, x))
        return found


def _dot(a: list[float], b: list[float], start: int) -> float:
    return summed(multiply(a[i], b[i]) for i in range(start, len(a)))


def _least_squares(
    columns: list[list[float]], target: list[float]
) -> tuple[list[float], list[list[float]], float, list[bool]]:
    """Solve min ||A b - y|| by Householder QR. Returns the coefficients,
    the inverse of R (rows and columns of the kept variables), the residual
    sum of squares, and which variables were kept: one that is a
    combination of those before it gets a coefficient of 0."""
    a = [list(column) for column in columns]
    y = list(target)
    rows = len(y)
    kept: list[bool] = []
    order: list[int] = []
    step = 0
    for index, column in enumerate(a):
        size = square_root(_dot(column, column, 0))
        norm = square_root(_dot(column, column, step)) if step < rows else 0.0
        if step >= rows or norm <= size * 1e-10 or norm == 0:
            kept.append(False)
            continue
        kept.append(True)
        order.append(index)
        alpha = -norm if column[step] >= 0 else norm
        v = [0.0] * rows
        for i in range(step, rows):
            v[i] = column[i]
        v[step] = subtract(v[step], alpha)
        vv = _dot(v, v, step)
        for other in a[index:]:
            factor = divide(multiply(2.0, _dot(v, other, step)), vv)
            for i in range(step, rows):
                other[i] = subtract(other[i], multiply(factor, v[i]))
        factor = divide(multiply(2.0, _dot(v, y, step)), vv)
        for i in range(step, rows):
            y[i] = subtract(y[i], multiply(factor, v[i]))
        step += 1
    size = len(order)
    r = [[a[order[j]][i] if j >= i else 0.0 for j in range(size)] for i in range(size)]
    solution = [0.0] * size
    for i in reversed(range(size)):
        value = y[i]
        for j in range(i + 1, size):
            value = subtract(value, multiply(r[i][j], solution[j]))
        solution[i] = divide(value, r[i][i])
    inverse = [[0.0] * size for _ in range(size)]
    for column in range(size):
        for i in reversed(range(size)):
            value = 1.0 if i == column else 0.0
            for j in range(i + 1, size):
                value = subtract(value, multiply(r[i][j], inverse[j][column]))
            inverse[i][column] = divide(value, r[i][i])
    coefficients = [0.0] * len(a)
    for position, index in enumerate(order):
        coefficients[index] = solution[position]
    residual = _dot(y, y, step) if step < rows else 0.0
    return coefficients, inverse, residual, kept


def _variables(context: Context, known_y: Value, known_x: Value | None) -> tuple[list[float], list[list[float]], bool]:
    """The y values, the x variables as columns, and whether y runs down."""
    grid = matrix(context, known_y)
    down = grid.width == 1
    if grid.height != 1 and grid.width != 1:
        raise ExcelError(REF)
    ys = _numbers(grid.items())
    count = len(ys)
    if known_x is None or isinstance(known_x, Empty):
        return ys, [[float(index) for index in range(1, count + 1)]], down
    xs_grid = matrix(context, known_x)
    if xs_grid.height * xs_grid.width == count:
        return ys, [_numbers(xs_grid.items())], down
    if down and xs_grid.height == count:
        return ys, [_numbers(row[column] for row in xs_grid.rows) for column in range(xs_grid.width)], down
    if not down and xs_grid.width == count:
        return ys, [_numbers(row) for row in xs_grid.rows], down
    raise ExcelError(REF)


def _numbers(items: Iterable[Scalar]) -> list[float]:
    """Every item as a number: an error is itself, anything else not a
    number ``#VALUE!``."""
    return [_one(item) for item in items]


def _flag(context: Context, value: Scalar | None, default: bool) -> bool:
    if value is None or isinstance(value, Empty):
        return default
    return context.logical(value)


def _fit(
    context: Context, known_y: Value, known_x: Value | None, constant: Scalar | None, *, logarithm: bool
) -> tuple[_Fit, list[list[float]], bool]:
    """The fit, the x variables as columns, and whether y runs down."""
    ys, xs, down = _variables(context, known_y, known_x)
    if logarithm:
        if any(value <= 0 for value in ys):
            raise ExcelError(NUM)
        ys = [math.log(value) for value in ys]
    return _Fit(ys, xs, _flag(context, constant, True)), xs, down


def _statistics(fit: _Fit, first_row: list[Scalar], statistics: bool) -> Array:
    if not statistics:
        return Array([first_row])
    width = len(first_row)
    freedom = fit.freedom
    residual = fit.residual
    regression = subtract(fit.total, residual)
    kept = sum(fit.kept)
    variance = divide(residual, float(freedom)) if freedom > 0 else 0.0
    standard: Scalar = square_root(variance) if freedom > 0 else NUM
    errors: list[Scalar] = []
    diagonal: list[float] = []
    for row in fit.inverse:
        diagonal.append(summed(multiply(value, value) for value in row))
    position = 0
    per_variable: list[Scalar] = []
    # Each standard error is the root of the variance times the diagonal
    # of (X'X)^-1, rather than the standard error of y times its root.
    for keep in fit.kept:
        if keep:
            spread = diagonal[position]
            per_variable.append(square_root(multiply(variance, spread)) if isinstance(standard, float) else standard)
            position += 1
        else:
            per_variable.append(0.0)
    errors.extend(reversed(per_variable))
    if fit.constant and isinstance(standard, float):
        # The intercept's variance: 1/n plus the means through (X'X)^-1.
        means = [mean for mean, keep in zip(fit.means, fit.kept, strict=True) if keep]
        projected = [summed(multiply(fit.inverse[j][i], means[j]) for j in range(len(means))) for i in range(len(means))]
        spread = add(divide(1.0, float(fit.count)), summed(multiply(value, value) for value in projected))
        errors.append(square_root(multiply(variance, spread)))
    elif fit.constant:
        errors.append(standard)
    else:
        errors.append(NA)
    determination: Scalar = divide(regression, fit.total) if fit.total else NUM
    f_value: Scalar = (
        divide(divide(regression, float(kept)), divide(residual, float(freedom))) if freedom > 0 and kept and residual else NUM
    )
    pad: list[Scalar] = [NA] * (width - 2)
    return Array(
        [
            first_row,
            [checked(value) if isinstance(value, float) else value for value in errors],
            [determination, standard, *pad],
            [f_value, float(freedom), *pad],
            [regression, residual, *pad],
        ]
    )


@function("LINEST", R, R, V, V, minimum=1)
def LINEST(
    context: Context,
    known_y: Value,
    known_x: Value | None = None,
    constant: Scalar | None = None,
    stats: Scalar | None = None,
) -> Value:
    fit, _, _ = _fit(context, known_y, known_x, constant, logarithm=False)
    first: list[Scalar] = [checked(value) for value in reversed(fit.coefficients)]
    first.append(checked(fit.intercept))
    return _statistics(fit, first, _flag(context, stats, False))


@function("LOGEST", R, R, V, V, minimum=1)
def LOGEST(
    context: Context,
    known_y: Value,
    known_x: Value | None = None,
    constant: Scalar | None = None,
    stats: Scalar | None = None,
) -> Value:
    fit, _, _ = _fit(context, known_y, known_x, constant, logarithm=True)
    first: list[Scalar] = [checked(math.exp(value)) for value in reversed(fit.coefficients)]
    first.append(checked(math.exp(fit.intercept)))
    return _statistics(fit, first, _flag(context, stats, False))


def _projection(
    context: Context,
    known_y: Value,
    known_x: Value | None,
    new_x: Value | None,
    constant: Scalar | None,
    *,
    logarithm: bool,
) -> Value:
    """TREND and GROWTH: the fitted values at new x, by default the known
    ones. With one variable the result has new x's shape; with several,
    one value for each observation, running the way y does."""
    fit, xs, down = _fit(context, known_y, known_x, constant, logarithm=logarithm)
    if fit.variables == 1:
        if new_x is None or isinstance(new_x, Empty):
            new = Array([[value] for value in xs[0]]) if down else Array([list(xs[0])])
        else:
            new = matrix(context, new_x)
        return Array([[_predicted(fit, [_one(item)], logarithm) for item in row] for row in new.rows])
    if new_x is None or isinstance(new_x, Empty):
        observations = [list(row) for row in zip(*xs, strict=True)]
    else:
        grid = matrix(context, new_x)
        lines = grid.rows if down else [list(column) for column in zip(*grid.rows, strict=True)]
        observations = [_numbers(line) for line in lines]
    values = [_predicted(fit, row, logarithm) for row in observations]
    return Array([[value] for value in values]) if down else Array([values])


def _one(item: Scalar) -> float:
    if isinstance(item, CellError):
        raise ExcelError(item)
    if not isinstance(item, float):
        raise ExcelError(VALUE)
    return item


def _predicted(fit: _Fit, row: list[float], logarithm: bool) -> Scalar:
    value = fit.predict(row)
    return checked(math.exp(value) if logarithm else value)


@function("TREND", R, R, R, V, minimum=1)
def TREND(
    context: Context,
    known_y: Value,
    known_x: Value | None = None,
    new_x: Value | None = None,
    constant: Scalar | None = None,
) -> Value:
    return _projection(context, known_y, known_x, new_x, constant, logarithm=False)


@function("GROWTH", R, R, R, V, minimum=1)
def GROWTH(
    context: Context,
    known_y: Value,
    known_x: Value | None = None,
    new_x: Value | None = None,
    constant: Scalar | None = None,
) -> Value:
    return _projection(context, known_y, known_x, new_x, constant, logarithm=True)


__all__ = [
    "interpolate",
    "percentile_exclusive",
    "percentile_inclusive",
    "quartile_exclusive",
    "quartile_inclusive",
    "sorted_numbers",
]
