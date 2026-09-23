"""Financial functions: the time value of money, depreciation, internal
rates of return, and bonds, bills and discounted securities.

Measured: the time-value functions raise ``1 + rate`` to a power the way
the ``^`` operator does (see :func:`~.evaluator.power`), so a whole number
of periods is multiplied out square by square, and they are plain double
arithmetic: FV, PV, NPER, NPV and the rest give Excel's bits even where
those are a hundred units in the last place from the exact value. PMT is
the exception, being within one unit of exact, so it is computed exactly
and rounded once.

RATE, IRR, XIRR and YIELD solve by iteration, and Excel stops short of
the root. RATE and IRR follow Excel's own iterations step by step, so they
stop where it does, to the bit. XIRR, which Excel takes to within its
documented 0.000001 percent, and YIELD converge fully, so they agree with
Excel to that accuracy and not to the bit.
"""

from __future__ import annotations

import itertools
import math
from collections.abc import Callable
from decimal import ROUND_HALF_UP, Decimal, localcontext

from pyopenvba.formula._calc import dates, precise, special
from pyopenvba.formula._calc.evaluator import Context, power
from pyopenvba.formula._calc.functions.arithmetic import checked, rounded
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.functions.dates import basis, days_between, serial_of, year_fraction
from pyopenvba.formula._calc.registry import R, V, function
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


def _power(base: float, exponent: float) -> float:
    found = power(base, exponent)
    if isinstance(found, CellError):
        raise ExcelError(found)
    assert isinstance(found, float)
    return found


def _optional(context: Context, value: Scalar | None, default: float) -> float:
    return default if value is None or isinstance(value, Empty) else context.number(value)


def _due(context: Context, value: Scalar | None) -> int:
    """The payment timing: 0 at the end of each period, 1 at the start."""
    return 0 if value is None or isinstance(value, Empty) or context.number(value) == 0 else 1


# ----------------------------------------------------------------------
# The time value of money
# ----------------------------------------------------------------------


def _annuity(rate: float, periods: float, due: int) -> tuple[float, float]:
    """``(1 + rate) ^ periods`` and the annuity factor, ``(1 + rate * due)``
    times ``((1 + rate) ^ periods - 1) / rate``, as FV, PV and RATE all form
    them. Measured on 2,400 values of FV and PV where the terms cancel."""
    growth = _power(precise.add(1.0, rate), periods)
    spread = precise.divide(precise.subtract(growth, 1.0), rate)
    return growth, precise.multiply(precise.add(1.0, precise.multiply(rate, due)), spread)


def future_value(rate: float, periods: float, payment: float, present: float, due: int) -> float:
    if rate == 0:
        return -precise.add(present, precise.multiply(payment, periods))
    growth, factor = _annuity(rate, periods, due)
    return -precise.add(precise.multiply(present, growth), precise.multiply(payment, factor))


def present_value(rate: float, periods: float, payment: float, future: float, due: int) -> float:
    if rate == 0:
        return -precise.add(future, precise.multiply(payment, periods))
    growth, factor = _annuity(rate, periods, due)
    return -precise.divide(precise.add(future, precise.multiply(payment, factor)), growth)


def _balance(rate: float, periods: float, payment: float, present: float, future: float, due: int) -> float:
    """What RATE drives to zero: the future value of the loan and its
    payments, plus ``future``; FV's own sum with ``future`` added last."""
    if rate == 0:
        return precise.add(precise.add(precise.multiply(payment, periods), present), future)
    growth, factor = _annuity(rate, periods, due)
    return precise.add(precise.add(precise.multiply(present, growth), precise.multiply(payment, factor)), future)


def _growth(rate: Decimal, periods: float) -> Decimal:
    """``(1 + rate) ^ periods``, exactly to fifty digits."""
    base = 1 + rate
    if base <= 0 and periods != math.trunc(periods):
        raise ExcelError(NUM)
    if base == 0:
        return Decimal(0)
    return base ** Decimal(periods)


def _exact_payment(rate: float, periods: float, present: float, future: float, due: int) -> Decimal:
    if periods == 0:
        raise ExcelError(NUM)
    r = Decimal(rate)
    if r == 0:
        return -(Decimal(present) + Decimal(future)) / Decimal(periods)
    growth = _growth(r, periods)
    if growth == 1:
        raise ExcelError(NUM)
    return -(r * (Decimal(future) + Decimal(present) * growth)) / ((1 + r * due) * (growth - 1))


def _exact_future(rate: Decimal, periods: float, paid: Decimal, present: float, due: int) -> Decimal:
    if rate == 0:
        return -(Decimal(present) + paid * Decimal(periods))
    growth = _growth(rate, periods)
    return -(Decimal(present) * growth + paid * (1 + rate * due) * (growth - 1) / rate)


def _exact_interest(rate: float, period: float, periods: float, present: float, future: float, due: int) -> Decimal:
    """IPMT's interest for a period: the balance owed going into it, times
    the rate."""
    r = Decimal(rate)
    paid = _exact_payment(rate, periods, present, future, due)
    if period == 1:
        interest = Decimal(0) if due else -Decimal(present)
    elif due:
        interest = _exact_future(r, period - 2, paid, present, 1) - paid
    else:
        interest = _exact_future(r, period - 1, paid, present, 0)
    return interest * r


def payment(rate: float, periods: float, present: float, future: float, due: int) -> float:
    """The periodic payment, from the exact formula rounded once.

    Measured: PMT, IPMT, PPMT, CUMIPMT and CUMPRINC are all within a few
    units in the last place of the exact values, which no double formula
    reaches, so these five are computed exactly."""
    with localcontext(special.CONTEXT):
        return checked(float(_exact_payment(rate, periods, present, future, due)))


@function("FV", V, V, V, V, V, minimum=3)
def FV(
    context: Context, rate: Scalar, nper: Scalar, pmt: Scalar, pv: Scalar | None = None, type_: Scalar | None = None
) -> Value:
    return checked(
        future_value(
            context.number(rate), context.number(nper), context.number(pmt), _optional(context, pv, 0.0), _due(context, type_)
        )
    )


@function("PV", V, V, V, V, V, minimum=3)
def PV(
    context: Context, rate: Scalar, nper: Scalar, pmt: Scalar, fv: Scalar | None = None, type_: Scalar | None = None
) -> Value:
    return checked(
        present_value(
            context.number(rate), context.number(nper), context.number(pmt), _optional(context, fv, 0.0), _due(context, type_)
        )
    )


@function("PMT", V, V, V, V, V, minimum=3)
def PMT(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, fv: Scalar | None = None, type_: Scalar | None = None
) -> Value:
    return payment(
        context.number(rate), context.number(nper), context.number(pv), _optional(context, fv, 0.0), _due(context, type_)
    )


@function("NPER", V, V, V, V, V, minimum=3)
def NPER(
    context: Context, rate: Scalar, pmt: Scalar, pv: Scalar, fv: Scalar | None = None, type_: Scalar | None = None
) -> Value:
    r = context.number(rate)
    paid = context.number(pmt)
    present = context.number(pv)
    future = _optional(context, fv, 0.0)
    due = _due(context, type_)
    if r == 0:
        if paid == 0:
            return NUM
        return checked(-(present + future) / paid)
    if r <= -1:
        return NUM
    scaled = paid * (1 + r * due)
    top = scaled - future * r
    bottom = scaled + present * r
    if bottom == 0:
        return DIV0
    ratio = top / bottom
    if ratio <= 0:
        return NUM
    return checked(math.log(ratio) / math.log(1 + r))


def _period_arguments(
    context: Context, rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar, fv: Scalar | None, type_: Scalar | None
) -> tuple[float, float, float, float, float, int]:
    period = context.number(per)
    periods = context.number(nper)
    if period < 1 or period > periods:
        raise ExcelError(NUM)
    return context.number(rate), period, periods, context.number(pv), _optional(context, fv, 0.0), _due(context, type_)


@function("IPMT", V, V, V, V, V, V, minimum=4)
def IPMT(
    context: Context,
    rate: Scalar,
    per: Scalar,
    nper: Scalar,
    pv: Scalar,
    fv: Scalar | None = None,
    type_: Scalar | None = None,
) -> Value:
    with localcontext(special.CONTEXT):
        return checked(float(_exact_interest(*_period_arguments(context, rate, per, nper, pv, fv, type_))))


@function("PPMT", V, V, V, V, V, V, minimum=4)
def PPMT(
    context: Context,
    rate: Scalar,
    per: Scalar,
    nper: Scalar,
    pv: Scalar,
    fv: Scalar | None = None,
    type_: Scalar | None = None,
) -> Value:
    r, period, periods, present, future, due = _period_arguments(context, rate, per, nper, pv, fv, type_)
    with localcontext(special.CONTEXT):
        paid = _exact_payment(r, periods, present, future, due)
        return checked(float(paid - _exact_interest(r, period, periods, present, future, due)))


def _cumulative(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> tuple[float, float, int, int, int, float]:
    r = context.number(rate)
    periods = context.number(nper)
    present = context.number(pv)
    first = math.trunc(context.number(start))
    last = math.trunc(context.number(end))
    due = context.number(type_)
    if r <= 0 or periods <= 0 or present <= 0 or first < 1 or last < first or last > periods or due not in (0, 1):
        raise ExcelError(NUM)
    return r, periods, first, last, int(due), present


@function("CUMIPMT", V, V, V, V, V, V)
def CUMIPMT(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> Value:
    r, periods, first, last, due, present = _cumulative(context, rate, nper, pv, start, end, type_)
    with localcontext(special.CONTEXT):
        total = sum(
            (_exact_interest(r, period, periods, present, 0.0, due) for period in range(first, last + 1)), Decimal(0)
        )
        return checked(float(total))


@function("CUMPRINC", V, V, V, V, V, V)
def CUMPRINC(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> Value:
    r, periods, first, last, due, present = _cumulative(context, rate, nper, pv, start, end, type_)
    with localcontext(special.CONTEXT):
        paid = _exact_payment(r, periods, present, 0.0, due)
        total = sum(
            (paid - _exact_interest(r, period, periods, present, 0.0, due) for period in range(first, last + 1)),
            Decimal(0),
        )
        return checked(float(total))


def _cash_flows(context: Context, values: tuple[Value, ...]) -> list[float]:
    """NPV's values: the numbers of a range or array, and a value given on
    its own as a number, an empty argument as 0."""
    found: list[float] = []
    for value in values:
        if isinstance(value, (Reference, Array)):
            for item, _ in context.scalars(value):
                if isinstance(item, CellError):
                    raise ExcelError(item)
                if isinstance(item, float):
                    found.append(item)
        elif isinstance(value, Empty):
            found.append(0.0)
        else:
            found.append(context.number(value))
    return found


def _npv(rate: float, values: list[float]) -> float:
    """The values discounted from the end of each period: each divided by
    ``(1 + rate)^(i - 1)``, a running product, and the sum divided by
    ``1 + rate`` once at the end, each step as the x87 rounds it.
    Measured, 300 of 300 random series and every isolated period."""
    growth = precise.add(1.0, rate)
    total = 0.0
    discount = 1.0
    for index, value in enumerate(values):
        if index:
            discount = precise.multiply(discount, growth)
        total = precise.add(total, precise.divide(value, discount))
    return precise.divide(total, growth)


@function("NPV", V, R, minimum=2, maximum=255)
def NPV(context: Context, rate: Scalar, *values: Value) -> Value:
    r = context.number(rate)
    if r == -1:
        return DIV0
    return checked(_npv(r, _cash_flows(context, values)))


def _series(context: Context, values: Value, when: Value) -> tuple[list[float], list[int]]:
    """XNPV's and XIRR's payments and their days, the first day earliest."""
    amounts = [_number(item) for item in matrix(context, values).items()]
    days = [math.trunc(_number(item)) for item in matrix(context, when).items()]
    if len(amounts) != len(days) or not amounts:
        raise ExcelError(NUM)
    if any(day < days[0] for day in days):
        raise ExcelError(NUM)
    return amounts, days


def _number(item: Scalar) -> float:
    if isinstance(item, CellError):
        raise ExcelError(item)
    if not isinstance(item, float):
        raise ExcelError(VALUE)
    return item


def _xnpv(rate: float, amounts: list[float], days: list[int]) -> float:
    total = 0.0
    for amount, day in zip(amounts, days, strict=True):
        total += amount / _power(1 + rate, (day - days[0]) / 365)
    return total


@function("XNPV", V, R, R)
def XNPV(context: Context, rate: Scalar, values: Value, when: Value) -> Value:
    r = context.number(rate)
    amounts, days = _series(context, values, when)
    if r <= -1:
        return NUM
    return checked(_xnpv(r, amounts, days))


def _newton(
    function_: Callable[[float], float], slope: Callable[[float], float], guess: float, *, limit: int = 100
) -> float:
    """A root by Newton's method, run until it stops moving; ``#NUM!``
    when it does not settle or leaves the range a rate can have."""
    x = guess
    best, best_error = x, math.inf
    previous = math.inf
    for _ in range(limit):
        try:
            value = function_(x)
            gradient = slope(x)
        except (OverflowError, ZeroDivisionError):
            raise ExcelError(NUM) from None
        if abs(value) < best_error:
            best, best_error = x, abs(value)
        if value == 0:
            return x
        if gradient == 0 or not math.isfinite(gradient) or not math.isfinite(value):
            raise ExcelError(NUM)
        step = value / gradient
        new = x - step
        if not math.isfinite(new) or new <= -1:
            raise ExcelError(NUM)
        if new == x or abs(step) <= abs(new) * 1e-15:
            return new
        if abs(step) >= abs(previous) and abs(step) < 1e-10 * max(1.0, abs(x)):
            # The steps have stopped shrinking: rounding, not the root, is
            # moving it now.
            return best
        previous = step
        x = new
    raise ExcelError(NUM)


@function("XIRR", R, R, V, minimum=2)
def XIRR(context: Context, values: Value, when: Value, guess: Scalar | None = None) -> Value:
    amounts, days = _series(context, values, when)
    if len(amounts) == 1:
        # One payment is #N/A, whatever its sign: XIRR(1,1) (tests/fixtures/formula/). pyOpenVBA's own
        # (docs/formula_engine.md).
        return NA
    if not (any(amount > 0 for amount in amounts) and any(amount < 0 for amount in amounts)):
        return NUM
    start = _optional(context, guess, 0.1)
    years = [(day - days[0]) / 365 for day in days]

    def slope(rate: float) -> float:
        return precise.summed(-year * amount / (1 + rate) ** (year + 1) for amount, year in zip(amounts, years, strict=True))

    return checked(_newton(lambda rate: _xnpv(rate, amounts, days), slope, start))


@function("IRR", R, V, minimum=1)
def IRR(context: Context, values: Value, guess: Scalar | None = None) -> Value:
    flows = [item for item, _ in context.scalars(values) if isinstance(item, float)]
    for item, _ in context.scalars(values):
        if isinstance(item, CellError):
            return item
    if not (any(flow > 0 for flow in flows) and any(flow < 0 for flow in flows)):
        return NUM
    start = _optional(context, guess, 0.1)
    seen: list[tuple[float, float]] = []
    for begin in (start, 0.1) if start != 0.1 else (start,):
        found = _irr(flows, begin, seen)
        if found is not None:
            return checked(found)
    return checked(_irr_bracketed(flows, seen))


def _irr(flows: list[float], guess: float, seen: list[tuple[float, float]]) -> float | None:
    """IRR's root from one starting guess, found as Excel finds it, or None
    where Excel gives up on the guess.

    Measured: a secant iteration in the discount rate ``d = r / (1 + r)``,
    each point's residual the flows discounted by ``1 / (1 - d)`` (see
    :func:`_discounted`). It starts from ``guess / (1 + guess)`` and a second
    point 0.001 below it when the residual there is positive, above it
    otherwise, and stops at the first new point whose residual is under
    1e-7 in size, having moved less than 1e-7, giving ``1 / (1 - d) - 1``.
    A point whose residual equals the one before it, 200 steps, or two
    equal residuals to divide by, and Excel starts again from 0.1. Held to
    2,920 of 2,920 probes bit for bit, among them series built so their
    residual is exact and scaled so only an exact zero passes the 1e-7.
    ``seen`` collects each point and its residual.
    """
    try:
        a = precise.divide(guess, precise.add(1.0, guess))
        at_a = _discounted(a, flows)
        b = precise.add(a, -0.001 if at_a > 0 else 0.001)
        at_b = _discounted(b, flows)
        seen += [(a, at_a), (b, at_b)]
        for _ in range(_IRR_STEPS):
            c = precise.subtract(b, precise.divide(precise.multiply(at_b, precise.subtract(b, a)), precise.subtract(at_b, at_a)))
            at_c = _discounted(c, flows)
            seen.append((c, at_c))
            if abs(at_c) < _IRR_CLOSE and abs(precise.subtract(c, b)) < _IRR_CLOSE:
                return _rate_of(c)
            if at_c == at_b:
                return None
            a, at_a, b, at_b = b, at_b, c, at_c
    except (ZeroDivisionError, OverflowError):
        return None
    return None


def _irr_bracketed(flows: list[float], seen: list[tuple[float, float]]) -> float:
    """IRR's root where both secant runs gave up: halved down from the
    nearest two points they visited with residuals of opposite sign, and
    kept only if its residual passes the same 1e-7 test.

    Excel finds a root here too, a third way not yet pinned down, stopping
    short of it as the secant does: on 216 series built to throw the secant
    far from the root, this gives Excel's bits on 54 and the rest within
    1e-11 of Excel's answer. Where no point's residual can pass the test
    Excel gives #NUM!, as this does.
    """
    points = sorted((point for point in seen if math.isfinite(point[1])), key=lambda point: point[0])
    brackets = [(low, high) for low, high in itertools.pairwise(points) if (low[1] < 0) != (high[1] < 0)]
    if not brackets:
        raise ExcelError(NUM)
    (low, at_low), (high, at_high) = min(brackets, key=lambda pair: pair[1][0] - pair[0][0])
    while True:
        middle = (low + high) / 2
        if middle in (low, high):
            break
        at_middle = _discounted(middle, flows)
        if (at_middle < 0) == (at_low < 0):
            low, at_low = middle, at_middle
        else:
            high, at_high = middle, at_middle
    closest, at_closest = (low, at_low) if abs(at_low) <= abs(at_high) else (high, at_high)
    if not abs(at_closest) < _IRR_CLOSE:
        raise ExcelError(NUM)
    return _rate_of(closest)


def _rate_of(discount_rate: float) -> float:
    """The rate a discount rate stands for, ``1 / (1 - d) - 1``."""
    return precise.subtract(precise.divide(1.0, precise.subtract(1.0, discount_rate)), 1.0)


def _discounted(discount_rate: float, flows: list[float]) -> float:
    """The flows discounted at the discount rate ``d``: each divided by
    ``x^i`` for the growth ``x = 1 / (1 - d)``, the powers a running
    product. Measured on two flows scaled by 2^40, where only an exact zero
    passes the 1e-7: Excel's IRR is #NUM! exactly where no growth brings
    ``v1 / x`` to ``-v0``, 60 of 60."""
    growth = precise.divide(1.0, precise.subtract(1.0, discount_rate))
    total = 0.0
    discount = 1.0
    for index, value in enumerate(flows):
        if index:
            discount = precise.multiply(discount, growth)
        total = precise.add(total, precise.divide(value, discount))
    return total


#: IRR's step limit and its tolerance, measured.
_IRR_STEPS = 200
_IRR_CLOSE = 1e-7


@function("RATE", V, V, V, V, V, V, minimum=3)
def RATE(
    context: Context,
    nper: Scalar,
    pmt: Scalar,
    pv: Scalar,
    fv: Scalar | None = None,
    type_: Scalar | None = None,
    guess: Scalar | None = None,
) -> Value:
    periods = context.number(nper)
    paid = context.number(pmt)
    present = context.number(pv)
    future = _optional(context, fv, 0.0)
    due = _due(context, type_)
    start = _optional(context, guess, 0.1)
    if periods <= 0:
        return NUM
    return checked(_rate(periods, paid, present, future, due, start))


def _rate(periods: float, payment: float, present: float, future: float, due: int, guess: float) -> float:
    """RATE's root of :func:`_balance`, found as Excel finds it.

    Measured: a secant iteration, not Newton's method. It starts from the
    guess and a second point 0.001 below it when the balance there is
    positive, above it otherwise, and stops at the first new point whose
    balance is under 1e-7 in size, having moved less than 1e-7. A point
    whose balance equals the one before it replaces only that one. After
    200 steps, or when the points stop moving, the last is kept only if
    its balance is under 1e-4; a rate at or below -1 + 1e-7 is #NUM!.
    Held to 1,345 of 1,359 probes bit for bit, the rest within 3 units in
    the last place, where the balance itself is rounding noise.
    """

    def balance(rate: float) -> float:
        return _balance(rate, periods, payment, present, future, due)

    try:
        a = guess
        at_a = balance(a)
        b = precise.add(a, -0.001 if at_a > 0 else 0.001)
        at_b = balance(b)
        for _ in range(_RATE_STEPS):
            if at_b == at_a and a == b:
                break
            c = precise.subtract(b, precise.divide(precise.multiply(at_b, precise.subtract(b, a)), precise.subtract(at_b, at_a)))
            at_c = balance(c)
            if abs(at_c) < _RATE_CLOSE and abs(c - b) < _RATE_CLOSE:
                return _rate_found(c)
            if at_c == at_b:
                b = c
            else:
                a, at_a, b, at_b = b, at_b, c, at_c
    except (ZeroDivisionError, OverflowError):
        raise ExcelError(NUM) from None
    if not abs(at_b) < _RATE_LOOSE:
        raise ExcelError(NUM)
    return _rate_found(b)


def _rate_found(rate: float) -> float:
    if rate <= -1 + _RATE_CLOSE:
        raise ExcelError(NUM)
    return rate


#: RATE's step limit and its two tolerances, all measured.
_RATE_STEPS = 200
_RATE_CLOSE = 1e-7
_RATE_LOOSE = 1e-4


@function("MIRR", R, V, V)
def MIRR(context: Context, values: Value, finance: Scalar, reinvest: Scalar) -> Value:
    flows = [item for item, _ in context.scalars(values) if isinstance(item, float)]
    borrowed = context.number(finance)
    earned = context.number(reinvest)
    count = len(flows)
    positive = _npv(earned, [flow if flow > 0 else 0.0 for flow in flows])
    negative = _npv(borrowed, [flow if flow < 0 else 0.0 for flow in flows])
    if positive == 0 or negative == 0 or count < 2:
        return DIV0
    ratio = (-positive * _power(1 + earned, count)) / (negative * (1 + borrowed))
    return checked(_power(ratio, 1 / (count - 1)) - 1)


# ----------------------------------------------------------------------
# Rates
# ----------------------------------------------------------------------


@function("EFFECT", V, V)
def EFFECT(context: Context, nominal: Scalar, npery: Scalar) -> Value:
    rate = context.number(nominal)
    periods = math.trunc(context.number(npery))
    if rate <= 0 or periods < 1:
        return NUM
    return checked(_power(1 + rate / periods, periods) - 1)


@function("NOMINAL", V, V)
def NOMINAL(context: Context, effect: Scalar, npery: Scalar) -> Value:
    rate = context.number(effect)
    periods = math.trunc(context.number(npery))
    if rate <= 0 or periods < 1:
        return NUM
    return checked(periods * (_power(1 + rate, 1 / periods) - 1))


@function("FVSCHEDULE", V, R)
def FVSCHEDULE(context: Context, principal: Scalar, schedule: Value) -> Value:
    value = context.number(principal)
    for item in matrix(context, schedule).items():
        if isinstance(item, CellError):
            return item
        if isinstance(item, Empty):
            continue
        if not isinstance(item, float):
            return VALUE
        value *= 1 + item
    return checked(value)


@function("PDURATION", V, V, V)
def PDURATION(context: Context, rate: Scalar, pv: Scalar, fv: Scalar) -> Value:
    r = context.number(rate)
    present = context.number(pv)
    future = context.number(fv)
    if r <= 0 or present <= 0 or future <= 0:
        return NUM
    return checked((math.log(future) - math.log(present)) / math.log(1 + r))


@function("RRI", V, V, V)
def RRI(context: Context, nper: Scalar, pv: Scalar, fv: Scalar) -> Value:
    periods = context.number(nper)
    present = context.number(pv)
    future = context.number(fv)
    if periods <= 0:
        return NUM
    if present == 0:
        return NUM
    return checked(_power(future / present, 1 / periods) - 1)


@function("ISPMT", V, V, V, V)
def ISPMT(context: Context, rate: Scalar, per: Scalar, nper: Scalar, pv: Scalar) -> Value:
    periods = context.number(nper)
    if periods == 0:
        return DIV0
    return checked(context.number(pv) * context.number(rate) * (context.number(per) / periods - 1))


def _fraction_digits(context: Context, fraction: Scalar) -> tuple[int, int]:
    whole = math.trunc(context.number(fraction))
    if whole < 0:
        raise ExcelError(NUM)
    if whole == 0:
        raise ExcelError(DIV0)
    return whole, math.ceil(math.log10(whole))


@function("DOLLARDE", V, V)
def DOLLARDE(context: Context, dollar: Scalar, fraction: Scalar) -> Value:
    value = context.number(dollar)
    whole, digits = _fraction_digits(context, fraction)
    integer = math.trunc(value)
    return checked(integer + (value - integer) * 10**digits / whole)


@function("DOLLARFR", V, V)
def DOLLARFR(context: Context, dollar: Scalar, fraction: Scalar) -> Value:
    value = context.number(dollar)
    whole, digits = _fraction_digits(context, fraction)
    integer = math.trunc(value)
    return checked(integer + (value - integer) * whole / 10**digits)


# ----------------------------------------------------------------------
# Depreciation
# ----------------------------------------------------------------------


@function("SLN", V, V, V)
def SLN(context: Context, cost: Scalar, salvage: Scalar, life: Scalar) -> Value:
    span = context.number(life)
    if span == 0:
        return DIV0
    return checked((context.number(cost) - context.number(salvage)) / span)


@function("SYD", V, V, V, V)
def SYD(context: Context, cost: Scalar, salvage: Scalar, life: Scalar, per: Scalar) -> Value:
    span = context.number(life)
    period = context.number(per)
    if span <= 0 or period <= 0 or period > span:
        return NUM
    return checked((context.number(cost) - context.number(salvage)) * (span - period + 1) * 2 / (span * (span + 1)))


@function("DB", V, V, V, V, V, minimum=4)
def DB(
    context: Context, cost: Scalar, salvage: Scalar, life: Scalar, period: Scalar, month: Scalar | None = None
) -> Value:
    value = context.number(cost)
    rest = context.number(salvage)
    span = context.number(life)
    which = context.number(period)
    months = _optional(context, month, 12.0)
    if value < 0 or rest < 0 or span <= 0 or which <= 0 or not 1 <= months <= 12 or which > span + 1:
        return NUM
    if value == 0:
        return 0.0
    rate = rounded(1 - _power(rest / value, 1 / span), 3, ROUND_HALF_UP)
    # Measured: the book value carried from period to period, and the part
    # of a year left at the end as a fraction of twelve months.
    depreciation = value * rate * months / 12
    if math.trunc(which) == 1:
        return checked(depreciation)
    book = value - depreciation
    for _ in range(2, math.trunc(min(span, which)) + 1):
        depreciation = book * rate
        book -= depreciation
    if which > span:
        depreciation = book * rate * ((12 - months) / 12)
    return checked(depreciation)


def _declining(cost: float, salvage: float, life: float, period: float, factor: float) -> float:
    """DDB's depreciation in one period, as Excel documents it: the book
    value going in times the rate, but no further than the salvage."""
    rate = min(factor / life, 1.0)
    before = cost * _power(1 - rate, period - 1) if rate < 1 else (cost if period == 1 else 0.0)
    depreciation = before * rate
    if before - depreciation < salvage:
        depreciation = before - salvage
    return max(depreciation, 0.0)


@function("DDB", V, V, V, V, V, minimum=4)
def DDB(
    context: Context, cost: Scalar, salvage: Scalar, life: Scalar, period: Scalar, factor: Scalar | None = None
) -> Value:
    value = context.number(cost)
    rest = context.number(salvage)
    span = context.number(life)
    which = context.number(period)
    rate = _optional(context, factor, 2.0)
    if value < 0 or rest < 0 or span <= 0 or which <= 0 or which > span or rate <= 0:
        return NUM
    return checked(_declining(value, rest, span, which, rate))


def _switching(cost: float, salvage: float, life: float, first: float, last: float, factor: float) -> float:
    """VDB with the switch to straight line: the book value carried from
    the first period, each period's declining balance compared with the
    straight line over what is left, and a period only partly inside the
    span counted in proportion. Measured: this carried value, not DDB's
    power, gives Excel's bits here."""
    rate = min(factor / life, 1.0)
    book = cost
    total = 0.0
    switched = False
    straight = 0.0
    for index in range(1, math.ceil(last) + 1):
        declining = book * rate
        if book - declining < salvage:
            declining = book - salvage
        declining = max(declining, 0.0)
        if not switched and life - index + 1 > 0:
            candidate = (book - salvage) / (life - index + 1)
            if candidate > declining:
                switched = True
                straight = candidate
        term = straight if switched else declining
        share = min(last, index) - max(first, index - 1)
        if share > 0:
            total += term if share == 1 else term * share
        book -= term
    return total


@function("VDB", V, V, V, V, V, V, V, minimum=5)
def VDB(
    context: Context,
    cost: Scalar,
    salvage: Scalar,
    life: Scalar,
    start: Scalar,
    end: Scalar,
    factor: Scalar | None = None,
    no_switch: Scalar | None = None,
) -> Value:
    value = context.number(cost)
    rest = context.number(salvage)
    span = context.number(life)
    first = context.number(start)
    last = context.number(end)
    rate = _optional(context, factor, 2.0)
    keep = no_switch is not None and not isinstance(no_switch, Empty) and context.logical(no_switch)
    if first < 0 or last < first or last > span or value < 0 or rest > value or rate <= 0:
        return NUM
    low = math.floor(first)
    high = math.ceil(last)
    if keep:
        total = 0.0
        for index in range(low + 1, high + 1):
            term = _declining(value, rest, span, index, rate)
            if index == low + 1:
                term *= min(last, low + 1) - first
            elif index == high:
                term *= last + 1 - high
            total += term
        return checked(total)
    return checked(_switching(value, rest, span, first, last, rate))


# ----------------------------------------------------------------------
# Bonds, bills and discounted securities
# ----------------------------------------------------------------------


def _frequency(context: Context, value: Scalar) -> int:
    found = math.trunc(context.number(value))
    if found not in (1, 2, 4):
        raise ExcelError(NUM)
    return found


def _dated(context: Context, settlement: Scalar, maturity: Scalar) -> tuple[int, int]:
    first = serial_of(context, settlement)
    last = serial_of(context, maturity)
    if first >= last:
        raise ExcelError(NUM)
    return first, last


def _months_back(context: Context, day: int, months: int, anchor_day: int, month_end: bool) -> int:
    """A coupon date ``months`` before (or after, when negative) ``day``,
    keeping the maturity's day of month, or the month's last day."""
    year, month, _ = dates.calendar(day, epoch_1904=context.epoch_1904)
    count = year * 12 + (month - 1) - months
    year, month = divmod(count, 12)
    month += 1
    last = dates.days_in_month(year, month)
    return dates.serial(year, month, last if month_end else min(anchor_day, last), epoch_1904=context.epoch_1904)


def _coupons(context: Context, settlement: int, maturity: int, frequency: int) -> tuple[int, int]:
    """The coupon dates either side of settlement: counted back from the
    maturity a period at a time, month ends kept as month ends."""
    year, month, day = dates.calendar(maturity, epoch_1904=context.epoch_1904)
    month_end = day == dates.days_in_month(year, month)
    step = 12 // frequency
    settle_year = dates.calendar(settlement, epoch_1904=context.epoch_1904)[0]
    # Start at the maturity's day in the settlement's year, then walk back.
    candidate = _months_back(context, maturity, (year - settle_year) * 12, day, month_end)
    if candidate <= settlement:
        candidate = _months_back(context, candidate, -12, day, month_end)
    while candidate > settlement:
        candidate = _months_back(context, candidate, step, day, month_end)
    return candidate, _months_back(context, candidate, -step, day, month_end)


def _coupon_count(context: Context, settlement: int, maturity: int, frequency: int) -> int:
    previous, _ = _coupons(context, settlement, maturity, frequency)
    y1, m1, _ = dates.calendar(previous, epoch_1904=context.epoch_1904)
    y2, m2, _ = dates.calendar(maturity, epoch_1904=context.epoch_1904)
    return ((y2 - y1) * 12 + (m2 - m1)) * frequency // 12


def _period_days(context: Context, settlement: int, maturity: int, frequency: int, kind: int) -> float:
    """COUPDAYS: the days of the coupon period holding settlement."""
    if kind == 1:
        previous, following = _coupons(context, settlement, maturity, frequency)
        return float(following - previous)
    return (365.0 if kind == 3 else 360.0) / frequency


def _days_before(context: Context, settlement: int, maturity: int, frequency: int, kind: int) -> float:
    """COUPDAYBS: the days from the last coupon to settlement."""
    previous, _ = _coupons(context, settlement, maturity, frequency)
    return float(days_between(previous, settlement, kind, epoch_1904=context.epoch_1904))


def _days_after(context: Context, settlement: int, maturity: int, frequency: int, kind: int) -> float:
    """COUPDAYSNC: the days from settlement to the next coupon; for the
    30/360 bases, what is left of the period."""
    if kind in (0, 4):
        return _period_days(context, settlement, maturity, frequency, kind) - _days_before(
            context, settlement, maturity, frequency, kind
        )
    _, following = _coupons(context, settlement, maturity, frequency)
    return float(following - settlement)


def _coupon_arguments(
    context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None
) -> tuple[int, int, int, int]:
    first, last = _dated(context, settlement, maturity)
    return first, last, _frequency(context, frequency), basis(context, basis_)


@function("COUPDAYS", V, V, V, V, minimum=3)
def COUPDAYS(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    return _period_days(context, *_coupon_arguments(context, settlement, maturity, frequency, basis_))


@function("COUPDAYBS", V, V, V, V, minimum=3)
def COUPDAYBS(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    return _days_before(context, *_coupon_arguments(context, settlement, maturity, frequency, basis_))


@function("COUPDAYSNC", V, V, V, V, minimum=3)
def COUPDAYSNC(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    return _days_after(context, *_coupon_arguments(context, settlement, maturity, frequency, basis_))


@function("COUPNCD", V, V, V, V, minimum=3)
def COUPNCD(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    first, last, times, _ = _coupon_arguments(context, settlement, maturity, frequency, basis_)
    return float(_coupons(context, first, last, times)[1])


@function("COUPPCD", V, V, V, V, minimum=3)
def COUPPCD(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    first, last, times, _ = _coupon_arguments(context, settlement, maturity, frequency, basis_)
    return float(_coupons(context, first, last, times)[0])


@function("COUPNUM", V, V, V, V, minimum=3)
def COUPNUM(context: Context, settlement: Scalar, maturity: Scalar, frequency: Scalar, basis_: Scalar | None = None) -> Value:
    first, last, times, _ = _coupon_arguments(context, settlement, maturity, frequency, basis_)
    return float(_coupon_count(context, first, last, times))


def _price(
    context: Context, settlement: int, maturity: int, rate: float, yld: float, redemption: float, frequency: int, kind: int
) -> float:
    """A bond's price per 100 of face value at a yield."""
    period = _period_days(context, settlement, maturity, frequency, kind)
    remaining = _days_after(context, settlement, maturity, frequency, kind) / period
    count = _coupon_count(context, settlement, maturity, frequency)
    accrued = _days_before(context, settlement, maturity, frequency, kind)
    coupon = 100 * rate / frequency
    if count == 1:
        return (redemption + coupon) / (1 + remaining * yld / frequency) - coupon * accrued / period
    growth = 1 + yld / frequency
    price = redemption / _power(growth, count - 1 + remaining)
    price -= coupon * accrued / period
    for index in range(count):
        price += coupon / _power(growth, index + remaining)
    return price


def _bond(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    rate: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None,
) -> tuple[int, int, float, float, int, int]:
    first, last = _dated(context, settlement, maturity)
    coupon_rate = context.number(rate)
    value = context.number(redemption)
    times = _frequency(context, frequency)
    kind = basis(context, basis_)
    if coupon_rate < 0 or value <= 0:
        raise ExcelError(NUM)
    return first, last, coupon_rate, value, times, kind


@function("PRICE", V, V, V, V, V, V, V, minimum=6)
def PRICE(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    first, last, coupon_rate, value, times, kind = _bond(context, settlement, maturity, rate, redemption, frequency, basis_)
    y = context.number(yld)
    if y < 0:
        return NUM
    return checked(_price(context, first, last, coupon_rate, y, value, times, kind))


@function("YIELD", V, V, V, V, V, V, V, minimum=6)
def YIELD(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    first, last, coupon_rate, value, times, kind = _bond(context, settlement, maturity, rate, redemption, frequency, basis_)
    target = context.number(pr)
    if target <= 0:
        return NUM
    count = _coupon_count(context, first, last, times)
    if count == 1:
        period = _period_days(context, first, last, times, kind)
        accrued = _days_before(context, first, last, times, kind)
        remaining = _days_after(context, first, last, times, kind)
        paid = target / 100 + accrued / period * coupon_rate / times
        return checked((value / 100 + coupon_rate / times - paid) / paid * (times * period / remaining))

    def gap(y: float) -> float:
        return _price(context, first, last, coupon_rate, y, value, times, kind) - target

    def slope(y: float) -> float:
        step = max(abs(y), 1e-4) * 1e-7
        return (gap(y + step) - gap(y - step)) / (2 * step)

    return checked(_newton(gap, slope, coupon_rate or 0.05))


def _duration(
    context: Context, settlement: int, maturity: int, coupon: float, yld: float, frequency: int, kind: int
) -> float:
    """Macaulay duration, in years: each payment's time, counted in coupon
    periods from settlement, weighted by its present value."""
    period = _period_days(context, settlement, maturity, frequency, kind)
    offset = _days_after(context, settlement, maturity, frequency, kind) / period - 1
    count = _coupon_count(context, settlement, maturity, frequency)
    paid = coupon * 100 / frequency
    growth = yld / frequency + 1
    # Measured: the coupons summed first and the last payment added after,
    # the weighted sum and the plain one each on its own.
    weighted = 0.0
    for time in range(1, count):
        weighted += (time + offset) * paid / _power(growth, time + offset)
    weighted += (count + offset) * (paid + 100) / _power(growth, count + offset)
    value = 0.0
    for time in range(1, count):
        value += paid / _power(growth, time + offset)
    value += (paid + 100) / _power(growth, count + offset)
    return weighted / value / frequency


@function("DURATION", V, V, V, V, V, V, minimum=5)
def DURATION(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    coupon: Scalar,
    yld: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    first, last = _dated(context, settlement, maturity)
    rate = context.number(coupon)
    y = context.number(yld)
    if rate < 0 or y < 0:
        return NUM
    return checked(_duration(context, first, last, rate, y, _frequency(context, frequency), basis(context, basis_)))


@function("MDURATION", V, V, V, V, V, V, minimum=5)
def MDURATION(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    coupon: Scalar,
    yld: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    first, last = _dated(context, settlement, maturity)
    rate = context.number(coupon)
    y = context.number(yld)
    if rate < 0 or y < 0:
        return NUM
    times = _frequency(context, frequency)
    # Measured: times the reciprocal, not divided.
    return checked(_duration(context, first, last, rate, y, times, basis(context, basis_)) * (1 / (1 + y / times)))


@function("ACCRINT", V, V, V, V, V, V, V, V, minimum=6)
def ACCRINT(
    context: Context,
    issue: Scalar,
    first_interest: Scalar,
    settlement: Scalar,
    rate: Scalar,
    par: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
    method: Scalar | None = None,
) -> Value:
    start = serial_of(context, issue)
    serial_of(context, first_interest)
    end = serial_of(context, settlement)
    coupon_rate = context.number(rate)
    value = context.number(par)
    _frequency(context, frequency)
    kind = basis(context, basis_)
    if coupon_rate <= 0 or value <= 0 or start >= end:
        return NUM
    return checked(value * coupon_rate * year_fraction(start, end, kind, epoch_1904=context.epoch_1904))


@function("ACCRINTM", V, V, V, V, V, minimum=4)
def ACCRINTM(
    context: Context, issue: Scalar, settlement: Scalar, rate: Scalar, par: Scalar, basis_: Scalar | None = None
) -> Value:
    start = serial_of(context, issue)
    end = serial_of(context, settlement)
    coupon_rate = context.number(rate)
    value = context.number(par)
    kind = basis(context, basis_)
    # Settled the day it is issued, it has accrued nothing: ACCRINTM(1,1,1,1) is 0 (tests/fixtures/formula/).
    # pyOpenVBA's own (docs/formula_engine.md).
    if coupon_rate <= 0 or value <= 0 or start > end:
        return NUM
    return checked(value * coupon_rate * year_fraction(start, end, kind, epoch_1904=context.epoch_1904))


def _term(context: Context, settlement: Scalar, maturity: Scalar, basis_: Scalar | None) -> float:
    """The years from settlement to maturity, as the basis counts them."""
    first, last = _dated(context, settlement, maturity)
    return year_fraction(first, last, basis(context, basis_), epoch_1904=context.epoch_1904)


def _day_count(context: Context, first: int, last: int, kind: int) -> tuple[float, float]:
    """The days from one date to a later one and the days in a year, as a
    basis counts them; for actual/actual, the year YEARFRAC divides by."""
    epoch = context.epoch_1904
    if kind == 1:
        days = float(last - first)
        return days, days / year_fraction(first, last, 1, epoch_1904=epoch)
    return float(days_between(first, last, kind, epoch_1904=epoch)), 365.0 if kind == 3 else 360.0


def _span(context: Context, settlement: Scalar, maturity: Scalar, basis_: Scalar | None) -> tuple[float, float]:
    first, last = _dated(context, settlement, maturity)
    return _day_count(context, first, last, basis(context, basis_))


@function("DISC", V, V, V, V, V, minimum=4)
def DISC(context: Context, settlement: Scalar, maturity: Scalar, pr: Scalar, redemption: Scalar, basis_: Scalar | None = None) -> Value:
    price = context.number(pr)
    value = context.number(redemption)
    if price <= 0 or value <= 0:
        return NUM
    return checked((1 - price / value) / _term(context, settlement, maturity, basis_))


@function("INTRATE", V, V, V, V, V, minimum=4)
def INTRATE(
    context: Context, settlement: Scalar, maturity: Scalar, investment: Scalar, redemption: Scalar, basis_: Scalar | None = None
) -> Value:
    invested = context.number(investment)
    value = context.number(redemption)
    if invested <= 0 or value <= 0:
        return NUM
    days, year = _span(context, settlement, maturity, basis_)
    return checked((value - invested) / invested * year / days)


@function("RECEIVED", V, V, V, V, V, minimum=4)
def RECEIVED(
    context: Context, settlement: Scalar, maturity: Scalar, investment: Scalar, discount: Scalar, basis_: Scalar | None = None
) -> Value:
    invested = context.number(investment)
    rate = context.number(discount)
    if invested <= 0 or rate <= 0:
        return NUM
    days, year = _span(context, settlement, maturity, basis_)
    denominator = 1 - rate * days / year
    if denominator <= 0:
        return NUM
    return checked(invested / denominator)


@function("PRICEDISC", V, V, V, V, V, minimum=4)
def PRICEDISC(
    context: Context, settlement: Scalar, maturity: Scalar, discount: Scalar, redemption: Scalar, basis_: Scalar | None = None
) -> Value:
    rate = context.number(discount)
    value = context.number(redemption)
    if rate <= 0 or value <= 0:
        return NUM
    days, year = _span(context, settlement, maturity, basis_)
    return checked(value - rate * value * days / year)


@function("YIELDDISC", V, V, V, V, V, minimum=4)
def YIELDDISC(
    context: Context, settlement: Scalar, maturity: Scalar, pr: Scalar, redemption: Scalar, basis_: Scalar | None = None
) -> Value:
    price = context.number(pr)
    value = context.number(redemption)
    if price <= 0 or value <= 0:
        return NUM
    days, year = _span(context, settlement, maturity, basis_)
    return checked((value - price) / price * year / days)


def _matured(
    context: Context, settlement: Scalar, maturity: Scalar, issue: Scalar, basis_: Scalar | None
) -> tuple[float, float, float, float]:
    """The days from issue to maturity, issue to settlement and settlement
    to maturity, and the days in the year, as the basis counts them."""
    first, last = _dated(context, settlement, maturity)
    issued = serial_of(context, issue)
    kind = basis(context, basis_)
    if issued >= first:
        raise ExcelError(NUM)
    whole, year = _day_count(context, issued, last, kind)
    accrued, _ = _day_count(context, issued, first, kind)
    remaining, _ = _day_count(context, first, last, kind)
    return whole, accrued, remaining, year


@function("PRICEMAT", V, V, V, V, V, V, minimum=5)
def PRICEMAT(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    rate: Scalar,
    yld: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    whole, accrued, remaining, year = _matured(context, settlement, maturity, issue, basis_)
    coupon_rate = context.number(rate)
    y = context.number(yld)
    if coupon_rate < 0 or y < 0:
        return NUM
    # Excel's documented formula, in its order of operations.
    value = (100 + (whole / year * coupon_rate * 100)) / (1 + (remaining / year * y)) - (accrued / year * coupon_rate * 100)
    return checked(value)


@function("YIELDMAT", V, V, V, V, V, V, minimum=5)
def YIELDMAT(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    rate: Scalar,
    pr: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    whole, accrued, remaining, year = _matured(context, settlement, maturity, issue, basis_)
    coupon_rate = context.number(rate)
    price = context.number(pr)
    if coupon_rate < 0 or price <= 0:
        return NUM
    paid = price / 100 + accrued / year * coupon_rate
    return checked(((1 + whole / year * coupon_rate) - paid) / paid * (year / remaining))


def _bill(context: Context, settlement: Scalar, maturity: Scalar) -> int:
    first, last = _dated(context, settlement, maturity)
    if last - first > 365:
        raise ExcelError(NUM)
    return last - first


@function("TBILLPRICE", V, V, V)
def TBILLPRICE(context: Context, settlement: Scalar, maturity: Scalar, discount: Scalar) -> Value:
    days = _bill(context, settlement, maturity)
    rate = context.number(discount)
    if rate <= 0:
        return NUM
    value = 100 * (1 - rate * days / 360)
    return NUM if value <= 0 else checked(value)


@function("TBILLYIELD", V, V, V)
def TBILLYIELD(context: Context, settlement: Scalar, maturity: Scalar, pr: Scalar) -> Value:
    days = _bill(context, settlement, maturity)
    price = context.number(pr)
    if price <= 0:
        return NUM
    return checked((100 - price) / price * 360 / days)


@function("TBILLEQ", V, V, V)
def TBILLEQ(context: Context, settlement: Scalar, maturity: Scalar, discount: Scalar) -> Value:
    days = _bill(context, settlement, maturity)
    rate = context.number(discount)
    if rate <= 0:
        return NUM
    if days <= 182:
        return checked(365 * rate / (360 - rate * days))
    price = 100 * (1 - rate * days / 360)
    if price <= 0:
        return NUM
    half = days / 365 - 0.5
    root = math.sqrt((days / 365) ** 2 - 2 * half * (1 - 100 / price))
    return checked((-days / 365 + root) / half)


# ----------------------------------------------------------------------
# French depreciation
# ----------------------------------------------------------------------


def _round_half_up(value: float) -> float:
    return float(Decimal(value).quantize(Decimal(1), rounding=ROUND_HALF_UP))


@function("AMORDEGRC", V, V, V, V, V, V, V, minimum=6)
def AMORDEGRC(
    context: Context,
    cost: Scalar,
    purchased: Scalar,
    first_period: Scalar,
    salvage: Scalar,
    period: Scalar,
    rate: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    value = context.number(cost)
    bought = serial_of(context, purchased)
    ends = serial_of(context, first_period)
    rest = context.number(salvage)
    which = math.trunc(context.number(period))
    share = context.number(rate)
    kind = basis(context, basis_)
    if kind == 2 or share <= 0 or value < rest or which < 0:
        return NUM
    life = 1 / share
    coefficient = 1.0 if life < 3 else 1.5 if life < 5 else 2.0 if life <= 6 else 2.5
    share *= coefficient
    depreciation = _round_half_up(year_fraction(bought, ends, kind, epoch_1904=context.epoch_1904) * share * value)
    value -= depreciation
    left = value - rest
    for index in range(which):
        depreciation = _round_half_up(share * value)
        left -= depreciation
        if left < 0:
            return _round_half_up(value * 0.5) if which - index <= 1 else 0.0
        value -= depreciation
    return depreciation


@function("AMORLINC", V, V, V, V, V, V, V, minimum=6)
def AMORLINC(
    context: Context,
    cost: Scalar,
    purchased: Scalar,
    first_period: Scalar,
    salvage: Scalar,
    period: Scalar,
    rate: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    value = context.number(cost)
    bought = serial_of(context, purchased)
    ends = serial_of(context, first_period)
    rest = context.number(salvage)
    which = math.trunc(context.number(period))
    share = context.number(rate)
    kind = basis(context, basis_)
    if kind == 2 or share <= 0 or value < rest or which < 0:
        return NUM
    whole = value * share
    first = year_fraction(bought, ends, kind, epoch_1904=context.epoch_1904) * share * value
    full = math.trunc((value - rest - first) / whole)
    if which == 0:
        result = first
    elif which <= full:
        result = whole
    elif which == full + 1:
        result = value - rest - whole * full - first
    else:
        result = 0.0
    return checked(max(result, 0.0))


__all__ = ["future_value", "payment", "present_value"]
