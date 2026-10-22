"""Financial functions: the time value of money, depreciation, internal
rates of return, and bonds, bills and discounted securities.

Measured: the time-value functions raise ``1 + rate`` to a power the way
the ``^`` operator does (see :func:`~.evaluator.power`), so a whole number
of periods is multiplied out square by square, and they are plain double
arithmetic: FV, PV, NPER, NPV and the rest give Excel's bits even where
those are a hundred units in the last place from the exact value. PMT,
IPMT, PPMT, CUMIPMT and CUMPRINC are the exception: they form
``(1 + rate)^-n`` through Excel's own accurate logarithm and exponential
(see :func:`_loan`), and give Excel's bits too.

RATE, IRR, XIRR and YIELD solve by iteration, and Excel stops short of
the root. Each follows Excel's own iteration step by step, so it stops
where Excel does, to the bit.
"""

from __future__ import annotations

import itertools
import math
from decimal import ROUND_HALF_UP, Decimal

from pyopenvba.formula._calc import dates, precise
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


def _discount(periods: float, log_growth: float) -> float:
    """``(1 + r)^-periods`` from ``log_growth = ln(1 + r)``, as ``e^-x``,
    ``x = periods ln(1 + r)``."""
    return precise.exp(-precise.multiply(periods, log_growth))


def _repaid(periods: float, log_growth: float) -> float:
    """``1 - (1 + r)^-periods``, the share of a loan that that many
    payments repay, formed as ``-(e^-x - 1)``."""
    return -precise.exp_less_one(-precise.multiply(periods, log_growth))


def _loan(rate: float, periods: float, present: float, future: float, due: int) -> tuple[float, float, float]:
    """What PMT, IPMT, PPMT and the cumulative functions share, each step
    through Excel's own accurate logarithm and exponential: ``ln(1 + r)``;
    ``(pv + fv)/w``, with ``w`` the share of the loan the ``n`` payments
    repay; and the factor that turns a balance into a period's interest,
    ``r``, or at the start of each period ``1/(1/r + 1)``."""
    log_growth = precise.ln_one_plus(rate)
    repaid = _repaid(periods, log_growth)
    if repaid == 0 or math.isinf(repaid):
        raise ExcelError(NUM)
    share = precise.divide(precise.add(present, future), repaid)
    factor = precise.divide(1.0, precise.add(precise.divide(1.0, rate), 1.0)) if due else rate
    return log_growth, share, factor


def payment(rate: float, periods: float, present: float, future: float, due: int) -> float:
    """PMT as Excel computes it: ``-((pv + fv)/w - fv) r`` with the pieces
    of :func:`_loan`. Measured: 15,553 of 15,555 payments to the bit, the
    structure pinned by twelve present or future values sharing each rate
    and term. The two misses are one loan whose logarithm, inside e^x - 1,
    lies too near a midpoint for anything but the x87's own FYL2X to
    decide."""
    if periods == 0 or rate <= -1:
        raise ExcelError(NUM)
    if rate == 0:
        return checked(-precise.divide(precise.add(present, future), periods))
    _, share, factor = _loan(rate, periods, present, future, due)
    return checked(-precise.multiply(precise.subtract(share, future), factor))


def interest(rate: float, period: float, periods: float, present: float, future: float, due: int) -> float:
    """IPMT as Excel computes it: the balance going into ``period``,
    ``(pv + fv)/w_n w_m - fv`` with ``m = n - per + 1`` payments left,
    times the rate factor, each piece as :func:`payment` forms its own.
    Measured: 2,640 of 2,640 to the bit, fv = 0 in groups of twelve present
    values sharing everything else and fv random."""
    if rate == 0 or (period == 1 and due):
        return 0.0
    log_growth, share, factor = _loan(rate, periods, present, future, due)
    owed = precise.subtract(precise.multiply(share, _repaid(periods - period + 1, log_growth)), future)
    return checked(-precise.multiply(owed, factor))


def principal(rate: float, period: float, periods: float, present: float, future: float, due: int) -> float:
    """PPMT as Excel computes it: ``-((pv + fv)/w_n (1 + r)^-m) f``, the
    loan's share discounted over the ``m`` payments left. Measured: 2,640
    of 2,640 to the bit, on the same loans as :func:`interest`, where the
    payment less the interest gives Excel's bits on fewer than half."""
    if rate == 0 or (period == 1 and due):
        return payment(rate, periods, present, future, due)
    log_growth, share, factor = _loan(rate, periods, present, future, due)
    return checked(-precise.multiply(precise.multiply(share, _discount(periods - period + 1, log_growth)), factor))


def repayments(rate: float, periods: float, present: float, first: int, last: int, due: int) -> tuple[float, float]:
    """CUMPRINC and CUMIPMT from period ``first`` to ``last``, in closed
    form. The principal those payments repay is ``pv/w_n (1 + r)^-(n -
    last + 1) w_count``, times ``1 + r`` when payments fall at the end of
    each period; at the start, a first payment in the span is all principal
    and is added on its own. The interest is the payments less the
    principal. Measured: 2,640 of 2,640 of each to the bit, half of them in
    groups of twelve present values sharing everything else."""
    if first > last:
        return 0.0, 0.0
    log_growth, share, factor = _loan(rate, periods, present, 0.0, due)
    paid = -precise.multiply(share, factor)
    opening = bool(due) and first == 1
    if opening:
        first = 2
    count = float(last - first + 1)
    span = precise.multiply(share, -precise.multiply(_discount(periods - last + 1, log_growth), _repaid(count, log_growth)))
    if opening:
        return checked(precise.add(paid, span)), checked(precise.subtract(precise.multiply(count, paid), span))
    repaid = span if due else precise.multiply(span, precise.add(1.0, rate))
    return checked(repaid), checked(precise.subtract(precise.multiply(count, paid), repaid))


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
    """IPMT's and PPMT's arguments. Measured: the period may run past the
    term, up to the term plus one, so ``IPMT(r,10.5,10,pv)`` has a value;
    a rate of -1 or below is refused, as PMT refuses it."""
    r = context.number(rate)
    period = context.number(per)
    periods = context.number(nper)
    if r <= -1 or period < 1 or periods - period + 1 <= 0:
        raise ExcelError(NUM)
    return r, period, periods, context.number(pv), _optional(context, fv, 0.0), _due(context, type_)


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
    return interest(*_period_arguments(context, rate, per, nper, pv, fv, type_))


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
    return principal(*_period_arguments(context, rate, per, nper, pv, fv, type_))


def _cumulative(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> tuple[float, float, float, int, int, int]:
    """CUMIPMT's and CUMPRINC's arguments. Measured: the checks read the
    periods as given, then a fractional start rounds up and a fractional
    end down, so a span holding no whole period is 0."""
    r = context.number(rate)
    periods = context.number(nper)
    present = context.number(pv)
    first = context.number(start)
    last = context.number(end)
    due = context.number(type_)
    if r <= 0 or periods <= 0 or present <= 0 or first < 1 or last < first or last > periods or due not in (0, 1):
        raise ExcelError(NUM)
    return r, periods, present, math.ceil(first), math.floor(last), int(due)


@function("CUMIPMT", V, V, V, V, V, V)
def CUMIPMT(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> Value:
    return repayments(*_cumulative(context, rate, nper, pv, start, end, type_))[1]


@function("CUMPRINC", V, V, V, V, V, V)
def CUMPRINC(
    context: Context, rate: Scalar, nper: Scalar, pv: Scalar, start: Scalar, end: Scalar, type_: Scalar
) -> Value:
    return repayments(*_cumulative(context, rate, nper, pv, start, end, type_))[0]


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
    return _xnpv_sums(rate, amounts, days)[0]


def _xnpv_sums(rate: float, amounts: list[float], days: list[int]) -> tuple[float, float]:
    """The amounts discounted to the first day at ``rate`` and summed, as
    XNPV sums them, and the sum of their sizes."""
    total = size = 0.0
    for amount, day in zip(amounts, days, strict=True):
        term = amount / _power(1 + rate, (day - days[0]) / 365)
        total += term
        size += abs(term)
    return total, size


@function("XNPV", V, R, R)
def XNPV(context: Context, rate: Scalar, values: Value, when: Value) -> Value:
    r = context.number(rate)
    amounts, days = _series(context, values, when)
    if r <= -1:
        return NUM
    return checked(_xnpv(r, amounts, days))


@function("XIRR", R, R, V, minimum=2)
def XIRR(context: Context, values: Value, when: Value, guess: Scalar | None = None) -> Value:
    amounts, days = _series(context, values, when)
    if len(amounts) == 1:
        # One payment is #N/A, whatever its sign: XIRR(1,1) (tests/fixtures/formula/). pyOpenVBA's own
        # (docs/formula_engine.md).
        return NA
    if not (any(amount > 0 for amount in amounts) and any(amount < 0 for amount in amounts)):
        return NUM
    return checked(_xirr(amounts, days, _optional(context, guess, 0.1)))


def _xirr(amounts: list[float], days: list[int], guess: float) -> float:
    """XIRR's root, found as Excel finds it.

    Measured: not Newton's method but halving. The residual takes the first
    amount's sign once the rate passes the root, and a rate whose residual
    has that sign, or is 0, counts as past it. From the guess g, a root
    above a positive guess is sought in [g, 2g], then [2g, 4g] and on, each
    twice the last; above a negative one only in [g, 0]. Below a positive
    guess it is sought in [0, g], [-g, 0] and [-1, -g] in turn, below a
    negative one in [-1, g]; a guess at or below -1 is #NUM!, and a guess
    of 0 is taken as 0.00001. The bracket is halved from its top,
    ``hi - (hi - lo) / 2``, until it is narrower than 1e-8 of
    ``0.5 + |mid|`` and the residual is under 1e-8 of the sum of the
    discounted amounts' sizes, and the last midpoint is the answer. Held to
    2,183 of 2,183 probes bit for bit, a root exactly at a bracket's end
    among them.
    """
    if guess <= -1:
        raise ExcelError(NUM)
    if guess == 0:
        guess = 1e-5

    def past(rate: float) -> bool:
        return _past_root(_xnpv_sums(rate, amounts, days)[0], amounts)

    if not past(guess):
        low, high = guess, precise.add(guess, guess) if guess > 0 else 0.0
        for _ in range(_XIRR_STEPS):
            if high <= low:
                break
            if past(high):
                return _xirr_halved(amounts, days, low, high)
            low, high = high, precise.multiply(high, 2.0)
        raise ExcelError(NUM)
    if guess < 0:
        return _xirr_halved(amounts, days, -1.0, guess)
    top = guess
    for low in (0.0, -guess):
        if low <= -1:
            raise ExcelError(NUM)
        if not past(low):
            return _xirr_halved(amounts, days, low, top)
        top = low
    return _xirr_halved(amounts, days, -1.0, top)


def _past_root(residual: float, amounts: list[float]) -> bool:
    """Whether a rate with this XNPV lies at or past XIRR's root: the
    residual is 0 or has the first amount's sign."""
    return residual == 0 or (residual > 0) == (amounts[0] > 0)


def _xirr_halved(amounts: list[float], days: list[int], low: float, high: float) -> float:
    """The root in ``[low, high]`` by halving, to XIRR's tolerances; the
    residual is never taken at ``low``, which may be -1."""
    for _ in range(_XIRR_STEPS):
        middle = precise.subtract(high, precise.divide(precise.subtract(high, low), 2.0))
        if middle <= -1:
            break
        at_middle, size = _xnpv_sums(middle, amounts, days)
        if _past_root(at_middle, amounts):
            high = middle
        else:
            low = middle
        if precise.subtract(high, low) < _XIRR_CLOSE * (0.5 + abs(middle)) and abs(at_middle) < _XIRR_CLOSE * size:
            return middle
    raise ExcelError(NUM)


#: XIRR's limit on halving or widening steps, and its tolerance, the
#: documented 0.000001 percent.
_XIRR_STEPS = 100
_XIRR_CLOSE = 1e-8


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


def _to_coupon(context: Context, settlement: int, maturity: int, frequency: int, kind: int) -> float:
    """The days from settlement to the next coupon as PRICE, YIELD and
    DURATION count them: what is left of the period, E - A, on every basis
    but actual/actual, where it is the days on the calendar. COUPDAYSNC
    counts calendar days on actual/360 and actual/365 too; measured, PRICE,
    DURATION and MDURATION do not, 160 of 160 bonds on those bases."""
    if kind == 1:
        return _days_after(context, settlement, maturity, frequency, kind)
    return _period_days(context, settlement, maturity, frequency, kind) - _days_before(context, settlement, maturity, frequency, kind)


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
    """A bond's price per 100 of face value at a yield, each step as the x87
    rounds it, which YIELD's iteration shows where PRICE alone does not."""
    period = _period_days(context, settlement, maturity, frequency, kind)
    remaining = precise.divide(_to_coupon(context, settlement, maturity, frequency, kind), period)
    count = _coupon_count(context, settlement, maturity, frequency)
    accrued = _days_before(context, settlement, maturity, frequency, kind)
    coupon = precise.divide(precise.multiply(100.0, rate), frequency)
    if count == 1:
        discount = precise.add(1.0, precise.divide(precise.multiply(remaining, yld), frequency))
        return precise.subtract(precise.divide(precise.add(redemption, coupon), discount), precise.multiply(coupon, precise.divide(accrued, period)))
    # Measured, 1,078 of 1,078 bonds: the coupons summed first, then the
    # redemption, then the accrued interest taken off, formed from the rate
    # as A/E * rate * 100 / frequency rather than from the coupon.
    growth = precise.add(1.0, precise.divide(yld, frequency))
    price = 0.0
    for index in range(count):
        price = precise.add(price, precise.divide(coupon, _power(growth, precise.add(float(index), remaining))))
    price = precise.add(price, precise.divide(redemption, _power(growth, precise.add(float(count - 1), remaining))))
    interest = precise.multiply(precise.multiply(precise.divide(accrued, period), rate), 100.0)
    return precise.subtract(price, precise.divide(interest, frequency))


def _price_slope(
    context: Context, settlement: int, maturity: int, rate: float, yld: float, redemption: float, frequency: int, kind: int
) -> float:
    """The slope of :func:`_price` in the yield, from its own terms."""
    period = _period_days(context, settlement, maturity, frequency, kind)
    remaining = precise.divide(_to_coupon(context, settlement, maturity, frequency, kind), period)
    count = _coupon_count(context, settlement, maturity, frequency)
    return _slope(count, remaining, rate, yld, redemption, frequency)


def _slope(count: int, remaining: float, rate: float, yld: float, redemption: float, frequency: int) -> float:
    """A bond's price slope in the yield over ``count`` full coupons, the
    first ``remaining`` periods away: each payment's time t, in periods,
    times the payment over the growth to the power t + 1, summed and
    divided by the frequency."""
    coupon = precise.divide(precise.multiply(100.0, rate), frequency)
    growth = precise.add(1.0, precise.divide(yld, frequency))
    total = 0.0
    for index in range(count):
        time = precise.add(float(index), remaining)
        total = precise.add(total, precise.divide(precise.multiply(time, coupon), _power(growth, precise.add(time, 1.0))))
    time = precise.add(float(count - 1), remaining)
    total = precise.add(total, precise.divide(precise.multiply(time, redemption), _power(growth, precise.add(time, 1.0))))
    return -precise.divide(total, frequency)


def _yield_start(context: Context, settlement: int, maturity: int, rate: float, price: float, kind: int) -> float:
    """YIELD's first estimate: the coupons over the term and the discount,
    over the mean of par and price less a quarter of the discount, which on
    a zero-coupon bond is one Halley step from 0. Par stands in for the
    redemption. The term is whole years, then the days from settlement moved
    into the last of them, 29 February to 1 March, over the basis' year, on
    actual/actual the maturity's."""
    epoch = context.epoch_1904
    y1, m1, d1 = dates.calendar(settlement, epoch_1904=epoch)
    y2, m2, d2 = dates.calendar(maturity, epoch_1904=epoch)
    whole = y2 - y1
    later = (m1, d1) > (m2, d2)
    if whole and later:
        whole -= 1
    anchor = settlement
    if whole:
        year = y2 - 1 if later else y2
        month, day = (3, 1) if (m1, d1) == (2, 29) and dates.days_in_month(year, 2) == 28 else (m1, d1)
        anchor = dates.serial(year, month, day, epoch_1904=epoch)
    stub = float(days_between(anchor, maturity, kind, epoch_1904=epoch))
    length = 365.0 if kind == 3 else (366.0 if dates.days_in_month(y2, 2) == 29 else 365.0) if kind == 1 else 360.0
    years = precise.add(float(whole), precise.divide(stub, length))
    discount = precise.subtract(100.0, price)
    # Measured to the bit on 562 starts: (100 r Y + B) / (100 Y - B Y / 2 -
    # B / 4), B the discount, and no other arrangement.
    top = precise.add(precise.multiply(100.0, precise.multiply(rate, years)), discount)
    mean = precise.subtract(precise.multiply(100.0, years), precise.divide(precise.multiply(discount, years), 2.0))
    return precise.divide(top, precise.subtract(mean, precise.divide(discount, 4.0)))


def _yield_one_coupon(
    context: Context, settlement: int, maturity: int, rate: float, price: float, redemption: float, frequency: int, kind: int
) -> float:
    """YIELD with one coupon left, in closed form. Measured: on every
    actual basis the period is the calendar's, and the days to redemption
    are counted by the basis, not as what is left of the period."""
    if kind in (0, 4):
        period = _period_days(context, settlement, maturity, frequency, kind)
    else:
        previous, following = _coupons(context, settlement, maturity, frequency)
        period = float(following - previous)
    accrued = _days_before(context, settlement, maturity, frequency, kind)
    remaining = float(days_between(settlement, maturity, kind, epoch_1904=context.epoch_1904))
    per_coupon = precise.divide(rate, frequency)
    paid = precise.add(precise.divide(price, 100.0), precise.multiply(precise.divide(accrued, period), per_coupon))
    gain = precise.subtract(precise.add(precise.divide(redemption, 100.0), per_coupon), paid)
    return precise.multiply(precise.divide(gain, paid), precise.divide(precise.multiply(float(frequency), period), remaining))


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
    if _coupon_count(context, first, last, times) == 1:
        return checked(_yield_one_coupon(context, first, last, coupon_rate, target, value, times, kind))
    # Measured: Newton's method from _yield_start with PRICE's own slope,
    # stopping when a step would be under 1e-10 and answering with the
    # yield it would have stepped from, so up to 1e-10 short of the root;
    # 1,494 of 1,496 probes to the bit, the other two a unit in the last
    # place away, one of them where ^ itself is.
    y = _yield_start(context, first, last, coupon_rate, target, kind)
    try:
        for _ in range(_YIELD_STEPS):
            gap = precise.subtract(_price(context, first, last, coupon_rate, y, value, times, kind), target)
            step = precise.divide(gap, _price_slope(context, first, last, coupon_rate, y, value, times, kind))
            if abs(step) < _YIELD_CLOSE:
                return checked(y)
            y = precise.subtract(y, step)
    except ZeroDivisionError:
        return NUM
    return NUM


_YIELD_STEPS = 100
_YIELD_CLOSE = 1e-10


# ----------------------------------------------------------------------
# Bonds with an odd first or last period
#
# Measured on 1,500 random bonds of each kind and a few hundred designed
# ones, Microsoft's formulas hold, with day counts and quasi-coupon dates
# of Excel's own. ODDLPRICE and ODDLYIELD give Excel's bits on all of
# them, ODDFPRICE on all but four a unit in the last place away.
# ----------------------------------------------------------------------


def _step(context: Context, day: int, months: int) -> int:
    """The quasi-coupon date ``months`` after ``day`` (before, when
    negative), keeping the day of month the month allows, so that a date
    clipped to the 28th stays on the 28th from then on."""
    try:
        return _months_back(context, day, -months, dates.calendar(day, epoch_1904=context.epoch_1904)[2], False)
    except (ValueError, OverflowError) as error:
        raise ExcelError(NUM) from error


def _is_month_end(context: Context, day: int) -> bool:
    year, month, number = dates.calendar(day, epoch_1904=context.epoch_1904)
    return number == dates.days_in_month(year, month)


def _month_end_days(context: Context, first: int, last: int) -> int:
    """30/360 counting every month's last day, February's included, as the
    30th: the length ODDLPRICE gives a quasi period on basis 0, and the
    days it counts into one."""
    y1, m1, d1 = dates.calendar(first, epoch_1904=context.epoch_1904)
    y2, m2, d2 = dates.calendar(last, epoch_1904=context.epoch_1904)
    d1 = 30 if d1 == dates.days_in_month(y1, m1) else d1
    d2 = 30 if d2 == dates.days_in_month(y2, m2) else d2
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def _odd_last_dates(context: Context, last: int, maturity: int, frequency: int) -> list[int]:
    """The quasi-coupon dates from the last interest date on, each a period
    after the one before, to the first that covers the maturity. From a
    month's last day, a date covers any maturity in its month, clipped to
    the 28th or not."""
    step = 12 // frequency
    month_end = _is_month_end(context, last)
    epoch = context.epoch_1904
    ends = dates.calendar(maturity, epoch_1904=epoch)[:2]
    found = [last]
    while True:
        following = _step(context, found[-1], step)
        found.append(following)
        if month_end and ends <= dates.calendar(following, epoch_1904=epoch)[:2]:
            return found
        if not month_end and maturity <= following:
            return found


def _odd_last_terms(
    context: Context, settlement: int, maturity: int, last: int, frequency: int, kind: int
) -> tuple[float, float, float]:
    """ODDLPRICE's sums over the quasi periods, each count over the
    period's normal length NL: DC, the days counted in each; A, those
    accrued; DSC, those from settlement to the maturity. On basis 0, NL and
    DC count month ends as the 30th while A and DSC count as YEARFRAC
    does; a quasi period wholly accrued counts as one; the last runs to
    the maturity for DC, even past its own end."""
    epoch = context.epoch_1904

    def part(first: int, final: int) -> float:
        return float(days_between(first, final, kind, epoch_1904=epoch))

    def whole(first: int, final: int) -> float:
        return float(_month_end_days(context, first, final)) if kind == 0 else part(first, final)

    quasi = _odd_last_dates(context, last, maturity, frequency)
    counted = accrued = remaining = 0.0
    for index in range(1, len(quasi)):
        start, end = quasi[index - 1], quasi[index]
        final = index == len(quasi) - 1
        length = whole(start, end)
        counted = precise.add(counted, precise.divide(whole(start, maturity if final else end), length))
        if settlement >= end and not final:
            accrued = precise.add(accrued, 1.0)
        elif settlement > start:
            accrued = precise.add(accrued, precise.divide(part(start, settlement), length))
        stop = min(end, maturity)
        if settlement < stop:
            remaining = precise.add(remaining, precise.divide(part(max(start, settlement), stop), length))
    return counted, accrued, remaining


def _odd_last(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None,
) -> tuple[float, float, float, float, float, float]:
    """ODDLPRICE's and ODDLYIELD's arguments, checked: the redemption, the
    coupon per period, DC, A and DSC, and the frequency."""
    settled = serial_of(context, settlement)
    due = serial_of(context, maturity)
    previous = serial_of(context, last_interest)
    coupon_rate = context.number(rate)
    value = context.number(redemption)
    times = _frequency(context, frequency)
    kind = basis(context, basis_)
    if not due > settled > previous or coupon_rate < 0 or value <= 0:
        raise ExcelError(NUM)
    counted, accrued, remaining = _odd_last_terms(context, settled, due, previous, times, kind)
    coupon = precise.divide(precise.multiply(100.0, coupon_rate), times)
    return value, coupon, counted, accrued, remaining, float(times)


@function("ODDLPRICE", V, V, V, V, V, V, V, V, minimum=7)
def ODDLPRICE(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    value, coupon, counted, accrued, remaining, times = _odd_last(
        context, settlement, maturity, last_interest, rate, redemption, frequency, basis_
    )
    y = context.number(yld)
    if y < 0:
        return NUM
    discount = precise.add(1.0, precise.divide(precise.multiply(remaining, y), times))
    paid = precise.add(value, precise.multiply(counted, coupon))
    return checked(precise.subtract(precise.divide(paid, discount), precise.multiply(accrued, coupon)))


@function("ODDLYIELD", V, V, V, V, V, V, V, V, minimum=7)
def ODDLYIELD(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    last_interest: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    value, coupon, counted, accrued, remaining, times = _odd_last(
        context, settlement, maturity, last_interest, rate, redemption, frequency, basis_
    )
    price = context.number(pr)
    if price <= 0:
        return NUM
    # Measured: no days left to maturity, as 30/360 counts from the 30th to
    # the 31st, gives 0.
    if remaining == 0:
        return 0.0
    paid = precise.add(value, precise.multiply(counted, coupon))
    cost = precise.add(price, precise.multiply(accrued, coupon))
    return checked(precise.multiply(precise.divide(precise.subtract(paid, cost), cost), precise.divide(times, remaining)))


def _extra_period(context: Context, settlement: int, first: int, frequency: int) -> int:
    """The whole period ODDFPRICE adds to a long odd first period's
    discounting, measured: when the first coupon falls on a month's last
    day and settlement does not, unless settlement shares its month with a
    quasi-coupon date other than the first coupon."""
    if not _is_month_end(context, first) or _is_month_end(context, settlement):
        return 0
    epoch = context.epoch_1904
    y1, m1, _ = dates.calendar(settlement, epoch_1904=epoch)
    y2, m2, _ = dates.calendar(first, epoch_1904=epoch)
    months = (y2 - y1) * 12 + m2 - m1
    return 1 if months == 0 or months % (12 // frequency) else 0


class _OddFirst:
    """A bond with an odd first period, as ODDFPRICE prices it.

    Measured: short or long is decided by the days from issue to the first
    coupon against the coupon period E, the first coupon's own; an odd
    period exactly E long is refused. A short one is Microsoft's formula
    with DFC, A and DSC counted by the basis. A long one steps quasi-coupon
    dates back from the first coupon past the issue, each from the one
    before; each whole quasi period counts as one coupon's worth, the
    partial first by its days, each over E on every basis but
    actual/actual, where over its own length; the accrued periods count by
    their days, and both are summed from the first coupon back.

    ``time`` is the periods to the first coupon that discount every cash
    flow; ``remaining`` and ``count``, the same without the extra period
    and the coupons from settlement, shape ODDFYIELD's slope.
    """

    def __init__(
        self, context: Context, settled: int, due: int, issued: int, first: int, frequency: int, kind: int
    ) -> None:
        epoch = context.epoch_1904

        def days(start: int, end: int) -> float:
            return float(days_between(start, end, kind, epoch_1904=epoch))

        period = _period_days(context, settled, first, frequency, kind)
        odd = days(issued, first)
        if odd == period:
            raise ExcelError(NUM)
        self.frequency = frequency
        self.after = _coupon_count(context, first, due, frequency)
        self.long = odd > period
        if not self.long:
            self.counted, self.length = odd, period
            self.accrued = precise.divide(days(issued, settled), period)
            self.time = self.remaining = precise.divide(days(settled, first), period)
            self.count = self.after + 1
            return
        quasi = [first]
        while quasi[-1] > issued:
            quasi.append(_step(context, quasi[-1], -(12 // frequency)))
        shares: list[float] = []
        accrued: list[float] = []
        for position, (start, end) in enumerate(reversed(list(zip(quasi[1:], quasi[:-1], strict=True)))):
            length = float(end - start) if kind == 1 else period
            start = max(start, issued)
            shares.append(precise.divide(days(start, end), length) if position == 0 else 1.0)
            if settled >= end:
                accrued.append(precise.divide(days(start, end), length))
            elif settled > start:
                accrued.append(precise.divide(days(start, settled), length))
        self.counted, self.length = precise.summed(reversed(shares)), 1.0
        self.accrued = precise.summed(reversed(accrued))
        whole = _coupon_count(context, settled, first, frequency) - 1
        self.remaining = precise.add(float(whole), precise.divide(_days_after(context, settled, first, frequency, kind), period))
        self.time = precise.add(self.remaining, float(_extra_period(context, settled, first, frequency)))
        self.count = self.after + whole + 1

    def price(self, rate: float, yld: float, redemption: float) -> float:
        """The redemption, then the odd coupon, then the regular coupons,
        less the accrued interest; the time to the first coupon discounts
        each of them."""
        coupon = precise.divide(precise.multiply(100.0, rate), self.frequency)
        growth = precise.add(1.0, precise.divide(yld, self.frequency))
        odd = precise.divide(precise.divide(precise.multiply(coupon, self.counted), self.length), _power(growth, self.time))
        if self.long:
            owed = precise.multiply(coupon, self.accrued)
        else:
            owed = precise.divide(precise.multiply(precise.multiply(self.accrued, rate), 100.0), self.frequency)
        coupons = 0.0
        for index in range(1, self.after + 1):
            coupons = precise.add(coupons, precise.divide(coupon, _power(growth, precise.add(float(index), self.time))))
        value = precise.divide(redemption, _power(growth, precise.add(float(self.after), self.time)))
        return precise.subtract(precise.add(precise.add(value, odd), coupons), owed)


def _odd_first(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None,
) -> tuple[_OddFirst, int, int, float, float, int]:
    """ODDFPRICE's and ODDFYIELD's arguments, checked, and the bond.
    Measured: the first coupon has to be one of the maturity's own coupon
    dates."""
    settled = serial_of(context, settlement)
    due = serial_of(context, maturity)
    issued = serial_of(context, issue)
    first = serial_of(context, first_coupon)
    coupon_rate = context.number(rate)
    value = context.number(redemption)
    times = _frequency(context, frequency)
    kind = basis(context, basis_)
    if not due > first > settled > issued or coupon_rate < 0 or value <= 0:
        raise ExcelError(NUM)
    if _coupons(context, first, due, times)[0] != first:
        raise ExcelError(NUM)
    return _OddFirst(context, settled, due, issued, first, times, kind), settled, due, coupon_rate, value, kind


@function("ODDFPRICE", V, V, V, V, V, V, V, V, V, minimum=8)
def ODDFPRICE(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    yld: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    bond, _, _, coupon_rate, value, _ = _odd_first(
        context, settlement, maturity, issue, first_coupon, rate, redemption, frequency, basis_
    )
    y = context.number(yld)
    if y < 0:
        return NUM
    return checked(bond.price(coupon_rate, y, value))


@function("ODDFYIELD", V, V, V, V, V, V, V, V, V, minimum=8)
def ODDFYIELD(
    context: Context,
    settlement: Scalar,
    maturity: Scalar,
    issue: Scalar,
    first_coupon: Scalar,
    rate: Scalar,
    pr: Scalar,
    redemption: Scalar,
    frequency: Scalar,
    basis_: Scalar | None = None,
) -> Value:
    bond, settled, due, coupon_rate, value, kind = _odd_first(
        context, settlement, maturity, issue, first_coupon, rate, redemption, frequency, basis_
    )
    target = context.number(pr)
    if target <= 0:
        return NUM
    # Measured: YIELD's iteration from YIELD's start, with the slope of a
    # regular bond whose coupons from settlement all come in full, the first
    # at the odd coupon's time. It gives Excel's bits on 678 of 681 bonds
    # with a short odd first period; with a long one, on 465 of 678, and
    # within 1e-10 of Excel's answer on the rest, Excel's slope there being
    # another. That slope can converge slowly where Excel's does not, so the
    # iteration runs longer than YIELD's before giving up.
    y = _yield_start(context, settled, due, coupon_rate, target, kind)
    try:
        for _ in range(_ODD_YIELD_STEPS):
            gap = precise.subtract(bond.price(coupon_rate, y, value), target)
            step = precise.divide(gap, _slope(bond.count, bond.remaining, coupon_rate, y, value, bond.frequency))
            if abs(step) < _YIELD_CLOSE:
                return checked(y)
            y = precise.subtract(y, step)
    except ZeroDivisionError:
        return NUM
    return NUM


_ODD_YIELD_STEPS = 1000


def _duration(
    context: Context, settlement: int, maturity: int, coupon: float, yld: float, frequency: int, kind: int
) -> float:
    """Macaulay duration, in years: each payment's time, counted in coupon
    periods from settlement, weighted by its present value."""
    period = _period_days(context, settlement, maturity, frequency, kind)
    remaining = precise.divide(_to_coupon(context, settlement, maturity, frequency, kind), period)
    count = _coupon_count(context, settlement, maturity, frequency)
    paid = precise.divide(precise.multiply(coupon, 100.0), frequency)
    growth = precise.add(precise.divide(yld, frequency), 1.0)
    # Measured, 1,800 of 1,800 bonds to the bit: each coupon's present value
    # summed, and its time, index + DSC/E as PRICE counts it, times that
    # present value; then the redemption on its own, at DSC/E + N - 1,
    # which rounds another way, its time times 100 over the power.
    weighted = value = 0.0
    for index in range(count):
        time = precise.add(float(index), remaining)
        present = precise.divide(paid, _power(growth, time))
        value = precise.add(value, present)
        weighted = precise.add(weighted, precise.multiply(time, present))
    time = precise.subtract(precise.add(remaining, float(count)), 1.0)
    discount = _power(growth, time)
    weighted = precise.add(weighted, precise.divide(precise.multiply(time, 100.0), discount))
    value = precise.add(value, precise.divide(100.0, discount))
    return precise.divide(precise.divide(weighted, value), frequency)


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
    duration = _duration(context, first, last, rate, y, times, basis(context, basis_))
    return checked(precise.divide(duration, precise.add(1.0, precise.divide(y, times))))


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
