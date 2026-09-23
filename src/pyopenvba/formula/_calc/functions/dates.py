"""Date and time functions.

A date is a serial number of days and a time a fraction of one, see
:mod:`~pyopenvba.formula._calc.dates`. A function reading a date takes
the whole part of the number, and one outside 0 to 31 December 9999 is
``#NUM!``. Text that reads as a date is a date, so ``YEAR("2020-01-15")``
is 2020.
"""

from __future__ import annotations

import calendar
import datetime as dt
import math

from pyopenvba.formula._calc import dates, precise
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import numbers
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import NUM, VALUE, Empty, ExcelError, Scalar, Value


def serial_of(context: Context, value: Scalar) -> int:
    """A date argument's day: the whole part, in range."""
    number = context.number(value)
    if number < 0 or number >= dates.last_serial(context.epoch_1904) + 1:
        raise ExcelError(NUM)
    return math.floor(number)


def _calendar(context: Context, value: Scalar) -> tuple[int, int, int]:
    return dates.calendar(serial_of(context, value), epoch_1904=context.epoch_1904)


@function("DATE", V, V, V)
def DATE(context: Context, year: Scalar, month: Scalar, day: Scalar) -> Value:
    y = int(context.number(year))
    m = int(context.number(month))
    d = int(context.number(day))
    if y < 0 or y >= 10000:
        return NUM
    if y < 1900:
        # DATE(99,1,1) is 1999.
        y += 1900
    try:
        serial = dates.serial(y, m, d, epoch_1904=context.epoch_1904)
    except (ValueError, OverflowError):
        return NUM
    if serial < 0 or serial > dates.last_serial(context.epoch_1904):
        return NUM
    return float(serial)


@function("YEAR", V)
def YEAR(context: Context, value: Scalar) -> Value:
    return float(_calendar(context, value)[0])


@function("MONTH", V)
def MONTH(context: Context, value: Scalar) -> Value:
    return float(_calendar(context, value)[1])


@function("DAY", V)
def DAY(context: Context, value: Scalar) -> Value:
    return float(_calendar(context, value)[2])


#: WEEKDAY's return types: the day numbered 1, and the numbering's start.
_WEEKDAY_TYPES = {1: (1, 1), 2: (2, 1), 3: (2, 0), 11: (2, 1), 12: (3, 1), 13: (4, 1), 14: (5, 1), 15: (6, 1), 16: (7, 1), 17: (1, 1)}


@function("WEEKDAY", V, V, minimum=1)
def WEEKDAY(context: Context, value: Scalar, kind: Scalar | None = None) -> Value:
    serial = serial_of(context, value)
    which = 1 if kind is None or isinstance(kind, Empty) else int(context.number(kind))
    if which not in _WEEKDAY_TYPES:
        return NUM
    sunday_based = dates.weekday(serial, epoch_1904=context.epoch_1904)  # 1 = Sunday
    first, base = _WEEKDAY_TYPES[which]
    return float((sunday_based - first) % 7 + base)


def _real_days(year: int, month: int) -> int:
    """Days in a month of the real calendar, in which 1900 has no 29
    February: EDATE and EOMONTH count by it, so a month after 31 January
    1900 is the 28th, serial 59."""
    return calendar.monthrange(year, month)[1]


def _add_months(year: int, month: int, day: int, months: int) -> tuple[int, int, int]:
    count = year * 12 + (month - 1) + months
    year, month = divmod(count, 12)
    month += 1
    return year, month, min(day, _real_days(year, month))


@function("EDATE", V, V)
def EDATE(context: Context, start: Scalar, months: Scalar) -> Value:
    year, month, day = _calendar(context, start)
    y, m, d = _add_months(year, month, max(day, 1), int(context.number(months)))
    if not 1900 <= y <= 9999:
        return NUM
    return float(dates.serial(y, m, d, epoch_1904=context.epoch_1904))


@function("EOMONTH", V, V)
def EOMONTH(context: Context, start: Scalar, months: Scalar) -> Value:
    year, month, _ = _calendar(context, start)
    y, m, _ = _add_months(year, month, 1, int(context.number(months)))
    if not 1900 <= y <= 9999:
        return NUM
    return float(dates.serial(y, m, _real_days(y, m), epoch_1904=context.epoch_1904))


@function("DATEDIF", V, V, V)
def DATEDIF(context: Context, start: Scalar, end: Scalar, unit: Scalar) -> Value:
    first = serial_of(context, start)
    last = serial_of(context, end)
    kind = context.text(unit).upper()
    if first > last:
        return NUM
    y1, m1, d1 = dates.calendar(first, epoch_1904=context.epoch_1904)
    y2, m2, d2 = dates.calendar(last, epoch_1904=context.epoch_1904)
    months = (y2 - y1) * 12 + (m2 - m1) - (1 if d2 < d1 else 0)
    if kind == "D":
        return float(last - first)
    if kind == "M":
        return float(months)
    if kind == "Y":
        return float(months // 12)
    if kind == "YM":
        return float(months % 12)
    if kind == "MD":
        if d2 >= d1:
            return float(d2 - d1)
        previous_year, previous_month = (y2, m2 - 1) if m2 > 1 else (y2 - 1, 12)
        return float(dates.days_in_month(previous_year, previous_month) - d1 + d2)
    if kind == "YD":
        shifted_year = y2 if (m2, d2) >= (m1, d1) else y2 - 1
        anchor = dates.serial(shifted_year, m1, min(d1, dates.days_in_month(shifted_year, m1)), epoch_1904=context.epoch_1904)
        return float(last - anchor)
    return NUM


def _date_text(context: Context, value: Scalar) -> float:
    """A date or time typed as text, as DATEVALUE and TIMEVALUE read it:
    text only, and not a plain number."""
    if not isinstance(value, str):
        raise ExcelError(VALUE)
    body = value.strip(" ")
    found = dates.parse_date_time(body, context.today, epoch_1904=context.epoch_1904)
    if found is None:
        raise ExcelError(VALUE)
    return found


@function("DATEVALUE", V)
def DATEVALUE(context: Context, text: Scalar) -> Value:
    return float(math.floor(_date_text(context, text)))


@function("TIMEVALUE", V)
def TIMEVALUE(context: Context, text: Scalar) -> Value:
    found = _date_text(context, text)
    return found - math.floor(found)


@function("TIME", V, V, V)
def TIME(context: Context, hour: Scalar, minute: Scalar, second: Scalar) -> Value:
    h = int(context.number(hour))
    m = int(context.number(minute))
    s = int(context.number(second))
    if h > 32767 or m > 32767 or s > 32767:
        return NUM
    total = (h * 3600 + m * 60 + s) / 86400
    if total < 0:
        return NUM
    return total - math.floor(total)


def _clock(context: Context, value: Scalar) -> int:
    """A time argument as whole seconds into its day. Measured, 3493 of
    3493 serials a few units either side of a half second: half a second
    is added to the serial, and the hours, minutes and seconds taken off
    its fraction one after another, each step as the x87 rounds it."""
    number = context.number(value)
    if number < 0:
        raise ExcelError(NUM)
    moved = precise.add(number, 0.5 / 86400)
    part = precise.multiply(precise.subtract(moved, float(math.floor(moved))), 24.0)
    hours = math.floor(part)
    part = precise.multiply(precise.subtract(part, float(hours)), 60.0)
    minutes = math.floor(part)
    part = precise.multiply(precise.subtract(part, float(minutes)), 60.0)
    return (hours * 3600 + minutes * 60 + math.floor(part)) % 86400


@function("HOUR", V)
def HOUR(context: Context, value: Scalar) -> Value:
    return float(_clock(context, value) // 3600)


@function("MINUTE", V)
def MINUTE(context: Context, value: Scalar) -> Value:
    return float(_clock(context, value) // 60 % 60)


@function("SECOND", V)
def SECOND(context: Context, value: Scalar) -> Value:
    return float(_clock(context, value) % 60)


@function("NOW", volatile=True)
def NOW(context: Context) -> Value:
    moment = context.now
    day = float(dates.serial(moment.year, moment.month, moment.day, epoch_1904=context.epoch_1904))
    return day + (moment.hour * 3600 + moment.minute * 60 + moment.second + moment.microsecond / 1e6) / 86400


@function("TODAY", volatile=True)
def TODAY(context: Context) -> Value:
    day = context.today
    return float(dates.serial(day.year, day.month, day.day, epoch_1904=context.epoch_1904))


@function("DAYS", V, V)
def DAYS(context: Context, end: Scalar, start: Scalar) -> Value:
    return float(serial_of(context, end) - serial_of(context, start))


@function("DAYS360", V, V, V, minimum=2)
def DAYS360(context: Context, start: Scalar, end: Scalar, method: Scalar | None = None) -> Value:
    european = method is not None and not isinstance(method, Empty) and context.logical(method)
    y1, m1, d1 = _calendar(context, start)
    y2, m2, d2 = _calendar(context, end)
    if european:
        d1 = min(d1, 30)
        d2 = min(d2, 30)
    else:
        last_feb = m1 == 2 and d1 == dates.days_in_month(y1, 2)
        if d1 == 31 or last_feb:
            d1 = 30
        if d2 == 31 and d1 >= 30:
            d2 = 30
    return float((y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1))


def _holidays(context: Context, value: Value | None) -> set[int]:
    if value is None or isinstance(value, Empty):
        return set()
    return {math.floor(day) for day in numbers(context, (value,))}


def _is_workday(context: Context, serial: int, weekend: frozenset[int], holidays: set[int]) -> bool:
    return dates.weekday(serial, epoch_1904=context.epoch_1904) not in weekend and serial not in holidays


#: Sunday and Saturday, by WEEKDAY's numbering from 1 on Sunday.
_WEEKEND = frozenset({1, 7})

#: The weekend codes of NETWORKDAYS.INTL and WORKDAY.INTL: 1 to 7 a pair
#: of days from Saturday and Sunday on, 11 to 17 a single day from Sunday.
_WEEKENDS = {
    **{code: frozenset({(code + 5) % 7 + 1, (code + 6) % 7 + 1}) for code in range(1, 8)},
    **{code: frozenset({code - 10}) for code in range(11, 18)},
}


def _weekend(context: Context, value: Scalar | None, *, every_day: bool = False) -> frozenset[int]:
    """A weekend by its code, or as seven 0s and 1s from Monday. A week of
    nothing but weekend is ``#VALUE!`` unless ``every_day`` allows it:
    measured, NETWORKDAYS.INTL counts no working days in it."""
    if value is None or isinstance(value, Empty):
        return _WEEKEND
    if isinstance(value, str):
        if len(value) != 7 or set(value) - {"0", "1"} or (value == "1111111" and not every_day):
            raise ExcelError(VALUE)
        # Monday first: position i is WEEKDAY's day (i + 1) % 7 + 1.
        return frozenset((index + 1) % 7 + 1 for index, flag in enumerate(value) if flag == "1")
    code = math.trunc(context.number(value))
    if code not in _WEEKENDS:
        raise ExcelError(NUM)
    return _WEEKENDS[code]


def _working_days(context: Context, start: Scalar, end: Scalar, weekend: frozenset[int], holidays: Value | None) -> Value:
    first = serial_of(context, start)
    last = serial_of(context, end)
    skip = _holidays(context, holidays)
    step = 1 if last >= first else -1
    count = sum(1 for day in range(first, last + step, step) if _is_workday(context, day, weekend, skip))
    return float(count * step)


def _working_day(context: Context, start: Scalar, days: Scalar, weekend: frozenset[int], holidays: Value | None) -> Value:
    current = serial_of(context, start)
    remaining = int(context.number(days))
    skip = _holidays(context, holidays)
    step = 1 if remaining >= 0 else -1
    while remaining:
        current += step
        if current < 0 or current > dates.last_serial(context.epoch_1904):
            return NUM
        if _is_workday(context, current, weekend, skip):
            remaining -= step
    return float(current)


@function("NETWORKDAYS", V, V, R, minimum=2)
def NETWORKDAYS(context: Context, start: Scalar, end: Scalar, holidays: Value | None = None) -> Value:
    return _working_days(context, start, end, _WEEKEND, holidays)


@function("NETWORKDAYS.INTL", V, V, V, R, minimum=2)
def NETWORKDAYS_INTL(
    context: Context, start: Scalar, end: Scalar, weekend: Scalar | None = None, holidays: Value | None = None
) -> Value:
    return _working_days(context, start, end, _weekend(context, weekend, every_day=True), holidays)


@function("WORKDAY", V, V, R, minimum=2)
def WORKDAY(context: Context, start: Scalar, days: Scalar, holidays: Value | None = None) -> Value:
    return _working_day(context, start, days, _WEEKEND, holidays)


@function("WORKDAY.INTL", V, V, V, R, minimum=2)
def WORKDAY_INTL(
    context: Context, start: Scalar, days: Scalar, weekend: Scalar | None = None, holidays: Value | None = None
) -> Value:
    return _working_day(context, start, days, _weekend(context, weekend), holidays)


def _basis_days(year: int) -> int:
    return 366 if dates.days_in_month(year, 2) == 29 else 365


def basis(context: Context, value: Scalar | None) -> int:
    """A day-count basis, 0 to 4, or ``#NUM!``."""
    kind = 0 if value is None or isinstance(value, Empty) else math.trunc(context.number(value))
    if kind not in (0, 1, 2, 3, 4):
        raise ExcelError(NUM)
    return kind


def days_between(first: int, last: int, kind: int, *, epoch_1904: bool = False) -> int:
    """The days from one serial to a later one under a day-count basis:
    30/360 for 0 (the US rules YEARFRAC measures) and 4 (European), the
    actual days for 1, 2 and 3."""
    if kind not in (0, 4):
        return last - first
    y1, m1, d1 = dates.calendar(first, epoch_1904=epoch_1904)
    y2, m2, d2 = dates.calendar(last, epoch_1904=epoch_1904)
    if kind == 4:
        d1, d2 = min(d1, 30), min(d2, 30)
    elif d1 == 31 and d2 == 31:
        d1 = d2 = 30
    elif d1 == 31:
        d1 = 30
    elif d1 == 30 and d2 == 31:
        d2 = 30
    elif m1 == 2 and m2 == 2 and d1 == dates.days_in_month(y1, 2) and d2 == dates.days_in_month(y2, 2):
        d1 = d2 = 30
    elif m1 == 2 and d1 == dates.days_in_month(y1, 2):
        d1 = 30
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)


def year_fraction(first: int, last: int, kind: int, *, epoch_1904: bool = False) -> float:
    """YEARFRAC of two serials, in either order, under a basis."""
    if first > last:
        first, last = last, first
    if kind in (0, 4):
        return days_between(first, last, kind, epoch_1904=epoch_1904) / 360
    if kind == 2:
        return (last - first) / 360
    if kind == 3:
        return (last - first) / 365
    # Actual/actual.
    y1, m1, d1 = dates.calendar(first, epoch_1904=epoch_1904)
    y2, m2, d2 = dates.calendar(last, epoch_1904=epoch_1904)
    if y1 == y2 or (y2 == y1 + 1 and (m1, d1) >= (m2, d2)):
        spans_leap = any(
            dates.days_in_month(year, 2) == 29
            and first <= dates.serial(year, 2, 29, epoch_1904=epoch_1904) <= last
            for year in {y1, y2}
        )
        per_year = 366 if (y1 == y2 and _basis_days(y1) == 366) or spans_leap else 365
        return (last - first) / per_year
    years = range(y1, y2 + 1)
    average = sum(_basis_days(year) for year in years) / len(years)
    return (last - first) / average


@function("YEARFRAC", V, V, V, minimum=2)
def YEARFRAC(context: Context, start: Scalar, end: Scalar, basis_: Scalar | None = None) -> Value:
    first = serial_of(context, start)
    last = serial_of(context, end)
    return checked(year_fraction(first, last, basis(context, basis_), epoch_1904=context.epoch_1904))


@function("WEEKNUM", V, V, minimum=1)
def WEEKNUM(context: Context, value: Scalar, kind: Scalar | None = None) -> Value:
    serial = serial_of(context, value)
    which = 1 if kind is None or isinstance(kind, Empty) else int(context.number(kind))
    if which == 21:
        return ISOWEEKNUM(context, float(serial))
    starts = {1: 1, 17: 1, 2: 2, 11: 2, 12: 3, 13: 4, 14: 5, 15: 6, 16: 7}
    if which not in starts:
        return NUM
    year, _, _ = dates.calendar(serial, epoch_1904=context.epoch_1904)
    january_first = dates.serial(year, 1, 1, epoch_1904=context.epoch_1904)
    offset = (dates.weekday(january_first, epoch_1904=context.epoch_1904) - starts[which]) % 7
    return float((serial - january_first + offset) // 7 + 1)


@function("ISOWEEKNUM", V)
def ISOWEEKNUM(context: Context, value: Scalar) -> Value:
    """The ISO week, counted from Excel's own weekdays: serial 1 is a
    Sunday to Excel, so it is in the last week of 1899, week 52."""
    serial = serial_of(context, value)
    year, _, _ = dates.calendar(serial, epoch_1904=context.epoch_1904)
    iso_weekday = (dates.weekday(serial, epoch_1904=context.epoch_1904) + 5) % 7 + 1
    first = dates.serial(year, 1, 1, epoch_1904=context.epoch_1904)
    week = (serial - first + 1 - iso_weekday + 10) // 7
    if week < 1:
        return float(_iso_weeks(year - 1))
    if week > _iso_weeks(year):
        return 1.0
    return float(week)


def _iso_weeks(year: int) -> int:
    """How many ISO weeks a year has: 53 when it starts on a Thursday, or
    on a Wednesday in a leap year."""
    starts = dt.date(year, 1, 1).isoweekday()
    leap = calendar.isleap(year)
    return 53 if starts == 4 or (leap and starts == 3) else 52


__all__ = ["basis", "days_between", "serial_of", "year_fraction"]
