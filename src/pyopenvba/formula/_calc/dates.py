"""Dates as Excel's formulas count them, and dates and times read from text.

A date is a serial number of days. In the 1900 system serial 1 is
1 January 1900 and serial 60 is 29 February 1900, a day that never was,
kept from Lotus 1-2-3; serial 0 is "0 January 1900", which ``DAY(0)`` is 0
and ``YEAR(0)`` 1900. In the 1904 system serial 0 is 1 January 1904. The
last date either system has is 31 December 9999.

Text becomes a date or a time the way it does when typed into a cell, in
the en-US locale the corpus was measured in: ``1/2/2020``, ``2020-01-15``,
``15-Jan-2020``, ``Jan 15, 2020``, ``12:30 PM``. A date without a year is in
the current one, so reading ``"1/15"`` needs to know what day it is.
"""

from __future__ import annotations

import datetime as dt
import re

#: The last serial either date system reaches: 31 December 9999.
LAST_SERIAL_1900 = 2958465
LAST_SERIAL_1904 = LAST_SERIAL_1900 - 1462
_EPOCH_1900 = dt.date(1899, 12, 31)
_EPOCH_1900_LATE = dt.date(1899, 12, 30)
_EPOCH_1904 = dt.date(1904, 1, 1)
#: Serial 60, Excel's 29 February 1900.
PHANTOM_LEAP_DAY = 60


def last_serial(epoch_1904: bool) -> int:
    return LAST_SERIAL_1904 if epoch_1904 else LAST_SERIAL_1900


def serial(year: int, month: int, day: int, *, epoch_1904: bool = False) -> int:
    """The serial of a day, with months and days past their end carried,
    as ``DATE`` does: month 13 is January of the next year and day 0 the
    last of the month before. ``year`` is the full year."""
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    first = dt.date(year, month, 1)
    if epoch_1904:
        return (first - _EPOCH_1904).days + day - 1
    if first < dt.date(1900, 3, 1):
        return (first - _EPOCH_1900).days + day - 1
    return (first - _EPOCH_1900_LATE).days + day - 1


def calendar(number: int, *, epoch_1904: bool = False) -> tuple[int, int, int]:
    """The year, month and day of a serial, which must be in range:
    ``(1900, 2, 29)`` for 60 and ``(1900, 1, 0)`` for 0 in the 1900
    system."""
    if epoch_1904:
        day = _EPOCH_1904 + dt.timedelta(days=number)
        return day.year, day.month, day.day
    if number == 0:
        return 1900, 1, 0
    if number == PHANTOM_LEAP_DAY:
        return 1900, 2, 29
    if number < PHANTOM_LEAP_DAY:
        day = _EPOCH_1900 + dt.timedelta(days=number)
    else:
        day = _EPOCH_1900_LATE + dt.timedelta(days=number)
    return day.year, day.month, day.day


def weekday(number: int, *, epoch_1904: bool = False) -> int:
    """The day of the week of a serial, 1 for Sunday through 7 for
    Saturday. Serial 1 was a Sunday by Excel's count, which is wrong
    before March 1900 and kept so."""
    if epoch_1904:
        number += 1462
    return (number - 1) % 7 + 1


def days_in_month(year: int, month: int) -> int:
    """Days in a month, February 1900 having 29 as Excel counts."""
    if month == 2:
        leap = year % 4 == 0 and (year % 100 != 0 or year % 400 == 0 or year == 1900)
        return 29 if leap else 28
    return 30 if month in (4, 6, 9, 11) else 31


# ----------------------------------------------------------------------
# Text
# ----------------------------------------------------------------------

_MONTH_NAMES = (
    "january", "february", "march", "april", "may", "june",
    "july", "august", "september", "october", "november", "december",
)  # fmt: skip
_MONTHS = {name: index for index, name in enumerate(_MONTH_NAMES, start=1)}
_MONTHS.update({name[:3]: index for index, name in enumerate(_MONTH_NAMES, start=1)})

_TIME = re.compile(
    r"(?P<hour>\d+)(?::(?P<minute>\d+)(?::(?P<second>\d+(?:\.\d*)?))?)?"
    r"(?:\s+(?P<meridiem>am|pm|a|p))?$",
    re.IGNORECASE,
)
_NUMERIC_DATE = re.compile(r"(\d+)([/-])(\d+)(?:\2(\d+))?$")
_DAY_MONTH = re.compile(r"(\d+)([- ])([a-z]+)(?:\2(\d+))?$", re.IGNORECASE)
_MONTH_DAY = re.compile(r"([a-z]+)([- ])(\d+)(?:,?\s*(\d+))?$", re.IGNORECASE)


def _year(text: str) -> int | None:
    """A year as typed: two digits are 1930 to 2029, four are themselves."""
    value = int(text)
    if len(text) <= 2:
        return value + (2000 if value < 30 else 1900)
    if len(text) == 4 and 1900 <= value <= 9999:
        return value
    return None


def _day_serial(year: int, month: int, day: int, epoch_1904: bool) -> int | None:
    if not 1 <= month <= 12 or not 1 <= day <= days_in_month(year, month):
        return None
    if year < 1900 or year > 9999:
        return None
    number = serial(year, month, day, epoch_1904=epoch_1904)
    return number if number >= 0 else None


def _date_part(text: str, today: dt.date, epoch_1904: bool) -> int | None:
    """The serial of a date typed without a time, or ``None``."""
    numeric = _NUMERIC_DATE.match(text)
    if numeric:
        first, _, second, third = numeric.groups()
        if len(first) == 4 and third is not None:
            # Year first: 2020-01-15 and 2020/01/15.
            year = _year(first)
            return None if year is None else _day_serial(year, int(second), int(third), epoch_1904)
        if third is None:
            return _day_serial(today.year, int(first), int(second), epoch_1904)
        year = _year(third)
        return None if year is None else _day_serial(year, int(first), int(second), epoch_1904)
    named = _DAY_MONTH.match(text)
    if named:
        day, _, name, year_text = named.groups()
        month = _MONTHS.get(name.lower())
        if month is None:
            return None
        year = today.year if year_text is None else _year(year_text)
        return None if year is None else _day_serial(year, month, int(day), epoch_1904)
    named = _MONTH_DAY.match(text)
    if named:
        name, _, number, year_text = named.groups()
        month = _MONTHS.get(name.lower())
        if month is None:
            return None
        if year_text is not None:
            year = _year(year_text)
            return None if year is None else _day_serial(year, month, int(number), epoch_1904)
        # "Jan 20" is a day this year, and "Jan 2020" the first of a
        # month: a number that cannot be a day is a year.
        found = _day_serial(today.year, month, int(number), epoch_1904)
        if found is not None:
            return found
        year = _year(number)
        return None if year is None else _day_serial(year, month, 1, epoch_1904)
    return None


def _time_part(text: str) -> float | None:
    """The fraction of a day a time typed as text is, which may be a day
    or more: ``25:00`` is 1.0416..."""
    found = _TIME.match(text)
    if not found or (found.group("minute") is None and found.group("meridiem") is None):
        return None
    hour = int(found.group("hour"))
    minute = int(found.group("minute") or 0)
    second = float(found.group("second") or 0)
    meridiem = (found.group("meridiem") or "").lower()
    if meridiem:
        if hour > 12:
            return None
        hour = hour % 12 + (12 if meridiem.startswith("p") else 0)
    return (hour * 3600 + minute * 60 + second) / 86400


def parse_date_time(text: str, today: dt.date, *, epoch_1904: bool = False) -> float | None:
    """The serial a date, a time or both typed as text stand for, or
    ``None`` when the text is neither. ``text`` has its outer spaces
    removed already."""
    time = _time_part(text)
    if time is not None:
        return time
    date = _date_part(text, today, epoch_1904)
    if date is not None:
        return float(date)
    # A date, a space, a time.
    for index in range(len(text) - 1, 0, -1):
        if text[index] != " ":
            continue
        head, tail = text[:index].rstrip(" "), text[index + 1 :].lstrip(" ")
        time = _time_part(tail)
        if time is None:
            continue
        date = _date_part(head, today, epoch_1904)
        if date is not None:
            return date + time
    return None


__all__ = [
    "LAST_SERIAL_1900",
    "LAST_SERIAL_1904",
    "PHANTOM_LEAP_DAY",
    "calendar",
    "days_in_month",
    "last_serial",
    "parse_date_time",
    "serial",
    "weekday",
]
