"""Dates as Excel's formulas count them, and dates and times read from text.

A date is a serial number of days. In the 1900 system serial 1 is
1 January 1900 and serial 60 is 29 February 1900, a day that never was,
kept from Lotus 1-2-3; serial 0 is "0 January 1900", which ``DAY(0)`` is 0
and ``YEAR(0)`` 1900. In the 1904 system serial 0 is 1 January 1904. The
last date either system has is 31 December 9999.

Text becomes a date or a time as VALUE and DATEVALUE read it, measured in
Excel in the en-US locale on some sixty thousand strings: ``1/2/2020``,
``2020-01-15``, ``15-Jan-2020``, ``Jan 15, 2020``, ``1/2020``, ``12:30 PM``,
and a date and a time in either order. A date without a year is in the
current one, so reading ``"1/15"`` needs to know what day it is.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from typing import Literal

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
#: A run of digits or of letters: the parts a date or a time is made of.
_PART = re.compile(r"[0-9]+|[A-Za-z]+")
#: Measured: AM or A, and PM or P, where a time's field would be, stand for
#: these, as a month's name stands for its number.
_MERIDIEM_CODES = {"am": 22, "a": 22, "pm": 21, "p": 21}
#: Measured: a time's field may pass its range only when no field before it
#: did, so ``25:00`` and ``0:99`` are times and ``25:99`` is not.
_LIMITS = {"h": 24, "m": 60, "s": 60}
_SECONDS = {"h": 3600, "m": 60, "s": 1}
#: Measured: what may end the text after a time, spaces or one separator.
_TRAILING = re.compile(r" *(?:[/\-:.]|, +) *$")
#: Measured: after a date and a time, numbers are dropped, apart by what
#: may part a date's numbers.
_STRAYS = re.compile(r"(?:(?: *[/\-] *| +|, +)[0-9]+)+ *$")

_Join = Literal["none", "space", "slash", "dash", "comma"]


def month_named(word: str) -> int | None:
    """The month three letters or more of its English name stand for, in
    any case: ``Sept`` and ``DECEM`` name one, ``Ma`` and ``Juli`` do not."""
    lowered = word.lower()
    if len(lowered) < 3:
        return None
    for number, name in enumerate(_MONTH_NAMES, start=1):
        if name.startswith(lowered):
            return number
    return None


def _join(separator: str) -> _Join | None:
    """How a separator joins two parts of a date: nothing, spaces, a slash
    or a dash with spaces around it, or a comma with a space after it;
    ``None`` for anything else."""
    if not separator:
        return "none"
    body = separator.strip(" ")
    if not body:
        return "space"
    if body == "/":
        return "slash"
    if body == "-":
        return "dash"
    if body == "," and separator[separator.index(",") + 1 :].startswith(" "):
        return "comma"
    return None


def _year_of(digits: str) -> int | None:
    """A year as typed: one or two digits are 1930 to 2029, four are
    themselves from 1900 on, and three are no year."""
    value = int(digits)
    if len(digits) <= 2:
        return value + (2000 if value < 30 else 1900)
    return value if len(digits) == 4 and value >= 1900 else None


def _month_of(digits: str) -> int | None:
    value = int(digits)
    return value if len(digits) <= 2 and 1 <= value <= 12 else None


def _day_of(digits: str, year: int, month: int) -> int | None:
    value = int(digits)
    return value if len(digits) <= 2 and 1 <= value <= days_in_month(year, month) else None


class _Dates:
    """Dates read from text: a date without a year is in this day's year,
    and a serial counts in the date system given."""

    def __init__(self, today: dt.date, epoch_1904: bool) -> None:
        self._today = today
        self._epoch_1904 = epoch_1904

    def read(self, text: str) -> int | None:
        """The serial of the date ``text`` is, or ``None``.

        Measured: two or three parts, numbers or one month's name. Numbers
        alone, apart by slashes or dashes, are a month and a day this year,
        or a month and a year when the second is no day of that month; a
        month, a day and a year; or a year of four digits, a month and a
        day. A name and a number are a day this year, or, the name first, a
        month and a year. A day, a name and a year take no comma; a name, a
        day and a year take one, and more than spaces between the day and
        the year."""
        parts = _PART.findall(text)
        separators = _PART.split(text)
        if not 2 <= len(parts) <= 3 or separators[0] or separators[-1]:
            return None
        joins = [_join(separator) for separator in separators[1:-1]]
        if None in joins:
            return None
        numbers = [part for part in parts if part[0].isdigit()]
        if len(numbers) == len(parts):
            if any(join not in ("slash", "dash") for join in joins):
                return None
            return self._numbers(numbers)
        names = [month_named(part) for part in parts if not part[0].isdigit()]
        month = names[0]
        if len(names) != 1 or month is None:
            return None
        shape = "".join("n" if part[0].isdigit() else "m" for part in parts)
        if shape in ("mn", "nm") and "comma" not in joins:
            day = _day_of(numbers[0], self._today.year, month)
            if day is not None:
                return self._serial(self._today.year, month, day)
            year = _year_of(numbers[0]) if shape == "mn" else None
            return None if year is None else self._serial(year, month, 1)
        if (shape == "nmn" and "comma" not in joins) or (shape == "mnn" and "comma" in joins and joins[1] != "space"):
            year = _year_of(numbers[1])
            day = None if year is None else _day_of(numbers[0], year, month)
            return None if year is None or day is None else self._serial(year, month, day)
        return None

    def _numbers(self, numbers: list[str]) -> int | None:
        if len(numbers) == 2:
            month = _month_of(numbers[0])
            if month is None:
                return None
            day = _day_of(numbers[1], self._today.year, month)
            if day is not None:
                return self._serial(self._today.year, month, day)
            year = _year_of(numbers[1])
            return None if year is None else self._serial(year, month, 1)
        if len(numbers[0]) == 4:
            year, month, day_digits = _year_of(numbers[0]), _month_of(numbers[1]), numbers[2]
        else:
            month, year, day_digits = _month_of(numbers[0]), _year_of(numbers[2]), numbers[1]
        if year is None or month is None:
            return None
        day = _day_of(day_digits, year, month)
        return None if day is None else self._serial(year, month, day)

    def _serial(self, year: int, month: int, day: int) -> int | None:
        # Measured: in the 1904 system a date before 1904 is no date.
        found = serial(year, month, day, epoch_1904=self._epoch_1904)
        return found if found >= 0 else None


@dataclass(frozen=True)
class _Time:
    """A time read from the start of some text."""

    seconds: float
    #: Where in the text the time ends.
    end: int
    #: Whether a month's name stood for a field, which no date may join.
    month_named: bool
    #: Whether the time ends in a colon, which nothing may follow.
    open_colon: bool
    #: Whether the time holds a point, which no second point may follow.
    pointed: bool


def _spaces(text: str, index: int) -> int:
    while index < len(text) and text[index] == " ":
        index += 1
    return index


def _token(text: str, index: int) -> str:
    """The run of digits or of letters at ``index``, or ``""``."""
    found = _PART.match(text, index)
    return found.group() if found else ""


def _colon(text: str, index: int) -> int | None:
    """Where the next part starts past spaces, a colon and spaces at
    ``index``, or ``None`` when there is no colon."""
    after = _spaces(text, index)
    if after < len(text) and text[after] == ":":
        return _spaces(text, after + 1)
    return None


def _meridiem(text: str, index: int) -> tuple[str, int] | None:
    """AM, PM, A or P after spaces at ``index``, and where it ends."""
    after = _spaces(text, index)
    word = _token(text, after)
    if after > index and word.lower() in _MERIDIEM_CODES:
        return word.lower(), after + len(word)
    return None


def _field(part: str) -> int | None:
    """A time's field. Measured: a number has at most four digits, and
    three or four only when it is 100 or more, so ``0100:00`` is a time and
    ``010:00`` is not. A word stands for the negative of its code, read as
    an unsigned 32-bit integer; a month's code is its number, which is how
    ``Jan:5`` comes to be 4294967295 hours and five minutes."""
    if part[0].isdigit():
        return int(part) if len(part) <= 2 or (len(part) <= 4 and int(part) >= 100) else None
    code = _MERIDIEM_CODES.get(part.lower()) or month_named(part)
    return None if code is None else 2**32 - code


def _read_time(text: str, start: int) -> _Time | None:
    """The time at ``start``, read as far as it goes: up to three fields
    apart by colons, perhaps a fraction of a second, perhaps AM or PM."""
    first = _token(text, start)
    if not first or first.lower() in _MERIDIEM_CODES:
        return None
    parts = [first]
    index = start + len(first)
    # Measured: in "6 PM : 27" PM stands where the minutes would be.
    word = _meridiem(text, index) if first[0].isdigit() else None
    if word is not None and _colon(text, word[1]) is not None:
        parts.append(word[0])
        index = word[1]
    open_colon = False
    while len(parts) < 3:
        after = _colon(text, index)
        if after is None:
            break
        part = _token(text, after)
        if not part:
            open_colon, index = True, after
            break
        if part.lower() in _MERIDIEM_CODES:
            return None
        parts.append(part)
        index = after + len(part)
    if len(parts) == 1 and (not first[0].isdigit() or (not open_colon and _meridiem(text, index) is None)):
        # One field is a time only before a colon or AM or PM.
        return None
    names = "hms"[: len(parts)]
    fraction = ""
    pointed = not open_colon and index < len(text) and text[index] == "."
    if pointed:
        if len(parts) == 1:
            return None
        # Measured: spaces may come between the point and its digits, and
        # a point with no digits ends the text.
        after_point = _spaces(text, index + 1)
        fraction = _token(text, after_point)
        if not fraction.isdigit():
            if text[after_point:].strip(" "):
                return None
            fraction, index = "", index + 1
        else:
            index = after_point + len(fraction)
            after = _colon(text, index)
            following = "" if after is None else _token(text, after)
            if len(parts) == 2 and after is not None and following.isdigit():
                # Measured: "1:2.5:3" is 1:02:05.5, the digits after the
                # point read as whole seconds and as their fraction, and the
                # field after them dropped.
                parts.append(fraction)
                names = "hms"
                index = after + len(following)
                if index < len(text) and text[index] in ".:":
                    return None
            elif following:
                return None
            elif len(parts) == 2:
                # Minutes and seconds, as 93:22.15.
                names = "ms"
                if after is not None:
                    open_colon, index = True, after
    values = [_field(part) for part in parts]
    numbers = [float(value) for value in values if value is not None]
    if len(numbers) < len(parts):
        return None
    if fraction:
        # Measured: kept to the millisecond, a half rounded up.
        milliseconds = int(fraction[:3].ljust(3, "0")) + (len(fraction) > 3 and fraction[3] >= "5")
        numbers[-1] += milliseconds / 1000
    over = False
    for name, number in zip(names, numbers, strict=True):
        if number >= _LIMITS[name]:
            if over:
                return None
            over = True
    meridiem = None if open_colon else _meridiem(text, index)
    if meridiem is not None:
        # Measured: AM or PM takes an hour up to 12, no field out of range.
        if over or (names[0] == "h" and numbers[0] > 12):
            return None
        if names[0] == "h":
            numbers[0] %= 12
        index = meridiem[1]
    seconds = 0.0
    for name, number in zip(names, numbers, strict=True):
        seconds += number * _SECONDS[name]
    if meridiem is not None and meridiem[0].startswith("p"):
        seconds += 12 * 3600
    month_named_field = any(not part[0].isdigit() and part.lower() not in _MERIDIEM_CODES for part in parts)
    return _Time(seconds, index, month_named_field, open_colon, pointed)


def _ends(rest: str, time: _Time) -> bool:
    """Whether what follows a time may end the text: spaces, or one
    separator, but nothing after a colon and no second point."""
    if not rest.strip(" "):
        return True
    return not time.open_colon and _TRAILING.match(rest) is not None and not (time.pointed and "." in rest)


def _date_after(rest: str, dates: _Dates) -> int | None:
    """The serial of a date after a time: apart by spaces, a slash, a dash,
    a comma or nothing."""
    end = 0
    while end < len(rest) and rest[end] in " /-,":
        end += 1
    tail = rest[end:].rstrip(" ")
    return dates.read(tail) if tail and _join(rest[:end]) is not None else None


def parse_date_time(text: str, today: dt.date, *, epoch_1904: bool = False) -> float | None:
    """The serial a date, a time or both typed as text stand for, or
    ``None`` when the text is neither.

    Measured: a time is up to three fields apart by colons, hours, minutes
    and seconds, or minutes and seconds with a fraction, as ``93:22.15``.
    A field may pass its range when no field before it did, and seconds
    are kept to the millisecond. AM or PM, or A or P, after a space takes
    an hour up to 12. A date and then a time are apart by spaces, and
    numbers after them are dropped; a time and then a date are apart by
    spaces, a slash, a dash, a comma or nothing. Spaces around the text
    are dropped, and so is one separator after a time, so ``12:30, `` is
    a time and ``12:30,`` is not."""
    body = text.lstrip(" ")
    if not body.rstrip(" "):
        return None
    dates = _Dates(today, epoch_1904)
    time = _read_time(body, 0)
    if time is not None:
        rest = body[time.end :]
        if _ends(rest, time):
            return time.seconds / 86400
        if not time.open_colon and not time.month_named:
            date = _date_after(rest, dates)
            if date is not None:
                return date + time.seconds / 86400
    date = dates.read(body.rstrip(" "))
    if date is not None:
        return float(date)
    for index in range(1, len(body)):
        # A date, spaces, then a time.
        if body[index] != " " or body[index - 1] == " ":
            continue
        time = _read_time(body, _spaces(body, index))
        if time is None or time.month_named:
            continue
        rest = body[time.end :]
        if _ends(rest, time) or (not time.open_colon and _STRAYS.match(rest)):
            date = dates.read(body[:index])
            if date is not None:
                return date + time.seconds / 86400
    return None


__all__ = [
    "LAST_SERIAL_1900",
    "LAST_SERIAL_1904",
    "PHANTOM_LEAP_DAY",
    "calendar",
    "days_in_month",
    "last_serial",
    "month_named",
    "parse_date_time",
    "serial",
    "weekday",
]
