"""What a value written to a cell becomes, and the number format that comes with it.

Writing a string through ``Range.Value`` is typing it: Excel parses it as
it parses what someone types into the cell, on this model's US English
system. Every rule here was measured (scripts/measure_value_typing.py and
scripts/measure_typing_formats.py):

- "" leaves the cell empty; a string of spaces, tabs or no-break spaces
  stays text, and so does text with spaces around it, unless it reads as
  a number once the spaces go.
- A leading apostrophe keeps the rest as text and marks the cell with a
  prefix character; ``=`` starts a formula, but not after a space.
- TRUE and FALSE, and the seven error names, in any case and nothing
  around them, are the Boolean and the error.
- A number may carry a sign, parentheses for a negative, a dollar sign,
  thousands separators (three or more digits after each), a percent sign,
  an exponent, or a whole part and a fraction; each brings its own format.
- Dates read month first, with two-digit years before 30 in this century
  and years from 1900 to 9999 on Excel's calendar, which has a 29 February
  1900; a month and a number that cannot be its day is a month and year;
  month names, times with or without seconds, AM/PM and fractions of a
  second, and a date with a time each bring their own formats. A date
  that does not exist stays text. A run of spaces between the parts of a
  date or a time counts as one space, though not in a fraction.

What the cell already holds decides the rest. A Text cell (``@``) keeps
every string as it is written and turns a Date or a Currency a macro
computes into text. A cell with a number format of its own reads ``1/2``
as a half rather than the second of January. A format typing brings
replaces the cell's own only when that one is General, or one of Excel's
built-in currency, accounting, percent, scientific, fraction, date and
time formats and the new one is of another kind; thousands separators
never replace a format.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass
from decimal import Decimal

from pyopenvba.formula._display import format_value
from pyopenvba.formula._values import ERRORS, ExcelError
from pyopenvba.interpreter._values import EMPTY, VBACurrency, VBADate, VBAErrorValue, error

#: The errors typing turns into error values; #SPILL! and the rest stay text.
_TYPED_ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")
#: CVErr's numbers for the errors a cell can hold.
_ERROR_CODES = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029, "#NUM!": 2036,
                "#N/A": 2042}
_MONTHS = {name: index for index, names in enumerate(
    (("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"), ("may",), ("jun", "june"),
     ("jul", "july"), ("aug", "august"), ("sep", "september"), ("oct", "october"), ("nov", "november"),
     ("dec", "december")), start=1) for name in names}

_PLAIN = re.compile(r"(?P<whole>\d+(?:,\d{3,})*)?(?P<point>\.(?P<fraction>\d*))?")
_EXPONENT = re.compile(r"(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+")
_MIXED = re.compile(r"(\d+) (\d+)/(\d+)")
_FRACTION = re.compile(r"(\d+)/(\d+)")
_NUMERIC_DATE = re.compile(r"(\d{1,4})([/-])(\d{1,4})(?:\2(\d{1,4}))?")
_DAY_MONTH = re.compile(r"(\d{1,2})[ -]([A-Za-z]+)(?:[ -](\d{2,4}))?")
_MONTH_DAY = re.compile(r"([A-Za-z]+) (\d{1,2})(?:, (\d{2,4}))?")
_MONTH_YEAR = re.compile(r"([A-Za-z]+)[ -](\d{4})")
#: h:m, h:m:s, either with a fraction of a second -- m:s.f when there are two parts -- and AM or PM.
_TIME = re.compile(r"(\d{1,4}):(\d{1,4})(?::(\d{1,4}))?(\.\d+)?(?: ([AaPp][Mm]?))?")
_HOUR = re.compile(r"(\d{1,4}) ([AaPp][Mm]?)")

#: The formats typing gives a dollar amount, and the one a Currency value brings.
TYPED_CURRENCY = ("$#,##0_);[Red]($#,##0)", "$#,##0.00_);[Red]($#,##0.00)")
CURRENCY_VALUE = "$#,##0.00_);($#,##0.00)"

#: Excel's built-in formats that typing may replace, by kind; every other format stays.
_REPLACEABLE = {
    **dict.fromkeys(("$#,##0_);($#,##0)", "$#,##0_);[Red]($#,##0)", "$#,##0.00_);($#,##0.00)",
                     "$#,##0.00_);[Red]($#,##0.00)"), "currency"),
    **dict.fromkeys(('_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
                     '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
                     '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'), "accounting"),
    **dict.fromkeys(("0%", "0.00%"), "percent"),
    **dict.fromkeys(("0.00E+00", "##0.0E+0"), "scientific"),
    **dict.fromkeys(("# ?/?", "# ??/??"), "fraction"),
    **dict.fromkeys(("m/d/yyyy", "d-mmm-yy", "d-mmm", "mmm-yy", "m/d/yyyy h:mm"), "date"),
    **dict.fromkeys(("h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "mm:ss", "[h]:mm:ss", "mm:ss.0"), "time"),
}
#: The kind of each format typing brings; a number with thousands separators replaces nothing.
_BROUGHT = {**_REPLACEABLE, **dict.fromkeys((*TYPED_CURRENCY, CURRENCY_VALUE), "currency"),
            "#,##0": "number", "#,##0.00": "number", "General": "general"}


@dataclass(frozen=True, slots=True)
class Typed:
    """What a cell holds after a write: its value, the format it ends with, and a prefix character."""

    value: object
    #: The cell's format after the write, or None where it keeps its own.
    number_format: str | None = None
    prefix: bool = False


def typed(value: object, current: str = "General", *, raw: bool = False) -> Typed:
    """A value written to a cell whose format is ``current``, as the cell keeps it.

    ``raw`` is Value2's write, which puts a Date or a Currency in as the
    Double under it and so brings no format with it.
    """
    if raw and isinstance(value, VBADate):
        value = value.serial
    elif raw and isinstance(value, VBACurrency):
        value = float(value)
    if current == "@":
        return _as_text(value)
    found = _typed(value, fractions=reads_fractions(current))
    return Typed(found.value, format_after(current, found.number_format), found.prefix)


def format_after(current: str, brought: str | None) -> str | None:
    """The format a cell takes when typing brings ``brought``, or None where it keeps ``current``."""
    if brought is None or brought == current:
        return None
    if current in ("General", ""):
        return None if brought == "General" else brought
    held = _REPLACEABLE.get(current)
    kind = _BROUGHT.get(brought)
    if held is None or kind in (None, "number", held):
        return None
    return brought


def reads_fractions(current: str) -> bool:
    """Whether a cell reads a typed ``1/2`` as a half: it has a number format that is not a date or time."""
    return current not in ("General", "", "@") and not _codes(current)


def _typed(value: object, *, fractions: bool) -> Typed:
    if isinstance(value, str):
        return typed_text(value, fractions=fractions)
    if isinstance(value, VBADate):
        return _typed_date(value.serial)
    if isinstance(value, VBACurrency):
        return Typed(float(value), CURRENCY_VALUE)
    if isinstance(value, bool):
        return Typed(value)
    if isinstance(value, (int, float, Decimal)):
        return Typed(float(value))
    if isinstance(value, VBAErrorValue):
        name = next((name for name, number in _ERROR_CODES.items() if number == value.number), "#VALUE!")
        return Typed(ERRORS.get(name, ExcelError(name)))
    return Typed(value)


def _typed_date(serial: float) -> Typed:
    """A Date: a time alone is a number shown as one, a whole day m/d/yyyy, a day and a time m/d/yyyy h:mm."""
    if serial < 0:
        # Excel's calendar starts in 1900; a Date before it is refused.
        raise error(1004, "Application-defined or object-defined error")
    if serial < 1:
        return Typed(float(serial), "h:mm:ss AM/PM")
    if serial == math.floor(serial):
        return Typed(float(serial), "m/d/yyyy")
    return Typed(float(serial), "m/d/yyyy h:mm")


def _as_text(value: object) -> Typed:
    """A value written to a Text cell: strings as they are, a Date or a Currency as the text Excel shows for it."""
    if isinstance(value, str):
        return Typed(value[1:], prefix=True) if value.startswith("'") else Typed(value)
    if isinstance(value, VBADate):
        return Typed(_date_text(value.serial))
    if isinstance(value, VBACurrency):
        # The cell takes the Currency's format as well, which no other write to a Text cell does.
        return Typed(format_value(float(value), CURRENCY_VALUE), CURRENCY_VALUE)
    return _typed(value, fractions=False)


def _date_text(serial: float) -> str:
    """A Date as a Text cell keeps it: m/d/yyyy on Excel's calendar, h:mm:ss AM/PM, or both, two spaces apart."""
    if serial < 0:
        raise error(1004, "Application-defined or object-defined error")
    day = math.floor(serial)
    seconds = round((serial - day) * 86400)
    if seconds >= 86400:
        day, seconds = day + 1, 0
    code = "h:mm:ss AM/PM" if not day else "m/d/yyyy  h:mm:ss AM/PM" if seconds else "m/d/yyyy"
    return format_value(serial, code)


def _serial(year: int, month: int, day: int) -> float | None:
    """A date's serial on Excel's calendar, or None where it has no such date."""
    if not 1900 <= year <= 9999:
        return None
    if (year, month, day) == (1900, 2, 29):
        return 60.0
    try:
        when = _dt.date(year, month, day)
    except ValueError:
        return None
    days = (when - _dt.date(1899, 12, 30)).days
    return float(days - 1 if when < _dt.date(1900, 3, 1) else days)


def _this_year() -> int:
    """The year a date typed without one falls in."""
    return _dt.date.today().year


def typed_text(text: str, *, fractions: bool = False) -> Typed:
    """A string written to a cell, read as Excel reads typing; ``fractions`` reads ``1/2`` as a half."""
    if text == "":
        return Typed(EMPTY)
    if text.startswith("'"):
        return Typed(text[1:], prefix=True)
    upper = text.upper()
    if upper in _TYPED_ERRORS:
        return Typed(ERRORS.get(upper, ExcelError(upper)))
    if upper in ("TRUE", "FALSE"):
        return Typed(upper == "TRUE")
    body = text.strip(" ")
    if not body:
        return Typed(text)
    found = _number(body, fractions=fractions) or _moment(body)
    if found is None:
        return Typed(text)
    return found if found.value is not None else Typed(text, found.number_format)


# --- numbers ------------------------------------------------------------------------------------


def _fraction_format(denominator: str) -> str:
    """The fraction format a typed fraction brings: as many places as its denominator has, up to two."""
    width = min(len(denominator), 2)
    return f"# {'?' * width}/{'?' * width}"


def _number(body: str, *, fractions: bool) -> Typed | None:
    """A typed number, or None. A value of None is text that still brings a format."""
    text, negative, percent, currency = body, False, False, False
    if text.startswith("(") and text.endswith(")"):
        text, negative = text[1:-1].strip(" "), True
    if text.endswith("%"):
        text, percent = text[:-1].rstrip(" "), True
    sign = ""
    for _ in range(2):
        if text[:1] in ("+", "-") and not sign and not negative:
            sign, text = text[0], text[1:].lstrip(" ")
        elif text[:1] == "$" and not currency:
            currency, text = True, text[1:].lstrip(" ")
    if not text or (currency and percent):
        return None
    negative = negative or sign == "-"
    if _EXPONENT.fullmatch(text):
        if percent:
            return None
        value = float(text)
        if math.isinf(value):
            # Excel keeps a number too large as text, and still gives the cell the scientific format.
            return Typed(None, "0.00E+00")
        return Typed(-value if negative else value, "0.00E+00")
    mixed = _MIXED.fullmatch(text)
    if mixed is not None:
        whole, numerator, denominator = (int(part) for part in mixed.groups())
        if not denominator or currency or percent:
            return None
        value = whole + numerator / denominator
        return Typed(-value if negative else value, _fraction_format(mixed.group(3)))
    alone = _FRACTION.fullmatch(text) if fractions else None
    if alone is not None:
        numerator, denominator = int(alone.group(1)), int(alone.group(2))
        if not numerator or not denominator or currency or percent:
            return None
        value = numerator / denominator
        return Typed(-value if negative else value, _fraction_format(alone.group(2)))
    plain = _PLAIN.fullmatch(text)
    if plain is None or not (plain.group("whole") or plain.group("fraction")):
        return None
    value = float(text.replace(",", ""))
    if percent:
        value /= 100
    value = -value if negative else value
    decimals = plain.group("point") is not None
    if currency:
        return Typed(value, TYPED_CURRENCY[decimals])
    if percent:
        return Typed(value, "0.00%" if decimals else "0%")
    if "," in text:
        return Typed(value, "#,##0.00" if decimals else "#,##0")
    return Typed(value)


# --- dates and times ----------------------------------------------------------------------------


def _moment(body: str) -> Typed | None:
    """A typed date, time, or date and time, or None; a run of spaces separates as one space does."""
    body = re.sub(" {2,}", " ", body)
    time = _time(body)
    if time is not None:
        return time
    date = _date(body)
    if date is not None:
        return date
    head, _, tail = body.partition(" ")
    if tail:
        day, clock = _numeric_date(head), _time(tail)
        if day is not None and clock is not None and isinstance(clock.value, float):
            return Typed(day[0] + clock.value, "m/d/yyyy h:mm")
    return None


def _year(text: str) -> int:
    """A typed year: two digits before 30 are this century's, from 30 the last one's."""
    year = int(text)
    if len(text) <= 2:
        return year + (2000 if year < 30 else 1900)
    return year


def _numeric_date(body: str) -> tuple[float, str] | None:
    """A date typed in digits: m/d/y, y/m/d, m/d this year, or a month and year."""
    found = _NUMERIC_DATE.fullmatch(body)
    if found is None:
        return None
    first, _, second, third = found.groups()
    if third is not None:
        if len(first) == 4:
            serial = _serial(int(first), int(second), int(third))
        else:
            serial = _serial(_year(third), int(first), int(second))
        return None if serial is None else (serial, "m/d/yyyy")
    month, number = int(first), int(second)
    if len(first) > 2 or not 1 <= month <= 12:
        return None
    serial = _serial(_this_year(), month, number) if len(second) <= 2 else None
    if serial is not None:
        return serial, "d-mmm"
    # A number that cannot be the month's day is its year.
    serial = _serial(_year(second), month, 1)
    return None if serial is None else (serial, "mmm-yy")


def _date(body: str) -> Typed | None:
    numeric = _numeric_date(body)
    if numeric is not None:
        return Typed(numeric[0], numeric[1])
    for pattern, order in ((_DAY_MONTH, "dmy"), (_MONTH_DAY, "mdy")):
        found = pattern.fullmatch(body)
        if found is None:
            continue
        day_text, month_text, year_text = (found.group(1), found.group(2), found.group(3)) if order == "dmy" \
            else (found.group(2), found.group(1), found.group(3))
        month = _MONTHS.get(month_text.lower())
        if month is None:
            continue
        year = _year(year_text) if year_text else _this_year()
        serial = _serial(year, month, int(day_text))
        if serial is not None:
            return Typed(serial, "d-mmm-yy" if year_text else "d-mmm")
    found = _MONTH_YEAR.fullmatch(body)
    if found is not None:
        month = _MONTHS.get(found.group(1).lower())
        serial = _serial(int(found.group(2)), month, 1) if month is not None else None
        if serial is not None:
            return Typed(serial, "mmm-yy")
    return None


def _time(body: str) -> Typed | None:
    """A typed time: h:mm, h:mm:ss, m:ss.0 or h:mm:ss.0, any of them with AM or PM, or an hour with AM or PM.

    Each part runs to 9999; a minute or second past 59 makes the number
    General, and an hour past 23 an elapsed time. With AM or PM the hour
    runs from 0 to 12 and the minutes and seconds stop at 59.
    """
    found = _TIME.fullmatch(body)
    if found is not None:
        first, second, third, fraction, half = found.groups()
        if third is None and fraction is not None:
            hours, minutes, seconds = 0, int(first), int(second)
        else:
            hours, minutes, seconds = int(first), int(second), int(third or 0)
        with_seconds = third is not None
    else:
        found = _HOUR.fullmatch(body)
        if found is None:
            return None
        hours, minutes, seconds, fraction, half, with_seconds = int(found.group(1)), 0, 0, None, found.group(2), False
    if half is not None:
        if hours > 12 or minutes >= 60 or seconds >= 60:
            return None
        hours = hours % 12 + (12 if half[0] in "Pp" else 0)
    # Excel keeps a typed fraction of a second to the millisecond.
    part = round(float(fraction), 3) if fraction else 0
    value = (hours * 3600 + minutes * 60 + seconds + part) / 86400
    if fraction:
        code = "mm:ss.0"
    elif half is not None:
        code = "h:mm:ss AM/PM" if with_seconds else "h:mm AM/PM"
    elif minutes >= 60 or seconds >= 60:
        code = "General"
    elif hours >= 24:
        code = "[h]:mm:ss"
    else:
        code = "h:mm:ss" if with_seconds else "h:mm"
    return Typed(value, code)


# --- what a format makes a value read as -------------------------------------------------------------


def _codes(code: str) -> list[str]:
    """The date and time codes in a number format, in order: y, m, d, h and s runs, elapsed ones included."""
    body = re.sub(r'"[^"]*"|\\.', "", code)
    body = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", body, flags=re.IGNORECASE)
    body = re.sub(r"am/pm|a/p", "", body, flags=re.IGNORECASE)
    return re.findall(r"y+|m+|d+|h+|s+", body.lower())


def is_date_format(code: str) -> bool:
    """Whether a format makes Range.Value read its number as a Date: it names a year, a month or a day.

    A format of times alone -- h:mm, mm:ss, [h]:mm -- leaves the number a
    Double. ``m`` and ``mm`` are minutes straight after an hour or before
    seconds, and a month otherwise.
    """
    if code in ("General", "", "@"):
        return False
    codes = _codes(code)
    for index, found in enumerate(codes):
        if found[0] in "yd":
            return True
        if found[0] == "m":
            after_hour = index > 0 and codes[index - 1][0] == "h"
            before_second = index + 1 < len(codes) and codes[index + 1][0] == "s"
            if len(found) >= 3 or not (after_hour or before_second):
                return True
    return False


def is_currency_format(code: str) -> bool:
    """Whether a format makes Range.Value read its number as a Currency: a dollar sign outside a bracket."""
    return "$" in re.sub(r"\[[^\]]*\]", "", code)
