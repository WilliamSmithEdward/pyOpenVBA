"""What a value written to a cell becomes, and the number format that comes with it.

Writing a string through ``Range.Value`` is typing it: Excel parses it as
it parses what someone types into the cell, on this model's US English
system. Every rule here was measured (scripts/measure_value_typing.py,
scripts/measure_typing_formats.py and scripts/measure_typed_dates.py):

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
  Digits past the fifteenth significant one are cut off, not rounded.
- Dates read month first, with two-digit years before 30 in this century
  and years from 1900 to 9999 on Excel's calendar, which has a 29 February
  1900; a month, by its number or its name, and a number that cannot be
  its day is a month and year; month names, times with or without
  seconds, AM/PM and fractions of a second each bring their own formats.
  A date that does not exist stays text. A run of spaces between the
  parts of a date or a time counts as one space, though not in a fraction.
- A date and a time may come either way round and bring m/d/yyyy h:mm, or
  the time's own format where that is a fraction of a second's or General.
  After them Excel reads past numbers, up to nine numbers and words in
  all. A time AM or PM cannot hold, beside a date, leaves the text with
  the date's format.
- A space before a date or a time keeps it text, except before a day and
  a month's name. Excel misreads a space before a month's name and a
  time, and digits after AM or PM and a point, into numbers no rule
  gives; those are refused.

What the cell already holds decides the rest. A Text cell (``@``) keeps
every string as it is written and turns a Date or a Currency a macro
computes into text. A cell with a number format of its own reads ``1/2``
as a half rather than the second of January. A format typing brings
replaces the cell's own only when that one is General, or one of Excel's
built-in currency, accounting, percent, scientific, fraction, date and
time formats and the new one is of another kind; thousands separators
never replace a format.

Range.Replace types too, with rules of its own (``replaced``), into the
text the formula bar shows for a cell (``edit_text``).
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass
from decimal import Decimal

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import clock, format_value
from pyopenvba.formula._parse import literal
from pyopenvba.formula._values import ERRORS, ExcelError, number_text
from pyopenvba.interpreter._values import EMPTY, VBACurrency, VBADate, VBAErrorValue, error

#: The errors typing turns into error values; #SPILL! and the rest stay text.
_TYPED_ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")
_MONTHS = {name: index for index, names in enumerate(
    (("jan", "january"), ("feb", "february"), ("mar", "march"), ("apr", "april"), ("may",), ("jun", "june"),
     ("jul", "july"), ("aug", "august"), ("sep", "sept", "september"), ("oct", "october"), ("nov", "november"),
     ("dec", "december")), start=1) for name in names}

_PLAIN = re.compile(r"(?P<whole>\d+(?:,\d{3,})*)?(?P<point>\.(?P<fraction>\d*))?")
_EXPONENT = re.compile(r"(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+")
_MIXED = re.compile(r"(\d+) (\d+)/(\d+)")
_FRACTION = re.compile(r"(\d+)/(\d+)")
# The two separators need not match, and spaces may stand round either (tests/fixtures/value_typing.json).
_NUMERIC_DATE = re.compile(r"(\d{1,4}) ?[/-] ?(\d{1,4})(?: ?[/-] ?(\d{1,4}))?")
# Beside a month's name a slash or a dash, a space round either, a space alone or nothing parts the numbers:
# 2/Jan, Jan - 2, 5May2020 (tests/fixtures/typed_dates.json). Only a comma and a space part a year after a day.
_APART = r"(?: ?[/-] ?| )?"
_DAY_MONTH = re.compile(rf"(\d{{1,2}}){_APART}([A-Za-z]+)(?:{_APART}(\d{{1,2}}|\d{{4}}))?")
_MONTH_DAY = re.compile(rf"([A-Za-z]+){_APART}(\d{{1,2}})(?: ?, (\d{{1,2}}|\d{{4}}))?")
_MONTH_YEAR = re.compile(rf"([A-Za-z]+){_APART}(\d{{4}})")
_FIELD = re.compile(r"\d+")
_COLON = re.compile(r" ?: ?")
_POINT = re.compile(r" ?\. ?")
_HALF = re.compile(r" ([AaPp][Mm]?)(?![A-Za-z])")
_HALVES = frozenset(("a", "am", "p", "pm"))
#: The numbers and words a date and a time are read from, nine at most.
_TOKENS = re.compile(r"\d+|[A-Za-z]+")
#: What Excel reads past after a date and a time, once the spaces round slashes and dashes are gone.
_READ_PAST = re.compile(r"[/-]?\d+(?:[/ -]\d+)*[/-]?|[/-]")

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
    """The format a cell takes when typing brings ``brought``, or None where it keeps ``current``.

    Only Replace types into a Text cell, and a format it brings replaces
    Text as it would a built-in format of another kind.
    """
    if brought is None or brought == current:
        return None
    if current in ("General", ""):
        return None if brought == "General" else brought
    held = "text" if current == "@" else _REPLACEABLE.get(current)
    kind = _BROUGHT.get(brought)
    if held is None or kind in (None, "number", held):
        return None
    return brought


def reads_fractions(current: str) -> bool:
    """Whether a cell reads a typed ``1/2`` as a half: it has a number format that is not a date or time."""
    return current not in ("General", "", "@") and not date_time_codes(current)


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
        from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS

        name = next((name for name, number in ERROR_NUMBERS.items() if number == value.number), None)
        if name is None:
            # CVErr of a number no cell error has, CVErr(2044), is refused (tests/fixtures/excel_model/).
            raise error(1004, "Application-defined or object-defined error")
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


def date_serial(year: int, month: int, day: int) -> float | None:
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


def this_year() -> int:
    """The year a date read without one falls in, held where a test holds the one typing uses."""
    return _this_year()


def month_number(name: str) -> int | None:
    """A month's number from its name or abbreviation as typing reads them, Sept included, or None."""
    return _MONTHS.get(name.lower())


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
    found = _number(body, fractions=fractions) or (_after_space(body) if text[0] == " " else _moment(body))
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
        value = literal(text)
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
    value = literal(text.replace(",", ""))
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
    """A typed date, time, or date and time, or None; a run of spaces separates as one space does.

    A value of None is text that still brings a format: a date with a time
    AM or PM cannot hold keeps the text, with the date's format.
    """
    body = re.sub(" {2,}", " ", body)
    if len(_TOKENS.findall(body)) > 9:
        # Excel reads no more than nine numbers and words into a date and a time.
        return None
    time = _clock(body, 0)
    if time is not None and time.end == len(body):
        held = _clock_value(time)
        return None if held is None else Typed(*held)
    date = _date(body)
    if date is not None:
        return date
    # The date first and the time after it, with numbers Excel reads past after that.
    for space in re.finditer(" ", body):
        date = _date(body[:space.start()])
        after = _clock(body, space.end()) if date is not None else None
        if date is not None and after is not None and _read_past(body[after.end:]):
            return _dated(date, after)
    # The time first and the date after it, a slash or a dash between them if there is one.
    if time is not None and body[time.end] in " /-":
        date = _date(re.sub(r"^ ?(?:[/-] ?)?", "", body[time.end:]))
        if date is not None:
            return _dated(date, time)
    return None


def _after_space(body: str) -> Typed | None:
    """A date or time typed after a space: Excel reads only a day and a month's name there, d-mmm.

    Every other date or time with a space before it stays text, except a
    month's name with a time, in a string of nothing but numbers, month's
    names, AM or PM and separators: Excel misreads that into a number, as
    ``" 2-Jan 12:30"`` into 27 December 2011 at 7:20, which is not
    modelled. Words that are none of those keep the whole string text.
    """
    body = re.sub(" {2,}", " ", body)
    found = _DAY_MONTH.fullmatch(body)
    month = _MONTHS.get(found.group(2).lower()) if found is not None and found.group(3) is None else None
    if found is not None and month is not None:
        serial = date_serial(_this_year(), month, int(found.group(1)))
        if serial is not None:
            return Typed(serial, "d-mmm")
    words = [word.lower() for word in re.findall("[A-Za-z]+", body)]
    dated = body.lstrip("/-")[:1].isalnum() and all(word in _MONTHS or word in _HALVES for word in words)
    if dated and any(word in _MONTHS for word in words) and (":" in body or any(word in _HALVES for word in words)):
        raise VBAUnsupportedError(f"typing {' ' + body!r}, a space before a month's name and a time, is not "
                                  "implemented")
    return None


def _dated(date: Typed, time: _Clock) -> Typed:
    """A date and a time typed together: m/d/yyyy h:mm, or the time's own format where that is not a clock's."""
    held = _clock_value(time)
    if held is None:
        # A time AM or PM cannot hold leaves the text, with the date's format.
        return Typed(None, date.number_format)
    value, code = held
    if code != "mm:ss.0":
        # Past 23 hours, or 59 minutes or seconds, the number is General as a time's would be.
        code = "General" if code in ("General", "[h]:mm:ss") else "m/d/yyyy h:mm"
    assert isinstance(date.value, float)
    return Typed(date.value + value, code)


def _read_past(rest: str) -> bool:
    """Whether what follows a date and a time is nothing, or numbers Excel reads past.

    Up to the nine numbers and words a date and time may hold, apart by
    slashes, dashes or spaces, a separator before or after them allowed
    but not two together, each of up to two digits or 100 to 9999.
    """
    if not rest:
        return True
    if rest[0] not in " /-":
        return False
    rest = re.sub(r" ?([/-]) ?", r"\1", rest.lstrip(" "))
    return _READ_PAST.fullmatch(rest) is not None and all(
        len(number) <= 2 or (len(number) <= 4 and int(number) >= 100) for number in re.findall(r"\d+", rest))


def typed_year(text: str) -> int:
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
    first, second, third = found.groups()
    if any(part is not None and len(part) == 3 for part in (first, second, third)):
        # No part of a typed date has three digits: 001/2/2020 is text.
        return None
    if third is not None:
        if len(first) == 4:
            serial = date_serial(int(first), int(second), int(third))
        else:
            serial = date_serial(typed_year(third), int(first), int(second))
        return None if serial is None else (serial, "m/d/yyyy")
    month, number = int(first), int(second)
    if len(first) > 2 or not 1 <= month <= 12:
        return None
    serial = date_serial(_this_year(), month, number) if len(second) <= 2 else None
    if serial is not None:
        return serial, "d-mmm"
    # A number that cannot be the month's day is its year.
    serial = date_serial(typed_year(second), month, 1)
    return None if serial is None else (serial, "mmm-yy")


def _date(body: str) -> Typed | None:
    """A typed date alone: in digits, or with a month's name before or after its day.

    A day first must be one of its month's, and the year after it has one,
    two or four digits. A month's name first takes its year only after a
    comma and a space, and a number that cannot be its day is its year.
    """
    numeric = _numeric_date(body)
    if numeric is not None:
        return Typed(numeric[0], numeric[1])
    found = _DAY_MONTH.fullmatch(body)
    if found is not None:
        return _named(found.group(2), found.group(1), found.group(3), day_first=True)
    found = _MONTH_DAY.fullmatch(body)
    if found is not None:
        return _named(found.group(1), found.group(2), found.group(3), day_first=False)
    found = _MONTH_YEAR.fullmatch(body)
    if found is not None:
        month = _MONTHS.get(found.group(1).lower())
        serial = date_serial(int(found.group(2)), month, 1) if month is not None else None
        if serial is not None:
            return Typed(serial, "mmm-yy")
    return None


def _named(month_text: str, day_text: str, year_text: str | None, *, day_first: bool) -> Typed | None:
    """A date with a month's name, its day and perhaps its year."""
    month = _MONTHS.get(month_text.lower())
    if month is None:
        return None
    serial = date_serial(typed_year(year_text) if year_text else _this_year(), month, int(day_text))
    if serial is not None:
        return Typed(serial, "d-mmm-yy" if year_text else "d-mmm")
    if day_first or year_text is not None:
        return None
    # A number that cannot be the month's day is its year: Jan 45 is January 1945, as 1/45 is.
    serial = date_serial(typed_year(day_text), month, 1)
    return None if serial is None else Typed(serial, "mmm-yy")


@dataclass(frozen=True, slots=True)
class _Clock:
    """A time typed within a longer string, and where it stops."""

    #: The hour, minute and second as far as they are typed; a colon with nothing after it adds none.
    fields: tuple[int, ...]
    #: The digits after a point, where one is typed.
    fraction: str
    #: A, AM, P or PM as typed.
    half: str | None
    end: int


def _clock(text: str, start: int) -> _Clock | None:
    """The time typed at ``start``: up to three numbers apart by colons, a fraction of a second, AM or PM.

    A space may stand either side of a colon or before the point, and a
    space after a colon takes the number that follows; a colon or a point
    with no number after it ends the string. A time has a colon, or AM or
    PM after a space. Each number runs to four digits.
    """
    found = _FIELD.match(text, start)
    if found is None:
        return None
    fields, at, colons = [found.group()], found.end(), 0
    while colons < 2 and (colon := _COLON.match(text, at)) is not None:
        colons, at = colons + 1, colon.end()
        found = _FIELD.match(text, at)
        if found is None:
            if at < len(text):
                return None
            break
        fields.append(found.group())
        at = found.end()
    fraction = ""
    point = _POINT.match(text, at)
    if point is not None:
        found = _FIELD.match(text, point.end())
        if found is None and point.end() < len(text):
            return None
        fraction, at = (found.group(), found.end()) if found is not None else ("", point.end())
    half = _HALF.match(text, at)
    if half is not None:
        at = half.end()
        point = _POINT.match(text, at)
        after = _FIELD.match(text, point.end()) if point is not None else None
        if point is not None and point.end() == len(text):
            # A point after AM or PM with nothing after it is passed over: 10 am.
            at = point.end()
        elif after is not None and _read_past(text[after.end():]):
            raise VBAUnsupportedError(f"typing {text!r}, a fraction of a second after AM or PM, is not implemented")
    if any(len(field) > 4 for field in fields) or not (colons or half is not None) or (fraction and not colons):
        return None
    return _Clock(tuple(int(field) for field in fields), fraction, half.group(1) if half is not None else None, at)


def _clock_value(time: _Clock) -> tuple[float, str] | None:
    """A typed time's day fraction and format, or None where AM or PM cannot hold it.

    h:mm, h:mm:ss, m:ss.0 or h:mm:ss.0, any of them with AM or PM, or an
    hour with AM or PM. Each part runs to 9999; a minute or second past 59
    makes the number General, and an hour past 23 an elapsed time. With AM
    or PM the hour runs from 0 to 12 and the minutes and seconds stop at 59.
    """
    fields = time.fields
    if time.fraction and len(fields) == 2:
        hours, minutes, seconds = 0, *fields
    else:
        hours, minutes, seconds = (*fields, 0, 0)[:3]
    if time.half is not None:
        if hours > 12 or minutes >= 60 or seconds >= 60:
            return None
        hours = hours % 12 + (12 if time.half[0] in "Pp" else 0)
    # Excel keeps a typed fraction of a second to the millisecond.
    part = round(float("0." + time.fraction), 3) if time.fraction else 0
    value = (hours * 3600 + minutes * 60 + seconds + part) / 86400
    if time.fraction:
        code = "mm:ss.0"
    elif time.half is not None:
        code = "h:mm:ss AM/PM" if len(fields) == 3 else "h:mm AM/PM"
    elif minutes >= 60 or seconds >= 60:
        code = "General"
    elif hours >= 24:
        code = "[h]:mm:ss"
    else:
        code = "h:mm:ss" if len(fields) == 3 else "h:mm"
    return value, code


# --- the text a cell is edited as ------------------------------------------------------------------


def edit_text(value: object, number_format: str) -> str:
    """What the formula bar shows for a constant: the text Find looks in among formulas and Replace changes.

    Measured through 20 numbers in 26 formats (scripts/measure_replace.py)
    and 152 times on the half second (scripts/measure_time_rounding.py).
    A number is spelled as Range.Formula spells it, under a percent format
    as a percentage. Under a date or time format it is the date m/d/yyyy on
    Excel's calendar, the time h:mm:ss AM/PM, or both two spaces apart:
    the date where the format has one or the number is a day or more, the
    time where the format has one or the number is not a whole day. The
    time is the day's fraction rounded to the second as a cell rounds one,
    and never carried into the next day. A number no date holds, below 0
    or past 9999, is spelled as a number. Text is itself, without the
    apostrophe it may have been typed with.
    """
    if value is EMPTY:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, ExcelError):
        return value.name
    if isinstance(value, (int, float)):
        number = float(value)
        if date_time_codes(number_format) and 0 <= number < 2958466:
            return _moment_text(number, number_format)
        if _is_percent(number_format):
            return number_text(number * 100, formula=True) + "%"
        return number_text(number, formula=True)
    return str(value)


def _moment_text(number: float, code: str) -> str:
    day = math.floor(number)
    parts: list[str] = []
    if is_date_format(code) or day >= 1:
        parts.append(format_value(float(day), "m/d/yyyy"))
    if any(found[0] in "hsHMS" for found in date_time_codes(code)) or number != day:
        _, hour, minute, second, _ = clock(number - day)
        parts.append(f"{hour % 12 or 12}:{minute:02d}:{second:02d} {'AM' if hour < 12 else 'PM'}")
    return "  ".join(parts)


def _is_percent(code: str) -> bool:
    """Whether a format multiplies by 100: a % sign that is not quoted, escaped, bracketed or a padding's."""
    return "%" in re.sub(r'"[^"]*"|\\.|_.|\*.|\[[^\]]*\]', "", code)


def replaced(text: str, current: str, *, prefixed: bool) -> Typed:
    """What Replace leaves in a cell whose text it changed, the cell's format being ``current``.

    Replace types the new text as it would be typed into a General cell:
    a Text cell's string becomes a number, a date or a Boolean too, and
    ``1/2`` is the second of January even where the cell's format is a
    number's. The format the text brings then goes where typing's would,
    replacing Text as it replaces a built-in format of another kind. A
    cell that showed a prefix character keeps it, and its text stays text.
    """
    if prefixed:
        return typed_text("'" + text)
    found = typed_text(text)
    return Typed(found.value, format_after(current, found.number_format), found.prefix)


# --- what a format makes a value read as -------------------------------------------------------------


def date_time_codes(code: str) -> list[str]:
    """The date and time codes in a number format, in order: y, m, d, h and s runs.

    An elapsed [h], [m] or [s] comes back as H, M or S, since [m] is
    minutes wherever it stands.
    """
    body = re.sub(r'"[^"]*"|\\.', "", code)
    body = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", body, flags=re.IGNORECASE)
    body = re.sub(r"am/pm|a/p", "", body, flags=re.IGNORECASE).lower()
    body = re.sub(r"\[([hms])+\]", lambda found: found.group(1).upper(), body)
    return re.findall(r"y+|m+|d+|h+|s+|[HMS]", body)


def is_date_format(code: str) -> bool:
    """Whether a format makes Range.Value read its number as a Date: it names a year, a month or a day.

    A format of times alone -- h:mm, mm:ss, [h]:mm, [m] -- leaves the
    number a Double. ``m`` and ``mm`` are minutes straight after an hour
    or before seconds, and a month otherwise.
    """
    if code in ("General", "", "@"):
        return False
    codes = date_time_codes(code)
    for index, found in enumerate(codes):
        if found[0] in "yd":
            return True
        if found[0] == "m":
            after_hour = index > 0 and codes[index - 1][0] in "hH"
            before_second = index + 1 < len(codes) and codes[index + 1][0] in "sS"
            if len(found) >= 3 or not (after_hour or before_second):
                return True
    return False


def is_currency_format(code: str) -> bool:
    """Whether a format makes Range.Value read its number as a Currency: a dollar sign outside a bracket."""
    return "$" in re.sub(r"\[[^\]]*\]", "", code)
