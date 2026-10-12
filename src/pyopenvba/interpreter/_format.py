"""VBA's Format function, as a US-English Office runs it.

Measured over 245 calls in ``tests/fixtures/vba_format.json``
(``scripts/measure_vba_format.py``) and the hour-token probes recorded with
them. Access's field formats (``pyopenvba.access._format``) are a different
engine: there ``#.##`` pads with zeros and ``\\x`` shows its backslash.

A format is up to four sections split by ``;``. A section with a ``0`` or
``#`` placeholder is a number format; one with ``@`` or ``&`` is a text
format; anything else is a date format, so ``Format(5, "abc")`` reads ``c``
as the general date. Quoted text, a character after ``\\`` and everything a
section does not use are literals; a ``[...]`` directive is dropped.
"""
from __future__ import annotations

import datetime as _dt
from decimal import ROUND_HALF_UP, Decimal, localcontext
from typing import Final

from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBADate,


    date_text,
    error,
    parse_date_text,
    to_number,
    to_text,
)

ERR_OVERFLOW: Final = 6

#: Formats with a name, matched only when the name is the whole format.
NAMED_NUMBER: Final = {
    "currency": "$#,##0.00;($#,##0.00)",
    "fixed": "0.00",
    "standard": "#,##0.00",
    "percent": "0.00%",
    "scientific": "0.00E+00",
}
NAMED_DATE: Final = {
    "general date": "c",
    "long date": "dddd, mmmm d, yyyy",
    "medium date": "dd-mmm-yy",
    "short date": "m/d/yyyy",
    "long time": "h:nn:ss AM/PM",
    "medium time": "hh:nn AM/PM",
    "short time": "hh:nn",
}
NAMED_BOOLEAN: Final = {"yes/no": ("Yes", "No"), "true/false": ("True", "False"), "on/off": ("On", "Off")}

_MONTHS: Final = ("January", "February", "March", "April", "May", "June", "July", "August",
                  "September", "October", "November", "December")
_DAYS: Final = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


def vba_format(value: object, pattern: object = MISSING, first_day: int = 1, first_week: int = 1) -> object:
    """``Format(value, pattern)``: a String, or Null for Null with no Null section."""
    text = "" if pattern is MISSING or pattern is EMPTY else to_text(pattern)
    if value is EMPTY:
        return ""
    sections = _split_sections(text)
    if value is NULL:
        return _literal(sections[3]) if len(sections) >= 4 else ""
    if text == "":
        return to_text(value)
    named = text.lower()
    if named == "general number":
        return to_text(value)
    if named in NAMED_BOOLEAN:
        number = _number_or_none(value)
        if number is None:
            return to_text(value)
        true_text, false_text = NAMED_BOOLEAN[named]
        return true_text if number != 0 else false_text
    if named in NAMED_NUMBER:
        sections = _split_sections(NAMED_NUMBER[named])
    elif named in NAMED_DATE:
        sections = [NAMED_DATE[named]]
    kind = _kind(sections[0])
    if kind == "text":
        return _format_text(value, sections)
    if kind == "number":
        number = _number_or_none(value)
        if number is None:
            return to_text(value)
        return _format_number(number, sections)
    when = _date_or_none(value)
    if when is None:
        return to_text(value)
    return _format_date(when, sections[0], first_day, first_week)


# --- sections and tokens ----------------------------------------------------------------


def _split_sections(text: str) -> list[str]:
    sections: list[str] = []
    current: list[str] = []
    at = 0
    quoted = False
    while at < len(text):
        char = text[at]
        if char == '"':
            quoted = not quoted
        elif char == "\\" and not quoted and at + 1 < len(text):
            current.append(text[at:at + 2])
            at += 2
            continue
        elif char == ";" and not quoted:
            sections.append("".join(current))
            current = []
            at += 1
            continue
        current.append(char)
        at += 1
    sections.append("".join(current))
    return sections


def _scan(section: str) -> list[tuple[str, str]]:
    """A section as ("lit", text) and ("ch", character) pieces, quotes, escapes and [...] taken care of."""
    pieces: list[tuple[str, str]] = []
    at = 0
    while at < len(section):
        char = section[at]
        if char == '"':
            end = section.find('"', at + 1)
            end = len(section) if end < 0 else end
            pieces.append(("lit", section[at + 1:end]))
            at = end + 1
        elif char == "\\" and at + 1 < len(section):
            pieces.append(("lit", section[at + 1]))
            at += 2
        elif char == "[":
            end = section.find("]", at + 1)
            at = len(section) if end < 0 else end + 1
        else:
            pieces.append(("ch", char))
            at += 1
    return pieces


def _kind(section: str) -> str:
    """Number with a 0 or # placeholder; text with @, &, < or >; date with a
    date token or AM/PM; else number, so "." makes 5 into "5." (native probes)."""
    chars = {text for kind, text in _scan(section) if kind == "ch"}
    if chars & {"0", "#"}:
        return "number"
    if chars & {"@", "&", "<", ">"}:
        return "text"
    if any(kind != "lit" for kind, _text in _date_pieces(section)):
        return "date"
    return "number"


def _literal(section: str) -> str:
    return "".join(text for _kind_, text in _scan(section))


# --- values --------------------------------------------------------------------------------


def _number(value: object) -> Decimal:
    if isinstance(value, VBADate):
        return Decimal(repr(value.serial))
    number = to_number(value)
    if isinstance(number, Decimal):
        return number
    if isinstance(number, float):
        return Decimal(f"{number:.15g}")
    return Decimal(int(number))


def _number_or_none(value: object) -> Decimal | None:
    try:
        return _number(value)
    except Exception:  # noqa: BLE001 - text that is not a number is formatted as itself
        return None


def _date_or_none(value: object) -> VBADate | None:
    """The value as a Date, None for text that is not one; a serial past the
    dates VBA holds is error 6 (native probe: 2958466 with "yyyy")."""
    if isinstance(value, VBADate):
        return value
    if isinstance(value, str):
        return parse_date_text(value)
    # The double itself: cut to 15 digits, a time near a half second rounds the
    # other way (tests/fixtures/time_rounding.json).
    number = to_number(value)
    serial = float(number)
    if not -657434.0 <= serial < 2958466.0:
        raise error(ERR_OVERFLOW)
    return VBADate(serial)


# --- numbers -------------------------------------------------------------------------------


def _format_number(value: Decimal, sections: list[str]) -> str:
    with localcontext() as context:
        # A Decimal holds 29 digits, past the default context's 28: abs() and
        # the scaling below would round its last one away.
        context.prec = 60
        return _format_sections(value, sections)


def _format_sections(value: Decimal, sections: list[str]) -> str:
    positive = sections[0]
    negative = sections[1] if len(sections) > 1 and sections[1] != "" else None
    zero = sections[2] if len(sections) > 2 and sections[2] != "" else None
    section = negative if value < 0 and negative is not None else positive
    rendered, rounded_zero = _render_number(abs(value), section)
    if value == 0 or rounded_zero:
        if zero is not None:
            return _render_number(Decimal(0), zero)[0] if _kind(zero) == "number" else _literal(zero)
        return _render_number(Decimal(0), positive)[0]
    if value < 0 and negative is None and rendered:
        return "-" + rendered
    return rendered


def _render_number(value: Decimal, section: str) -> tuple[str, bool]:
    """The section filled with ``value`` (not negative), and whether it rounded to zero."""
    pieces = _scan(section)
    exponent_at = _exponent_index(pieces)
    body = pieces if exponent_at is None else pieces[:exponent_at]
    point = next((index for index, piece in enumerate(body) if piece == ("ch", ".")), None)
    integer_part = body if point is None else body[:point]
    fraction_part = [] if point is None else body[point + 1:]
    digit_indices = [index for index, (kind, char) in enumerate(integer_part) if kind == "ch" and char in "0#"]
    placeholders = "".join(integer_part[index][1] for index in digit_indices)
    decimals = [char for kind, char in fraction_part if kind == "ch" and char in "0#"]
    grouping, scale = False, 0
    if digit_indices:
        first, last = digit_indices[0], digit_indices[-1]
        grouping = any(piece == ("ch", ",") for piece in integer_part[first:last])
        scale = sum(1 for piece in integer_part[last + 1:] if piece == ("ch", ","))
    first_digit = next((index for index, (kind, char) in enumerate(pieces) if kind == "ch" and char in "0#"), None)
    percent = sum(1 for index, piece in enumerate(pieces)
                  if piece == ("ch", "%") and first_digit is not None and index > first_digit)
    value = value * Decimal(100) ** percent / Decimal(1000) ** scale
    if exponent_at is not None:
        return _render_scientific(value, pieces, exponent_at, placeholders, decimals, point)
    rounded = value.quantize(Decimal(1).scaleb(-len(decimals)), rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{rounded:f}".partition(".")
    fraction = fraction.ljust(len(decimals), "0")
    return _fill(pieces, point, whole, fraction, placeholders, decimals, grouping), rounded == 0


def _exponent_index(pieces: list[tuple[str, str]]) -> int | None:
    """Where E+, E-, e+ or e- starts: an E or e not followed by a sign is a literal."""
    for index, (kind, char) in enumerate(pieces[:-1]):
        if kind == "ch" and char in "Ee" and pieces[index + 1][0] == "ch" and pieces[index + 1][1] in "+-":
            return index
    return None


def _render_scientific(value: Decimal, pieces: list[tuple[str, str]], exponent_at: int, placeholders: str,
                       decimals: list[str], point: int | None) -> tuple[str, bool]:
    integer_digits = max(1, len(placeholders))
    if value == 0:
        exponent = 0
        mantissa = Decimal(0)
    else:
        exponent = value.adjusted() - (integer_digits - 1)
        mantissa = value.scaleb(-exponent).quantize(Decimal(1).scaleb(-len(decimals)), rounding=ROUND_HALF_UP)
        if mantissa.adjusted() >= integer_digits:
            exponent += 1
            mantissa = value.scaleb(-exponent).quantize(Decimal(1).scaleb(-len(decimals)), rounding=ROUND_HALF_UP)
    whole, _, fraction = f"{mantissa:f}".partition(".")
    fraction = fraction.ljust(len(decimals), "0")
    marker = pieces[exponent_at][1]
    sign_char = pieces[exponent_at + 1][1]
    exponent_places = [char for kind, char in pieces[exponent_at + 2:] if kind == "ch" and char in "0#"]
    minimum = len([char for char in exponent_places if char == "0"])
    digits = str(abs(exponent)).rjust(minimum, "0")
    sign = "-" if exponent < 0 else ("+" if sign_char == "+" else "")
    exponent_text = marker + sign + digits
    head = _fill(pieces[:exponent_at], point, whole, fraction, placeholders, decimals, False)
    tail = "".join(text for kind, text in pieces[exponent_at + 2:] if kind == "lit" or text not in "0#")
    return head + exponent_text + tail, mantissa == 0


def _fill(pieces: list[tuple[str, str]], point: int | None, whole: str, fraction: str, placeholders: str,
          decimals: list[str], grouping: bool) -> str:
    """The section's pieces with the digits put in their places.

    Integer digits fill the placeholders from the right, the first taking any
    extra; a 0 placeholder pads with zeros up to its place. Grouping puts the
    whole grouped integer at the first placeholder. Commas are never shown:
    they group, scale or, after the point, do nothing (native probes).
    """
    if whole == "0":
        whole = ""
    zero_from = placeholders.find("0")
    minimum = len(placeholders) - zero_from if zero_from >= 0 else 0
    whole = whole.rjust(minimum, "0")
    if grouping and whole:
        whole = f"{int(whole):,}"
    # Fraction: '#' places drop trailing zeros, '0' places keep them.
    keep = len(fraction)
    while keep > 0 and decimals[keep - 1] == "#" and fraction[keep - 1] == "0":
        keep -= 1
    fraction = fraction[:keep]
    integer_count = len(placeholders)
    out: list[str] = []
    integer_seen = fraction_seen = 0
    in_fraction = False
    for index, (kind, char) in enumerate(pieces):
        if point is not None and index == point:
            if integer_count == 0:
                # No integer placeholder: 5 with "." is "5.", 0.5 with ".0" is ".5".
                out.append(whole)
            out.append(".")
            in_fraction = True
            continue
        if kind == "lit":
            out.append(char)
            continue
        if char in "0#":
            if in_fraction:
                if fraction_seen < len(fraction):
                    out.append(fraction[fraction_seen])
                fraction_seen += 1
                continue
            integer_seen += 1
            if grouping:
                if integer_seen == 1:
                    out.append(whole)
                continue
            remaining = integer_count - integer_seen
            if integer_seen == 1:
                take = max(0, len(whole) - remaining)
                out.append(whole[:take])
            else:
                position = len(whole) - remaining - 1
                if position >= 0:
                    out.append(whole[position])
            continue
        if char == ",":
            continue
        out.append(char)
    return "".join(out)


# --- text ----------------------------------------------------------------------------------


def _format_text(value: object, sections: list[str]) -> str:
    text = date_text(value) if isinstance(value, VBADate) else to_text(value)
    section = sections[1] if text == "" and len(sections) > 1 else sections[0]
    pieces = _scan(section)
    chars = [char for kind, char in pieces if kind == "ch"]
    if ">" in chars:
        text = text.upper()
    elif "<" in chars:
        text = text.lower()
    left_to_right = "!" in chars
    slots = [index for index, (kind, char) in enumerate(pieces) if kind == "ch" and char in "@&"]
    count = len(slots)
    if not count:
        # ">" alone upper-cases the text (native probe); its literals follow it.
        return text + "".join(char for kind, char in pieces if kind == "lit" or char not in "<>!")
    fills: dict[int, str] = {}
    if left_to_right:
        if len(text) > count:
            text = text[len(text) - count:]
        for order, index in enumerate(slots):
            fills[index] = text[order] if order < len(text) else (" " if pieces[index][1] == "@" else "")
    else:
        extra = max(0, len(text) - count)
        padded = text[extra:]
        missing = count - len(padded)
        for order, index in enumerate(slots):
            if order < missing:
                fills[index] = " " if pieces[index][1] == "@" else ""
            else:
                fills[index] = padded[order - missing]
        if extra and slots:
            fills[slots[0]] = text[:extra] + fills[slots[0]]
    out: list[str] = []
    for index, (kind, char) in enumerate(pieces):
        if index in fills:
            out.append(fills[index])
        elif kind == "lit" or char not in "<>!":
            out.append(char)
    return "".join(out)


# --- dates ---------------------------------------------------------------------------------

_DATE_TOKENS: Final = ("dddddd", "ddddd", "dddd", "ddd", "dd", "d", "ww", "w", "mmmm", "mmm", "mm", "m", "q",
                       "yyyy", "yy", "y", "hh", "h", "nn", "n", "ss", "s", "ttttt", "c")
_MARKERS: Final = ("am/pm", "a/p", "ampm")


def week_of_year(day: _dt.date, first_day: int = 1, first_week: int = 1) -> int:
    """DatePart("ww"): weeks from the year's first week, as VBA counts them.

    ``first_day`` 1 is Sunday to 7 Saturday; ``first_week`` 1 is the week
    holding 1 January, 2 the first with four days of the year, 3 the first
    whole one. A day before the year's first week is in the last week of
    the year before (1/1/2021 with Monday and four days is week 53).
    """
    def week_one(year: int) -> _dt.date:
        january = _dt.date(year, 1, 1)
        sunday_based = (january.weekday() + 1) % 7 + 1
        start = january - _dt.timedelta(days=(sunday_based - first_day) % 7)
        if first_week == 2 and 7 - (january - start).days < 4:
            start += _dt.timedelta(days=7)
        elif first_week == 3 and start != january:
            start += _dt.timedelta(days=7)
        return start

    start = week_one(day.year)
    if day < start:
        start = week_one(day.year - 1)
    return (day - start).days // 7 + 1


def weekday_number(day: _dt.date, first_day: int = 1) -> int:
    sunday_based = (day.weekday() + 1) % 7 + 1
    return (sunday_based - first_day) % 7 + 1


def _format_date(value: VBADate, section: str, first_day: int, first_week: int) -> str:
    when = value.to_datetime()
    tokens = _date_pieces(section)
    # An AM/PM marker makes the hour before it 12-hour, toggling it if two
    # markers follow it, and every hour after it (native probes).
    twelve: dict[int, bool] = {}
    last_hour: int | None = None
    after_marker = False
    for index, (kind, text) in enumerate(tokens):
        if kind == "tok" and text in ("h", "hh"):
            twelve[index] = after_marker
            last_hour = index
        elif kind == "marker":
            if last_hour is not None:
                twelve[last_hour] = not twelve[last_hour]
            after_marker = True
    out: list[str] = []
    previous: str | None = None
    for index, (kind, text) in enumerate(tokens):
        if kind == "lit":
            out.append(text)
            continue
        if kind == "marker":
            out.append(_marker_text(text, when.hour))
            continue
        out.append(_token_text(text, when, value, previous, twelve.get(index, False), first_day, first_week))
        previous = text
    return "".join(out)


def _date_pieces(section: str) -> list[tuple[str, str]]:
    """Tokens ("tok", lower-case), AM/PM markers ("marker", as written) and literals."""
    out: list[tuple[str, str]] = []
    for kind, text in _scan(section):
        if kind == "lit":
            out.append(("lit", text))
        else:
            out.append(("ch", text))
    joined: list[tuple[str, str]] = []
    at = 0
    while at < len(out):
        kind, text = out[at]
        if kind == "lit":
            joined.append(("lit", text))
            at += 1
            continue
        run = ""
        end = at
        while end < len(out) and out[end][0] == "ch":
            run += out[end][1]
            end += 1
        joined.extend(_date_run(run))
        at = end
    return joined


def _date_run(run: str) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    at = 0
    lower = run.lower()
    while at < len(run):
        marker = next((name for name in _MARKERS if lower.startswith(name, at)), None)
        if marker is not None:
            out.append(("marker", run[at:at + len(marker)]))
            at += len(marker)
            continue
        token = next((name for name in _DATE_TOKENS if lower.startswith(name, at)), None)
        if token is not None:
            out.append(("tok", token))
            at += len(token)
            continue
        out.append(("lit", run[at]))
        at += 1
    return out


def _marker_text(marker: str, hour: int) -> str:
    morning = hour < 12
    lower = marker.lower()
    if lower == "ampm":
        return "AM" if morning else "PM"
    first, _, second = marker.partition("/")
    return first if morning else second


def _token_text(token: str, when: _dt.datetime, value: VBADate, previous: str | None, twelve: bool,
                first_day: int, first_week: int) -> str:
    if token == "d":
        return str(when.day)
    if token == "dd":
        return f"{when.day:02d}"
    if token == "ddd":
        return _DAYS[when.weekday()][:3]
    if token == "dddd":
        return _DAYS[when.weekday()]
    if token == "ddddd":
        return f"{when.month}/{when.day}/{when.year:04d}"
    if token == "dddddd":
        return f"{_DAYS[when.weekday()]}, {_MONTHS[when.month - 1]} {when.day}, {when.year:04d}"
    if token == "w":
        return str(weekday_number(when.date(), first_day))
    if token == "ww":
        return str(week_of_year(when.date(), first_day, first_week))
    if token in ("m", "mm"):
        # A minute right after an hour, a month anywhere else (native probes).
        number = when.minute if previous in ("h", "hh") else when.month
        return str(number) if token == "m" else f"{number:02d}"
    if token == "mmm":
        return _MONTHS[when.month - 1][:3]
    if token == "mmmm":
        return _MONTHS[when.month - 1]
    if token == "q":
        return str((when.month - 1) // 3 + 1)
    if token == "y":
        return str(when.timetuple().tm_yday)
    if token == "yy":
        return f"{when.year % 100:02d}"
    if token == "yyyy":
        return f"{when.year:04d}"
    if token in ("h", "hh"):
        hour = (when.hour % 12 or 12) if twelve else when.hour
        return str(hour) if token == "h" else f"{hour:02d}"
    if token == "n":
        return str(when.minute)
    if token == "nn":
        return f"{when.minute:02d}"
    if token == "s":
        return str(when.second)
    if token == "ss":
        return f"{when.second:02d}"
    if token == "ttttt":
        hour = when.hour % 12 or 12
        return f"{hour}:{when.minute:02d}:{when.second:02d} {'AM' if when.hour < 12 else 'PM'}"
    if token == "c":
        return date_text(value)
    return token

