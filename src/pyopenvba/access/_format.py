"""Jet's ``Format`` and ``Partition``.

Both answer with text, and both were measured against DAO on a table of
dates, doubles and currency: the named formats, the custom date and
number patterns, the negative section, and the widths ``Partition`` pads
its bounds to.  ``Format`` of Null is the empty string, not Null.
"""

from __future__ import annotations

import datetime as _dt
import re
from decimal import ROUND_HALF_UP, Decimal

from pyopenvba.access_read import AccessError

WEEKDAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
# Longest first: the scanner takes the first that matches.
DATE_TOKENS = ("yyyy", "yy", "mmmm", "mmm", "mm", "dddd", "ddd", "dd", "ww", "hh", "nn", "ss", "m", "d", "q", "y", "w", "h", "n", "s")
AMPM_TOKENS = ("am/pm", "a/p", "ampm")


def format_value(value: object, pattern: str) -> str:
    """``Format(value, pattern)``.  Null formats as the empty string."""
    if value is None:
        return ""
    named = pattern.strip().lower()
    if named in ("<", ">"):
        text = _as_text(value)
        return text.lower() if named == "<" else text.upper()
    if named in NAMED_DATES:
        return _date_pattern(_as_date(value), NAMED_DATES[named])
    if named == "general date":
        when = _as_date(value)
        if when.time() == _dt.time():
            return _date_pattern(when, "m/d/yyyy")
        if when.date() == _dt.date(1899, 12, 30):
            return _date_pattern(when, "h:nn:ss AM/PM")
        return _date_pattern(when, "m/d/yyyy h:nn:ss AM/PM")
    if named in NAMED_NUMBERS:
        return _number_pattern(_as_number(value), NAMED_NUMBERS[named])
    if named == "general number":
        return _plain(_as_number(value))
    if named == "scientific":
        return f"{float(_as_number(value)):.2E}"
    if named in TWO_WAY:
        yes, no = TWO_WAY[named]
        return no if _as_number(value) == 0 else yes
    if "@" in pattern or "&" in pattern:
        return _text_pattern(_as_text(value), pattern)
    if isinstance(value, (_dt.datetime, _dt.date)) or _looks_like_a_date(pattern):
        return _date_pattern(_as_date(value), pattern)
    return _number_pattern(_as_number(value), pattern)


NAMED_DATES = {
    "long date": "dddd, mmmm d, yyyy",
    "medium date": "dd-mmm-yy",
    "short date": "m/d/yyyy",
    "long time": "h:nn:ss AM/PM",
    "medium time": "hh:nn AM/PM",
    "short time": "hh:nn",
}
NAMED_NUMBERS = {
    "currency": "$#,##0.00;($#,##0.00)",
    "fixed": "0.00",
    "standard": "#,##0.00",
    "percent": "0.00%",
}
TWO_WAY = {"yes/no": ("Yes", "No"), "true/false": ("True", "False"), "on/off": ("On", "Off")}


def partition(number: float, start: int, stop: int, interval: int) -> str:
    """The range ``number`` falls in, as ``low:high``, each bound padded to
    the width of the widest one the arguments allow.  Below ``start`` the
    low bound is blank, above ``stop`` the high one is."""
    if interval < 1:
        raise AccessError("Partition needs an interval of at least 1")
    width = max(len(str(stop)), len(str(start - 1)))
    if number < start:
        return " " * width + ":" + str(start - 1).rjust(width)
    if number > stop:
        return str(stop + 1).rjust(width) + ":" + " " * width
    offset = int((number - start) // interval)
    low = start + offset * interval
    return str(low).rjust(width) + ":" + str(min(low + interval - 1, stop)).rjust(width)


def _text_pattern(text: str, pattern: str) -> str:
    """A text pattern: ``@`` takes one character or a space where there is
    none, ``&`` takes one or nothing, and everything else is literal.
    Both are filled from the right, as VBA fills them, and an empty
    string formats as an empty string whatever the pattern says
    (measured)."""
    if not text:
        return ""
    letters = list(text)
    out: list[str] = []
    for mark in reversed(pattern):
        if mark == "@":
            out.append(letters.pop() if letters else " ")
        elif mark == "&":
            out.append(letters.pop() if letters else "")
        else:
            out.append(mark)
    return "".join(letters) + "".join(reversed(out))


def _as_text(value: object) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    return str(value)


def _as_number(value: object) -> Decimal:
    if isinstance(value, bool):
        return Decimal(-1 if value else 0)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    if isinstance(value, _dt.datetime):
        delta = value - _dt.datetime(1899, 12, 30)
        return Decimal(delta.days) + Decimal(delta.seconds) / Decimal(86400)
    try:
        return Decimal(str(value).strip())
    except Exception as exc:  # noqa: BLE001
        raise AccessError(f"Format cannot read {value!r} as a number") from exc


def _as_date(value: object) -> _dt.datetime:
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime.combine(value, _dt.time())
    if isinstance(value, (int, float, Decimal)):
        whole = int(value)
        return _dt.datetime(1899, 12, 30) + _dt.timedelta(days=whole, seconds=round(abs(float(value) - whole) * 86400))
    from pyopenvba.access._sql import parse_date_literal

    return parse_date_literal(str(value))


def _looks_like_a_date(pattern: str) -> bool:
    return bool(re.search(r"(yyyy|mmmm|mmm|dddd|ddd|dd|hh|nn|ss)", pattern, re.IGNORECASE))


def _plain(number: Decimal) -> str:
    text = format(number.normalize(), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def _number_pattern(number: Decimal, pattern: str) -> str:
    """One custom number pattern, with up to four ``;``-separated sections
    for positive, negative, zero and null values.  A negative value
    handed to its own section loses its sign, as the section is expected
    to spell one out."""
    sections = pattern.split(";")
    if number < 0 and len(sections) > 1 and sections[1]:
        return _one_section(-number, sections[1])
    if number == 0 and len(sections) > 2 and sections[2]:
        return _one_section(number, sections[2])
    return _one_section(number, sections[0])


def _one_section(number: Decimal, section: str) -> str:
    if "%" in section:
        number = number * 100
    body = section.replace("%", "")
    grouped = "," in body
    body = body.replace(",", "")
    if not any(c in body for c in "0#"):
        return section
    head, _, tail = body.partition(".")
    decimals = sum(1 for c in tail if c in "0#")
    quantized = number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    sign = "-" if quantized < 0 else ""
    digits = format(abs(quantized), "f")
    whole, _, fraction = digits.partition(".")
    zeros = sum(1 for c in head if c == "0")
    whole = whole.lstrip("0") or ("" if "0" not in head and zeros == 0 else "0")
    whole = whole.rjust(zeros, "0")
    if grouped:
        whole = f"{int(whole or 0):,}" if whole else whole
    out = whole + ("." + fraction.ljust(decimals, "0") if decimals else "")
    prefix, suffix = _fixed_parts(section)
    return prefix + sign + out + suffix


def _fixed_parts(section: str) -> tuple[str, str]:
    """Whatever a section says before its first placeholder and after its
    last, which is where a currency sign or a bracket sits."""
    positions = [i for i, c in enumerate(section) if c in "0#"]
    if not positions:
        return "", ""
    return section[: positions[0]].replace("%", ""), section[positions[-1] + 1 :].replace(",", "")


def _date_pattern(when: _dt.datetime, pattern: str) -> str:
    twelve = any(t in pattern.lower() for t in AMPM_TOKENS)
    out: list[str] = []
    i = 0
    previous = ""
    while i < len(pattern):
        rest = pattern[i:]
        lowered = rest.lower()
        marker = next((t for t in AMPM_TOKENS if lowered.startswith(t)), None)
        if marker is not None:
            morning = when.hour < 12
            if marker == "a/p":
                out.append(("a" if morning else "p") if rest[0].islower() else ("A" if morning else "P"))
            else:
                out.append(("AM" if morning else "PM") if rest[0].isupper() else ("am" if morning else "pm"))
            i += len(marker)
            continue
        token = next((t for t in DATE_TOKENS if lowered.startswith(t)), None)
        if token is None:
            if rest[0] == '"':
                end = rest.find('"', 1)
                out.append(rest[1:end] if end > 0 else rest[1:])
                i += (end + 1) if end > 0 else len(rest)
                continue
            if rest[0] == chr(92) and len(rest) > 1:
                out.append(rest[1])
                i += 2
                continue
            out.append(rest[0])
            i += 1
            continue
        out.append(_date_token(token, when, twelve, previous))
        previous = token
        i += len(token)
    return "".join(out)


def _date_token(token: str, when: _dt.datetime, twelve: bool, previous: str) -> str:
    hour = when.hour % 12 or 12 if twelve else when.hour
    if token == "yyyy":
        return f"{when.year:04d}"
    if token == "yy":
        return f"{when.year % 100:02d}"
    if token == "mmmm":
        return MONTHS[when.month - 1]
    if token == "mmm":
        return MONTHS[when.month - 1][:3]
    if token in ("mm", "m"):
        # After an hour token, m means minutes, which is VBA's rule.
        value = when.minute if previous in ("h", "hh") else when.month
        return f"{value:02d}" if token == "mm" else str(value)
    if token == "dddd":
        return WEEKDAYS[(when.weekday() + 1) % 7]
    if token == "ddd":
        return WEEKDAYS[(when.weekday() + 1) % 7][:3]
    if token == "dd":
        return f"{when.day:02d}"
    if token == "d":
        return str(when.day)
    if token == "q":
        return str((when.month - 1) // 3 + 1)
    if token == "y":
        return str(when.timetuple().tm_yday)
    if token == "ww":
        first = _dt.date(when.year, 1, 1)
        sunday = first - _dt.timedelta(days=(first.weekday() + 1) % 7)
        return str((when.date() - sunday).days // 7 + 1)
    if token == "w":
        return str((when.weekday() + 1) % 7 + 1)
    if token == "hh":
        return f"{hour:02d}"
    if token == "h":
        return str(hour)
    if token == "nn":
        return f"{when.minute:02d}"
    if token == "n":
        return str(when.minute)
    if token == "ss":
        return f"{when.second:02d}"
    return str(when.second)
