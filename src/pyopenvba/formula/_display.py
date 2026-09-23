"""What a number format shows: Excel's formats, as TEXT shows a value and a wide enough cell does.

Measured in live Excel (scripts/measure_number_formats.py): 156 codes
through 28 values, read through Range.Text in a column wide enough not to
cut anything short and through WorksheetFunction.Text. The two agree
everywhere but where the width decides -- a ``*`` fill, which pads the
cell and which TEXT drops, and a value the format cannot show, which
fills the cell with ``#`` and makes TEXT an error.

- A code has up to four sections, for positive numbers, negative ones,
  zero and text. With one section a negative number gets a minus sign;
  with two or three the negative section shows the magnitude. A section
  with a condition, ``[>=100]``, is chosen by it, and then every number
  keeps its sign. A minus sign leads everything the section writes, and
  a number that rounds to nothing shows none, except as a mixed fraction.
- A number is first taken to fifteen significant digits, and rounded
  half away from zero to the places shown. ``0`` shows a digit, ``#``
  shows nothing for a zero that does not matter and ``?`` a space; a
  comma between digits groups thousands, a comma after the last digit
  divides by 1,000, and ``%`` multiplies by 100.
- ``E+`` or ``E-`` shows an exponent: a multiple of the number of digits
  before the point, with at least as many digits as follow the sign.
- A fraction takes the last continued-fraction convergent whose
  denominator has no more digits than its placeholders, or rounds to a
  denominator written out; a fraction of nothing leaves its places blank.
- A date or time reads the serial on Excel's calendar, day 0 being 0
  January 1900 and day 60 the 29 February 1900 Excel counts. Half a
  second, or half the fraction of one shown, is added to the serial
  before the day, hours, minutes and seconds are cut off it one by one
  (scripts/measure_time_rounding.py); a count of elapsed seconds alone is
  rounded from fifteen digits instead. A serial below 0 or past 31
  December 9999 cannot be shown.
- ``@`` shows text; a code with no text section shows text as it is, and
  a Boolean as TRUE or FALSE, through the text section when there is one.
- ``_x`` leaves a space the width of ``x``, which in text is a space.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from collections.abc import Iterator
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, localcontext
from fractions import Fraction


class UndisplayableError(Exception):
    """A value its format cannot show: TEXT answers #VALUE!, and a cell fills with #."""


# --- tokens ------------------------------------------------------------------------------------------


def tokens(code: str) -> Iterator[tuple[str, str]]:
    """A code in pieces: ("quoted", '"text"'), ("escaped", "\\x"), ("bracket", "[...]"), ("pair", "_x" or "*x"),
    ("exponent", "E+"), or ("char", c)."""
    index = 0
    while index < len(code):
        char = code[index]
        if char == '"':
            end = code.find('"', index + 1)
            end = len(code) - 1 if end < 0 else end
            yield "quoted", code[index:end + 1]
            index = end + 1
        elif char == "\\" and index + 1 < len(code):
            yield "escaped", code[index:index + 2]
            index += 2
        elif char == "[":
            end = code.find("]", index)
            end = len(code) - 1 if end < 0 else end
            yield "bracket", code[index:end + 1]
            index = end + 1
        elif char in "_*" and index + 1 < len(code):
            yield "pair", code[index:index + 2]
            index += 2
        elif char in "Ee" and code[index + 1:index + 2] in ("+", "-"):
            yield "exponent", code[index:index + 2]
            index += 2
        else:
            yield "char", char
            index += 1


# --- General -----------------------------------------------------------------------------------------


def general_text(number: float) -> str:
    """A number as the General format shows it in a column wide enough not to cut it short.

    General spends at most eleven characters, the sign apart. A number is
    written out in full when that fits; otherwise it is rounded to fit,
    or given an exponent, whichever keeps more significant digits, with a
    tie going to the rounding -- so 0.0833333333333333 shows 0.083333333,
    0.000012345678 shows 1.23457E-05, and 123456789012 shows 1.23457E+11.
    Measured from 1E-25 to 1E+300 (scripts/measure_number_spelling.py). A
    narrower column shows less, which is not modelled.
    """
    if number != number or number in (float("inf"), float("-inf")):
        return "#NUM!"
    if number == 0:
        return "0"
    widest = 11
    sign = "-" if number < 0 else ""
    exact = _significant(abs(number)).normalize()
    if len(plain := format(exact, "f")) <= widest:
        return sign + plain
    _, digits, exponent = exact.as_tuple()
    assert isinstance(exponent, int)
    highest = exponent + len(digits) - 1
    scientific = _exponent_form(exact, widest)
    in_exponent_form = len(scientific.partition("E")[0].replace(".", ""))
    # What rounding to fit keeps: the whole digits and as many places as the width leaves.
    places = widest - (highest + 2) if highest >= 0 else widest - 2
    kept = (highest + 1 + max(places, 0)) if highest >= 0 else places + highest + 1
    if highest + 1 <= widest and kept >= in_exponent_form:
        text = format(_rounded(exact, max(places, 0)).normalize(), "f")
        if len(text) <= widest:
            return sign + text
    return sign + scientific


def _exponent_form(number: Decimal, widest: int) -> str:
    """A positive number with an exponent of two digits at least, its mantissa cut to fit ``widest``."""
    for precision in range(15, 0, -1):
        with localcontext() as context:
            context.prec = precision
            context.rounding = ROUND_HALF_UP
            rounded = (+number).normalize()
        _, digits, exponent = rounded.as_tuple()
        assert isinstance(exponent, int)
        highest = exponent + len(digits) - 1
        mantissa = "".join(str(digit) for digit in digits)
        text = (mantissa[0] + ("." + mantissa[1:] if len(mantissa) > 1 else "")
                + f"E{'+' if highest >= 0 else '-'}{abs(highest):02d}")
        if len(text) <= widest:
            return text
    raise AssertionError("a one-digit mantissa always fits")


# --- sections ----------------------------------------------------------------------------------------

_CONDITION = re.compile(r"\[(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d*)?)\]")
_CURRENCY = re.compile(r"\[\$([^\]-]*)(?:-[^\]]*)?\]")
_ELAPSED = re.compile(r"\[(h+|m+|s+)\]", re.IGNORECASE)
_DATE_LETTERS = frozenset("ymdhsYMDHS")
_PAD = {"0": "0", "?": " ", "#": ""}


@dataclass(slots=True)
class _Item:
    """One element of a section: its kind -- digit, point, comma, percent, slash, exponent, general, text,
    date, elapsed, ampm or literal -- and the characters it stands for."""

    kind: str
    text: str = ""


@dataclass(slots=True)
class _Section:
    items: list[_Item]
    condition: tuple[str, float] | None
    has_fill: bool

    def has(self, *kinds: str) -> bool:
        return any(item.kind in kinds for item in self.items)


def _section(pieces: list[tuple[str, str]]) -> _Section:
    items: list[_Item] = []
    condition: tuple[str, float] | None = None
    fill = False
    chars: list[str] = []

    def flush() -> None:
        items.extend(_letters("".join(chars)))
        chars.clear()

    for kind, text in pieces:
        if kind == "char":
            chars.append(text)
            continue
        flush()
        if kind == "quoted":
            items.append(_Item("literal", text[1:-1]))
        elif kind == "escaped":
            items.append(_Item("literal", text[1]))
        elif kind == "pair":
            if text[0] == "_":
                items.append(_Item("literal", " "))
            else:
                fill = True
        elif kind == "exponent":
            items.append(_Item("exponent", text))
        elif (found := _CONDITION.fullmatch(text)) is not None:
            condition = (found.group(1), float(found.group(2)))
        elif (found := _CURRENCY.fullmatch(text)) is not None:
            items.append(_Item("literal", found.group(1)))
        elif (found := _ELAPSED.fullmatch(text)) is not None:
            items.append(_Item("elapsed", found.group(1).lower()))
        # Any other bracket -- a colour, a locale -- shows nothing.
    flush()
    return _Section(items, condition, fill)


def _letters(run: str) -> Iterator[_Item]:
    """A run of plain characters as the parts it makes: digits, General, AM/PM, date letters and literals."""
    index = 0
    while index < len(run):
        lowered = run[index:].lower()
        char = run[index]
        if lowered.startswith("general"):
            yield _Item("general")
            index += 7
        elif lowered.startswith("am/pm"):
            yield _Item("ampm", run[index:index + 5])
            index += 5
        elif lowered.startswith("a/p"):
            yield _Item("ampm", run[index:index + 3])
            index += 3
        elif char in "0#?":
            yield _Item("digit", char)
            index += 1
        elif char in _DATE_LETTERS:
            count = len(run[index:]) - len(run[index:].lstrip(char.lower() + char.upper()))
            yield _Item("date", char.lower() * count)
            index += count
        elif char in "eE":
            # e without a sign after it is the era year, which in this locale is the year: TEXT(0,"zero") shows
            # z1900ro (pyOfficeEditor's formula corpus).
            count = len(run[index:]) - len(run[index:].lstrip("eE"))
            yield _Item("date", "yyyy")
            index += count
        elif char in ".,%/@":
            yield _Item({".": "point", ",": "comma", "%": "percent", "/": "slash", "@": "text"}[char], char)
            index += 1
        else:
            yield _Item("literal", char)
            index += 1


def _sections(code: str) -> list[_Section]:
    pieces: list[list[tuple[str, str]]] = [[]]
    for kind, text in tokens(code):
        if kind == "char" and text == ";":
            pieces.append([])
        else:
            pieces[-1].append((kind, text))
    return [_section(piece) for piece in pieces]


def _meets(condition: tuple[str, float], value: float) -> bool:
    operator, bound = condition
    return {"<": value < bound, "<=": value <= bound, ">": value > bound, ">=": value >= bound,
            "=": value == bound, "<>": value != bound}[operator]


# --- the value, through its section ------------------------------------------------------------------


def format_value(value: object, code: str) -> str:
    """What ``code`` shows for ``value`` -- a number, a string or a Boolean -- as TEXT does.

    Raises UndisplayableError where the format cannot show the value.
    """
    return shown(value, code)[0]


def shown(value: object, code: str) -> tuple[str, bool]:
    """What ``code`` shows for ``value``, and whether the section it shows through fills the cell's width.

    A ``*`` fill repeats its character across whatever width the cell has
    to spare, which TEXT leaves out; the text returned leaves it out too.
    """
    sections = _sections(code or "General")
    text_section = sections[3] if len(sections) >= 4 else sections[-1] if sections[-1].has("text") else None
    if isinstance(value, (bool, str)):
        text = ("TRUE" if value else "FALSE") if isinstance(value, bool) else value
        if text_section is None:
            return text, False
        return "".join(text if item.kind == "text" else _literal(item) for item in text_section.items), \
            text_section.has_fill
    number = float(value)  # type: ignore[arg-type]
    numeric = [section for section in sections[:3] if section is not text_section]
    if not numeric:
        return general_text(number), False
    section, magnitude, signed = _choose(numeric, number)
    return _numeric(section, magnitude, signed), section.has_fill


def _choose(sections: list[_Section], value: float) -> tuple[_Section, float, bool]:
    """The section a number shows through, the number it shows, and whether a minus sign may lead it."""
    if any(section.condition is not None for section in sections):
        # The first section whose condition holds; a later section with none takes whatever is left.
        for index, section in enumerate(sections):
            if section.condition is None:
                if index:
                    return section, value, True
            elif _meets(section.condition, value):
                return section, value, True
        return sections[-1], value, True
    if len(sections) == 1:
        return sections[0], value, True
    if value < 0:
        return sections[1], -value, False
    if value == 0 and len(sections) >= 3:
        return sections[2], value, False
    return sections[0], value, False


def _literal(item: _Item) -> str:
    """What an item shows when it is not filled with a value."""
    return item.text if item.kind in ("literal", "percent", "slash", "comma", "point", "date", "ampm") else ""


def _numeric(section: _Section, value: float, signed: bool) -> str:
    if section.has("date", "elapsed", "ampm"):
        return _date(section.items, value)
    if section.has("general"):
        shown = general_text(value if signed else abs(value))
        return "".join(shown if item.kind == "general" else _literal(item) for item in section.items)
    if not section.has("digit"):
        return "".join(_literal(item) for item in section.items)
    if section.has("exponent"):
        text, zero = _scientific(section.items, abs(value))
    elif _slash(section.items) is not None:
        text, zero = _fraction(section.items, abs(value))
    else:
        text, zero = _decimal(section.items, abs(value))
    return ("-" if signed and value < 0 and not zero else "") + text


# --- numbers ---------------------------------------------------------------------------------------


def _significant(value: float) -> Decimal:
    """A number as Excel holds it for display: fifteen significant digits."""
    with localcontext() as context:
        context.prec = 15
        return +Decimal(value)


def _rounded(value: Decimal, places: int) -> Decimal:
    """Half away from zero to ``places``, however large the number."""
    with localcontext() as context:
        context.prec = 1000
        return value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)


def _whole(items: list[_Item], places: list[int], digits: str, grouped: bool) -> dict[int, str]:
    """Whole digits right-aligned over their placeholders, what is left over going to the first.

    A placeholder with no digit of its own shows 0 for ``0``, a space for
    ``?`` and nothing for ``#``. Grouping puts a comma every three digits,
    padding zeros included.
    """
    shown: dict[int, str] = {}
    remaining = digits
    for count, index in enumerate(reversed(places)):
        if count == len(places) - 1:
            piece, remaining = remaining, ""
        else:
            piece, remaining = remaining[-1:], remaining[:-1]
        shown[index] = piece or _PAD[items[index].text]
    if grouped and places:
        joined = "".join(shown[index] for index in places)
        body = joined.lstrip(" ")
        groups: list[str] = []
        while body:
            groups.append(body[-3:])
            body = body[:-3]
        shown = dict.fromkeys(places, "")
        shown[places[0]] = joined[: len(joined) - len(joined.lstrip(" "))] + ",".join(reversed(groups))
    return shown


def _places(items: list[_Item], places: list[int], fraction: str) -> dict[int, str]:
    """Decimal digits left to right; a trailing zero shows 0 for ``0``, nothing for ``#`` and a space for ``?``."""
    significant = len(fraction.rstrip("0"))
    return {index: (_PAD[items[index].text] if position >= significant else fraction[position])
            for position, index in enumerate(places)}


def _decimal(items: list[_Item], value: float) -> tuple[str, bool]:
    """A number through digit placeholders: the text, and whether it rounded to nothing."""
    digits = [index for index, item in enumerate(items) if item.kind == "digit"]
    point = next((index for index, item in enumerate(items) if item.kind == "point"), len(items))
    grouped, scale, percent = False, 0, 0
    kept: list[_Item] = []
    for index, item in enumerate(items):
        if item.kind == "comma" and index > digits[-1]:
            scale += 1
        elif item.kind == "comma" and digits[0] < index < point:
            grouped = True
        else:
            percent += item.kind == "percent"
            kept.append(item if item.kind != "comma" else _Item("literal", ","))
    point = next((index for index, item in enumerate(kept) if item.kind == "point"), len(kept))
    whole_places = [index for index in range(point) if kept[index].kind == "digit"]
    places = [index for index in range(point + 1, len(kept)) if kept[index].kind == "digit"]
    number = _rounded(_significant(value) * Decimal(100) ** percent / Decimal(1000) ** scale, len(places))
    whole, _, fraction = format(number, "f").partition(".")
    filled = {**_whole(kept, whole_places, whole.lstrip("0"), grouped),
              **_places(kept, places, fraction.ljust(len(places), "0"))}
    if not whole_places and whole.lstrip("0"):
        # With no place for them, whole digits still show, straight before the point.
        filled[point] = whole.lstrip("0") + "."
    return "".join(filled.get(index, _literal(item)) for index, item in enumerate(kept)), number == 0


def _scientific(items: list[_Item], value: float) -> tuple[str, bool]:
    """A number with an exponent: a multiple of the whole placeholders before the point."""
    at = next(index for index, item in enumerate(items) if item.kind == "exponent")
    mantissa_items, exponent_items = items[:at], items[at + 1:]
    point = next((index for index, item in enumerate(mantissa_items) if item.kind == "point"), len(mantissa_items))
    whole_places = [index for index in range(point) if mantissa_items[index].kind == "digit"]
    places = [index for index in range(point + 1, len(mantissa_items)) if mantissa_items[index].kind == "digit"]
    step = max(len(whole_places), 1)
    number = _significant(value)
    exponent, mantissa = 0, Decimal(0)
    if number:
        exponent = number.adjusted() // step * step
        mantissa = _rounded(number.scaleb(-exponent), len(places))
        if mantissa >= Decimal(10) ** step:
            exponent += step
            mantissa = _rounded(number.scaleb(-exponent), len(places))
    whole, _, fraction = format(mantissa, "f").partition(".")
    # Zero shows a 0 in every whole place.
    shown = dict.fromkeys(whole_places, "0") if not number else \
        _whole(mantissa_items, whole_places, whole.lstrip("0"), grouped=False)
    filled = {**shown, **_places(mantissa_items, places, fraction.ljust(len(places), "0"))}
    head = "".join(filled.get(index, _literal(item)) for index, item in enumerate(mantissa_items))
    letter, sign = items[at].text
    written = "-" if exponent < 0 else ("+" if sign == "+" else "")
    width = max(sum(item.kind == "digit" for item in exponent_items), 1)
    tail = "".join(_literal(item) for item in exponent_items if item.kind != "digit")
    return f"{head}{letter}{written}{str(abs(exponent)).rjust(width, '0')}{tail}", not number


# --- fractions ------------------------------------------------------------------------------------


def _slash(items: list[_Item]) -> int | None:
    """Where a fraction's slash is: a / with a digit placeholder straight before it."""
    return next((index for index, item in enumerate(items)
                 if item.kind == "slash" and index > 0 and items[index - 1].kind == "digit"), None)


def _convergent(value: float, limit: int) -> Fraction:
    """The fraction Excel shows for ``value``: the last continued-fraction convergent whose denominator fits.

    Excel works the continued fraction out in double arithmetic on the
    binary number, and takes no semiconvergents: 0.5625 over one digit is
    4/7, not the closer 5/9; the fraction of 1234567.891, a hair over 0.891
    in binary, is 703/789 over three; and 0.001 over three is nothing,
    since 1/0.001 comes out exactly 1000.
    """
    before, current = (0, 1), (1, 0)
    remainder = value
    while True:
        term = math.floor(remainder)
        numerator, denominator = term * current[0] + before[0], term * current[1] + before[1]
        if denominator > limit:
            return Fraction(current[0], current[1])
        before, current = current, (numerator, denominator)
        left = remainder - term
        if left <= 1e-300:
            return Fraction(numerator, denominator)
        remainder = 1 / left


def _fraction(items: list[_Item], value: float) -> tuple[str, bool]:
    """A number as a fraction, with a whole part before it when the code has one."""
    slash = _slash(items)
    assert slash is not None
    top = slash
    while top > 0 and items[top - 1].kind == "digit":
        top -= 1
    numerator_places = list(range(top, slash))
    whole_places = [index for index in range(top) if items[index].kind == "digit"]
    between = list(range(whole_places[-1] + 1, top)) if whole_places else []
    end = slash + 1
    while end < len(items) and (items[end].kind == "digit" or items[end].text.isdigit()):
        end += 1
    below = list(range(slash + 1, end))
    fixed = "".join(items[index].text for index in below) if below and items[below[0]].kind == "literal" else ""
    whole = math.floor(value) if whole_places else 0
    if fixed:
        rest = _significant(value) - whole
        numerator, denominator = int(_rounded(rest * int(fixed), 0)), int(fixed)
    else:
        found = _convergent(value - whole, 10 ** len(below) - 1)
        numerator, denominator = found.numerator, found.denominator
    if whole_places and numerator >= denominator:
        whole, numerator = whole + numerator // denominator, numerator % denominator
    if not whole_places and len(str(numerator)) > 15:
        raise UndisplayableError("a numerator too long to show")
    filled: dict[int, str] = {}
    if whole_places and not numerator:
        # A fraction of nothing: the whole number, then blanks where the fraction would be.
        filled.update(_whole(items, whole_places, str(whole), grouped=False))
        for index in [*between, *numerator_places, slash]:
            filled[index] = " " * max(len(_literal(items[index])), 1)
        for index in below:
            filled[index] = " "
    else:
        filled.update(_whole(items, whole_places, str(whole) if whole else "", grouped=False))
        head = items[numerator_places[0]].text
        width = len(numerator_places)
        filled.update(dict.fromkeys(numerator_places, ""))
        filled[numerator_places[0]] = str(numerator).rjust(width, " " if head == "?" else "0") \
            if head in "?0" else str(numerator)
        if fixed:
            # A denominator written out shows as written, its zeros included.
            filled.update({index: items[index].text for index in below})
        else:
            foot = items[below[0]].text
            filled.update(dict.fromkeys(below, ""))
            filled[below[0]] = str(denominator).ljust(len(below)) if foot == "?" else str(denominator)
    text = "".join(filled.get(index, _literal(item)) for index, item in enumerate(items))
    # A mixed fraction keeps the minus of a number that rounds to nothing, -0.01 showing -0; an improper one does not.
    return text, not whole_places and not numerator


# --- dates and times ------------------------------------------------------------------------------

_MONTHS = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
           "November", "December")
_DAYS = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")


def excel_date(day: int) -> tuple[int, int, int]:
    """A serial day on Excel's 1900 calendar, which counts a 29 February 1900 and a day 0 of January."""
    if day == 60:
        return 1900, 2, 29
    if day == 0:
        return 1900, 1, 0
    when = _dt.date(1899, 12, 30) + _dt.timedelta(days=day + 1 if day < 60 else day)
    return when.year, when.month, when.day


def clock(value: float, places: int = 0) -> tuple[int, int, int, int, int]:
    """A serial as the day, hour, minute, second and ``places`` decimals of a second Excel shows for it.

    Excel adds half of the last unit shown to the serial, in double
    arithmetic, and then takes the day, the hours, the minutes and the
    seconds off what that leaves, each by multiplying up and cutting off.
    So a time a hair under half a second can still round up, and one a
    hair over can round down; measured on 152 such times.
    """
    unit = 10**places
    moment = value + 0.5 / (86400 * unit)
    day = math.floor(moment)
    rest = (moment - day) * 24
    hour = math.floor(rest)
    rest = (rest - hour) * 60
    minute = math.floor(rest)
    rest = (rest - minute) * 60
    second = math.floor(rest)
    return day, hour, minute, second, math.floor((rest - second) * unit)


def _date(items: list[_Item], value: float) -> str:
    """A serial through date and time codes, rounded to the smallest part shown."""
    parts = [item for item in items if item.kind in ("date", "elapsed", "ampm")]
    if len(parts) == 1 and parts[0].kind == "elapsed" and parts[0].text[0] == "s" and not any(
            item.kind == "point" for item in items):
        # Elapsed seconds alone are a count of seconds, sign and all, with no calendar to leave: the
        # seconds to fifteen digits, then to the nearest one.
        seconds = int(math.copysign(int(_rounded(_significant(abs(value) * 86400), 0)), value))
        return "".join(("-" if seconds < 0 else "") + str(abs(seconds)).rjust(len(item.text), "0")
                       if item.kind == "elapsed" else _literal(item) for item in items)
    if value < 0 or value >= 2958466:
        raise UndisplayableError("a date before 1900 or past 9999")
    # Places of a second: 0s straight after a point that follows the seconds.
    tenths: set[int] = set()
    for index, item in enumerate(items):
        if item.kind == "point" and index and items[index - 1].kind in ("date", "elapsed") \
                and items[index - 1].text[0] == "s":
            look = index + 1
            while look < len(items) and items[look].kind == "digit":
                tenths.add(look)
                look += 1
    places = len(tenths)
    day, hour, minute, second, part = clock(value, places)
    if day >= 2958466:
        raise UndisplayableError("a date past 9999")
    total = day * 86400 + hour * 3600 + minute * 60 + second
    year, month, date = excel_date(day)
    twelve = any(item.kind == "ampm" for item in items)
    fraction = str(part).rjust(places, "0") if places else ""
    out: list[str] = []
    for index, item in enumerate(items):
        if index in tenths:
            out.append(fraction[sorted(tenths).index(index)])
        elif item.kind == "date":
            out.append(_part(item.text, items, index, (year, month, date, day), (hour, minute, second), twelve))
        elif item.kind == "elapsed":
            amount = {"h": total // 3600, "m": total // 60, "s": total}[item.text[0]]
            out.append(str(amount).rjust(len(item.text), "0"))
        elif item.kind == "ampm":
            morning = hour < 12
            out.append(("AM" if morning else "PM") if len(item.text) == 5 else item.text[0 if morning else 2])
        else:
            out.append(_literal(item))
    return "".join(out)


def _part(code: str, items: list[_Item], index: int, calendar: tuple[int, int, int, int],
          clock: tuple[int, int, int], twelve: bool) -> str:
    year, month, date, day = calendar
    hour, minute, second = clock
    letter, count = code[0], len(code)
    if letter == "y":
        return f"{year % 100:02d}" if count <= 2 else f"{year:04d}"
    if letter == "d":
        if count <= 2:
            return f"{date:0{count}d}"
        name = _DAYS[(day + 6) % 7]
        return name[:3] if count == 3 else name
    if letter == "h":
        shown = (hour % 12 or 12) if twelve else hour
        return f"{shown:0{min(count, 2)}d}"
    if letter == "s":
        return f"{second:0{min(count, 2)}d}"
    # m is minutes straight after an hour or before seconds, and a month otherwise.
    if count <= 2 and (_neighbour(items, index, -1) == "h" or _neighbour(items, index, 1) == "s"):
        return f"{minute:0{count}d}"
    if count <= 2:
        return f"{month:0{count}d}"
    name = _MONTHS[month - 1]
    return name[:3] if count == 3 else name if count == 4 else name[0]


def _neighbour(items: list[_Item], index: int, step: int) -> str:
    """The letter of the nearest date or time part before (-1) or after (+1) ``index``."""
    look = index + step
    while 0 <= look < len(items):
        if items[look].kind in ("date", "elapsed"):
            return items[look].text[0]
        look += step
    return ""
