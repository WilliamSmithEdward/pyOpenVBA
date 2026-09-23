"""Range.AutoFill, as Excel's object model answers it.

Measured in live Excel (scripts/measure_autofill.py):

- The destination is one block on the source's sheet that holds the
  source and runs on from one of its edges, down, up, right or left;
  anything else is error 1004, a missing destination error 449. One that
  runs on in two directions at once fills in a way the model does not
  follow and reports itself.
- The source repeats along the fill, one line at a time: each column of
  it for a fill down or up, each row for a fill across. Inside a line,
  neighbouring cells of one kind make a run, blanks between them skipped
  and anything else ending it, and each run carries on as a series while
  the rest repeats: formulas moved as a copy moves them, text, logical
  values, errors and blanks as they are. A cell takes the format of the
  source cell it repeats.
- Numbers run together while their formats agree: General goes with any
  number format that is not a date's or a time's, and otherwise the
  first format that is not General decides. One number repeats when it
  is the whole source and steps by 1 otherwise; two step by their
  difference; more follow their least-squares line. Each value is the
  first one plus the step times how far along it is, that product
  rounded to 15 significant digits and then the sum, half away from
  zero: so 1/3, 2/3 carries on 1, 1.33333333333333, 1.66666666666666.
- A lone cell repeats rather than steps when the cells beside it across
  the fill make a run of their own: a row of numbers filled down repeats.
- A date steps by a day, under the built-in format mmm-yy by a month. Two
  or more step by their difference, by whole months where they share a
  day of the month or all end one, and repeat where the steps differ. A
  time steps by an hour.
- Text ending in a number steps that number, keeping its leading zeros
  and whatever comes before it; text starting with one followed by a
  space steps that; 1st, 2nd and the rest step as ordinals. A run is a
  shared prefix, ignoring case and a last space, dot or hyphen, and its
  steps must be even. A value below 0 is shown without its sign, and one
  past 4294967295 wraps. Q1 to Q4, Qtr and Quarter forms, and 1st Qtr to
  4th Qtr wrap round the quarters.
- Day and month names, short or long and cased as the first is, step
  round the week and the year.
- xlFillCopy repeats everything, xlFillFormats only the formats and
  xlFillValues fills without them. xlFillDays, xlFillWeekdays,
  xlFillMonths and xlFillYears step dates by that unit and drop their
  times; everything else repeats under them. xlLinearTrend steps
  numbers and repeats the rest, and xlGrowthTrend multiplies.

What is not modelled reports itself: a trend through three or more
numbers that do not step evenly (Excel takes it from its LINEST
arithmetic, whose last digit the model cannot match), a growth trend,
dates whose times differ, weekday fills of several dates wider than a
week apart, Flash Fill, merged cells, and row or column formats.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from collections.abc import Callable
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel import _merges
from pyopenvba.apps.excel._typing import date_time_codes, is_date_format
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._display import excel_date
from pyopenvba.formula._parse import shift_text
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import EMPTY, MISSING, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Range, Worksheet
    from pyopenvba.apps.excel._styles import Style

DEFAULT, COPY, SERIES, FORMATS, VALUES, DAYS, WEEKDAYS, MONTHS, YEARS, LINEAR, GROWTH, FLASH = range(12)
_DATE_UNITS = {DAYS, WEEKDAYS, MONTHS, YEARS}
#: The fill types where a lone cell repeats beside a run across the fill.
_GUESSING = {DEFAULT, VALUES}

_DAYS = [["mon", "tue", "wed", "thu", "fri", "sat", "sun"],
         ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]]
_MONTHS = [["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"],
           ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
            "november", "december"]]
_QUARTERS = {"q", "qtr", "quarter"}
_WRAP = 2**32
_EPOCH = _dt.date(1899, 12, 30)


def autofill(source: Range, destination: object, kind_argument: object) -> object:
    from pyopenvba.apps.excel._model import Range

    if destination is MISSING:
        raise error(449)
    if not isinstance(destination, Range):
        raise error(1004, "AutoFill needs a range to fill")
    kind = DEFAULT if kind_argument is MISSING else int(to_integer(kind_argument, "Long"))
    if kind == FLASH:
        raise VBAUnsupportedError("AutoFill with xlFlashFill is not implemented")
    if not DEFAULT <= kind <= GROWTH:
        raise error(1004, "AutoFill has no such fill type")
    if len(source.areas) != 1 or len(destination.areas) != 1 or destination.sheet is not source.sheet:
        raise error(1004, "AutoFill fills one block from one block on the same sheet")
    area, target = source.first, destination.first
    down = _direction(area, target)
    sheet = source.sheet
    from pyopenvba.apps.excel._protection import whole

    if down:
        filled = (Area(area.bottom + 1, area.left, target.bottom, area.right) if target.bottom > area.bottom
                  else Area(target.top, area.left, area.top - 1, area.right))
    else:
        filled = (Area(area.top, area.right + 1, area.bottom, target.right) if target.right > area.right
                  else Area(area.top, target.left, area.bottom, area.left - 1))
    whole(sheet, [filled], "AutoFill")
    _check_modelled(sheet, area, target)
    plan = _Plan(sheet, area, target, down, kind)
    plan.fill()
    return True


def _direction(area: Area, target: Area) -> bool:
    """True for a fill down or up, False for one across; error 1004 when the blocks do not line up."""
    same_columns = target.left == area.left and target.right == area.right
    same_rows = target.top == area.top and target.bottom == area.bottom
    if same_columns and (target.top == area.top and target.bottom > area.bottom
                         or target.bottom == area.bottom and target.top < area.top):
        return True
    if same_rows and (target.left == area.left and target.right > area.right
                      or target.right == area.right and target.left < area.left):
        return False
    inside = target.top <= area.top and area.bottom <= target.bottom and target.left <= area.left \
        and area.right <= target.right
    if inside and target.rows > area.rows and target.columns > area.columns:
        raise VBAUnsupportedError("AutoFill into a destination that runs on in two directions is not implemented")
    raise error(1004, "The destination has to hold the source and run on from one of its edges")


def _check_modelled(sheet: Worksheet, area: Area, target: Area) -> None:
    if target.rows * target.columns > 1048576:
        raise VBAUnsupportedError("AutoFill into more than 1,048,576 cells is not implemented")
    if any(_merges.intersects(target, one) for one in sheet.merged_areas):
        raise VBAUnsupportedError("AutoFill involving merged cells is not implemented")
    dims = sheet.dims
    if any(target.top <= row <= target.bottom and record.style is not None for row, record in dims.rows.items()) \
            or any(dims.column_style(column) is not None for column in range(target.left, target.right + 1)):
        raise VBAUnsupportedError("AutoFill on rows or columns with formats of their own is not implemented")


# --- what a source cell holds ----------------------------------------------------------------------


@dataclass
class _Item:
    """One source cell as AutoFill reads it."""

    kind: str
    #: A number, date or time; a text's number; a name's place in its list.
    value: float = 0.0
    #: What a neighbour has to share to carry on the same run.
    key: tuple[str, ...] = ()
    #: A number's format family: general, number, percent, currency, scientific, fraction or text.
    family: str = ""
    #: A trailing number's text before it; a name's leading spaces.
    prefix: str = ""
    #: A leading number's or an ordinal's text after it.
    rest: str = ""
    #: A trailing number's width, when it is written with leading zeros.
    width: int = 0
    quarter: bool = False
    #: A name's list, short or long, empty for May; and its case.
    names: str = ""
    case: str = ""
    number_format: str = ""


def _read(cell: Cell | None, number_format: str) -> _Item:
    if cell is None or (cell.value is EMPTY and not cell.formula):
        return _Item("blank")
    if cell.formula:
        return _Item("copy")
    value = cell.value
    if isinstance(value, (bool, ExcelError)):
        return _Item("copy")
    if isinstance(value, (int, float)):
        number = float(value)
        if is_date_format(number_format):
            return _Item("date", number, ("date",), number_format=number_format)
        if date_time_codes(number_format):
            return _Item("time", number, ("time",), number_format=number_format)
        return _Item("number", number, ("number",), family=_family(number_format), number_format=number_format)
    if isinstance(value, str):
        return _text(value)
    return _Item("copy")


def _family(code: str) -> str:
    """The format family a number keeps a run with: General goes with any of the others."""
    if code in ("General", ""):
        return "general"
    body = re.sub(r'"[^"]*"|\\.|_.|\*.', "", code)
    if "@" in body:
        return "text"
    if "%" in re.sub(r"\[[^\]]*\]", "", body):
        return "percent"
    if re.search(r"[eE][+-]", body):
        return "scientific"
    if "/" in body:
        return "fraction"
    if "$" in code or "[$" in code or any(sign in code for sign in "€£¥"):
        return "currency"
    return "number"


_LEADING = re.compile(r"(\d+)( +)(.*)\Z", re.DOTALL)
_ORDINAL = re.compile(r"(\d+)(st|nd|rd|th)((?: .*)?)\Z", re.IGNORECASE | re.DOTALL)
_TRAILING = re.compile(r"(\d+)\Z")


def _text(text: str) -> _Item:
    trimmed = text.rstrip(" ")
    name = _name(text, trimmed)
    if name is not None:
        return name
    found = _LEADING.match(text)
    if found:
        if not found.group(3).strip(" "):
            return _Item("copy")
        number = int(found.group(1))
        if number >= _WRAP:
            return _Item("copy")
        rest = found.group(2) + found.group(3)
        tail = rest.strip(" ").lower()
        return _Item("leading", number, ("leading", tail), rest=rest, quarter=tail in _QUARTERS and 1 <= number <= 4)
    found = _ORDINAL.match(text)
    if found and found.group(2).lower() == _suffix(int(found.group(1))) and int(found.group(1)) < _WRAP:
        number = int(found.group(1))
        tail = found.group(3).strip(" ").lower()
        return _Item("ordinal", number, ("ordinal", tail), rest=found.group(3),
                     quarter=tail in _QUARTERS and 1 <= number <= 4)
    found = _TRAILING.search(trimmed)
    if found:
        digits = found.group(1)
        number = int(digits)
        if number >= _WRAP:
            return _Item("copy")
        prefix = trimmed[: found.start()]
        stem = _stem(prefix)
        width = len(digits) if len(digits) > 1 and digits[0] == "0" else 0
        return _Item("trailing", number, ("trailing", stem), prefix=prefix, width=width,
                     quarter=stem in _QUARTERS and len(digits) == 1 and 1 <= number <= 4)
    return _Item("copy")


def _stem(prefix: str) -> str:
    """A trailing number's prefix as runs compare it: no last spaces, then no last dot or hyphen, any case."""
    stem = prefix.rstrip(" ")
    if stem[-1:] in (".", "-"):
        stem = stem[:-1]
    return stem.lower()


def _name(text: str, trimmed: str) -> _Item | None:
    core = trimmed[:-1] if trimmed.endswith(".") else trimmed
    body = core.lstrip(" ")
    lowered = body.lower()
    lead = text[: len(text) - len(text.lstrip(" "))]
    case = "lower" if body[:1].islower() else "upper" if body[1:2].isupper() else "title"
    for kind, lists in (("day", _DAYS), ("month", _MONTHS)):
        for which, names in zip(("short", "long"), lists, strict=True):
            if lowered in names:
                listed = "" if kind == "month" and lowered == "may" else which
                return _Item(kind, float(names.index(lowered)), (kind,), prefix=lead, names=listed, case=case)
    return None


def _suffix(number: int) -> str:
    if 11 <= number % 100 <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(number % 10, "th")


# --- runs --------------------------------------------------------------------------------------------


def _runs(items: list[_Item]) -> list[list[int]]:
    """The runs in a line: indexes of neighbouring items of one kind, blanks skipped, anything else ending one."""
    runs: list[list[int]] = []
    current: list[int] = []
    family = names = ""
    for index, item in enumerate(items):
        if item.kind == "blank":
            continue
        if item.kind == "copy":
            current = []
            continue
        joins = bool(current) and items[current[0]].key == item.key
        if joins and item.kind == "number":
            joins = item.family == "general" or family in ("general", item.family)
        if joins and item.kind in ("day", "month"):
            joins = not item.names or not names or item.names == names
        if not joins:
            current = [index]
            runs.append(current)
            family, names = item.family, item.names
            continue
        current.append(index)
        if family == "general":
            family = item.family
        names = names or item.names
    return runs


# --- the fill ----------------------------------------------------------------------------------------


class _Plan:
    def __init__(self, sheet: Worksheet, area: Area, target: Area, down: bool, kind: int) -> None:
        self.sheet, self.area, self.target, self.down, self.kind = sheet, area, target, down, kind
        self.length = area.rows if down else area.columns
        self.single = area.rows == 1 and area.columns == 1
        self.cells = {position: cell for position, cell in sheet.cells_.items() if area.contains(*position)}

    def position(self, line: int, index: int) -> tuple[int, int]:
        return (self.area.top + index, line) if self.down else (line, self.area.left + index)

    def item(self, row: int, column: int) -> _Item:
        return _read(self.cells.get((row, column)), self.sheet.style_at(row, column).number_format)

    def crossed(self) -> set[tuple[int, int]]:
        """The source cells that make a run with their neighbours across the fill."""
        if self.kind not in _GUESSING:
            return set()
        marked: set[tuple[int, int]] = set()
        lines = range(self.area.left, self.area.right + 1) if self.down else range(self.area.top, self.area.bottom + 1)
        for index in range(self.length):
            positions = [self.position(line, index) for line in lines]
            items = [self.item(*position) for position in positions]
            for run in _runs(items):
                if len(run) > 1:
                    marked.update(positions[member] for member in run)
        return marked

    def fill(self) -> None:
        crossed = self.crossed()
        lines = range(self.area.left, self.area.right + 1) if self.down else range(self.area.top, self.area.bottom + 1)
        first, last = (self.target.top, self.target.bottom) if self.down else (self.target.left, self.target.right)
        start = self.area.top if self.down else self.area.left
        writes: list[tuple[tuple[int, int], tuple[int, int], object]] = []
        for line in lines:
            positions = [self.position(line, index) for index in range(self.length)]
            items = [self.item(*position) for position in positions]
            makers: dict[int, tuple[int, int, Callable[[int], object]]] = {}
            for run in _runs(items):
                series = _series([items[index] for index in run], self.kind, self.single,
                                 len(run) == 1 and positions[run[0]] in crossed)
                if series is not None:
                    for place, index in enumerate(run):
                        makers[index] = (place, len(run), series)
            for spot in range(first, last + 1):
                index = spot - start
                if 0 <= index < self.length:
                    continue
                slot = index % self.length
                repeat = index // self.length
                made = makers.get(slot)
                content: object = _COPIED
                if made is not None:
                    place, size, series = made
                    content = series(place + repeat * size)
                at = (spot, line) if self.down else (line, spot)
                writes.append((at, positions[slot], content))
        for at, origin, content in writes:
            self.write(at, origin, content)

    def write(self, at: tuple[int, int], origin: tuple[int, int], content: object) -> None:
        from pyopenvba.apps.excel._model import Cell

        sheet = self.sheet
        source = self.cells.get(origin)
        default = sheet.book.stylesheet.default
        style: Style = sheet.style_at(*at) if self.kind == VALUES else sheet.style_at(*origin)
        if self.kind == FORMATS:
            sheet.restyle(*at, style)
            return
        formula, value = "", EMPTY
        if content is _COPIED:
            if source is not None and source.formula:
                formula = shift_text(source.formula, at[0] - origin[0], at[1] - origin[1])
            elif source is not None:
                value = source.value
        else:
            value = content
        cell = Cell(value=value, formula=formula, stale=bool(formula), style=None if style == default else style)
        sheet.cells_[at] = cell
        sheet.settle(*at)
        sheet.cell_changed(*at)


class _Copied:
    """What a slot gets when its source cell repeats as it is."""


_COPIED = _Copied()


def _series(run: list[_Item], kind: int, single: bool, crossed: bool) -> Callable[[int], object] | None:
    """How a run carries on: its value at each step along it, or None where it repeats."""
    if kind in (COPY, FORMATS):
        return None
    first = run[0]
    lone = len(run) == 1
    if lone and crossed:
        return None
    if first.kind == "number":
        return _numbers(run, kind, single)
    if first.kind == "time":
        return _times(run, kind)
    if first.kind == "date":
        return _dates(run, kind)
    if first.kind in ("day", "month"):
        return _names(run, kind)
    return _texts(run, kind)


def _numbers(run: list[_Item], kind: int, single: bool) -> Callable[[int], object] | None:
    values = [item.value for item in run]
    if kind in _DATE_UNITS:
        return None
    if kind == GROWTH:
        if len(values) == 1 or len(set(values)) == 1:
            return None
        if any(value <= 0 for value in values):
            return lambda step: 0.0
        raise VBAUnsupportedError("AutoFill's growth trend is not implemented")
    if len(values) == 1:
        if single and kind in _GUESSING:
            return None
        return _line(values[0], 1.0)
    return _even(values)


def _even(values: list[float]) -> Callable[[int], object]:
    """Two numbers, or more that step evenly to 15 digits: the first plus the first difference."""
    step = values[1] - values[0]
    if len({_r15(later - earlier) for earlier, later in zip(values, values[1:])}) > 1:
        raise VBAUnsupportedError("AutoFill's trend through three or more numbers that do not step evenly is not "
                                  "implemented: Excel takes it from its LINEST arithmetic")
    return _line(values[0], step)


def _line(base: float, step: float) -> Callable[[int], object]:
    def value(along: int) -> object:
        found = _r15(base + _r15(along * step))
        if not math.isfinite(found):
            raise VBAUnsupportedError("AutoFill past the largest number is not implemented")
        return found
    return value


def _r15(value: float) -> float:
    """A number rounded to 15 significant digits, half away from zero."""
    if value == 0 or not math.isfinite(value):
        return value
    exact = Decimal(value)
    return float(exact.quantize(Decimal(1).scaleb(exact.adjusted() - 14), rounding=ROUND_HALF_UP))


def _times(run: list[_Item], kind: int) -> Callable[[int], object] | None:
    values = [item.value for item in run]
    if kind in _DATE_UNITS:
        return None
    if kind in (LINEAR, GROWTH):
        raise VBAUnsupportedError("AutoFill of times with a trend is not implemented")
    if len(values) == 1:
        return _line(values[0], 1 / 24)
    if any(not 0 <= value < 1 for value in values):
        raise VBAUnsupportedError("AutoFill of several times a day or more apart is not implemented")
    return _even(values)


def _dates(run: list[_Item], kind: int) -> Callable[[int], object] | None:
    if kind in (LINEAR, GROWTH):
        return None
    values = [item.value for item in run]
    days = [math.floor(value) for value in values]
    fractions = {value - day for value, day in zip(values, days, strict=True)}
    if kind in _DATE_UNITS:
        keep = 0.0
    elif len(fractions) > 1:
        raise VBAUnsupportedError("AutoFill of dates whose times differ is not implemented")
    else:
        keep = fractions.pop()
    if len(days) == 1:
        day = days[0]
        if kind == WEEKDAYS:
            return lambda along: _weekdays(day, along) + keep
        if kind == MONTHS or (kind != DAYS and run[0].number_format == "mmm-yy"):
            return lambda along: _months(day, along, end=False) + keep
        if kind == YEARS:
            return lambda along: _months(day, 12 * along, end=False) + keep
        return lambda along: day + along + keep
    steps = {later - earlier for earlier, later in zip(days, days[1:])}
    monthly = _monthly(days)
    count = len(days)
    if kind in (DEFAULT, VALUES, SERIES, MONTHS, YEARS):
        if monthly is not None and (kind != YEARS or monthly[1] % 12 == 0):
            base, months, end = monthly
            return lambda along: _months(base, along * months, end=end) + keep
    if len(steps) > 1:
        return None
    step = steps.pop()
    if kind in (DEFAULT, VALUES, SERIES, DAYS):
        return lambda along: days[0] + along * step + keep
    if kind in (MONTHS, YEARS):
        # Each time round, dates on one day of the month move a month or a year each; others keep
        # their distance in days from the first, which moves.
        unit = 1 if kind == MONTHS else 12
        same, ends = _aligned(days)
        if same or ends:
            return lambda along: _months(days[along % count], along // count * unit, end=ends and not same)
        return lambda along: _months(days[0], along // count * unit, end=False) + days[along % count] - days[0]
    if abs(step) == 1:
        return lambda along: _weekdays(days[0], along * step)
    if 1 < abs(step) <= 7:
        return lambda along: _weekdays(days[along % count], along // count * 5)
    raise VBAUnsupportedError("AutoFill of weekdays from dates more than a week apart is not implemented")


def _monthly(days: list[int]) -> tuple[int, int, bool] | None:
    """Dates a whole number of months apart: the first, the months between each, and whether all end a month."""
    if min(days) < 61:
        return None
    parts = [_calendar(day) for day in days]
    counts = [year * 12 + month for year, month, _ in parts]
    steps = {later - earlier for earlier, later in zip(counts, counts[1:])}
    if len(steps) != 1:
        return None
    same, ends = _aligned(days)
    if ends and same and parts[0][2] < 31:
        raise VBAUnsupportedError("AutoFill of month ends that share a day of the month is not implemented")
    if not (ends or same):
        return None
    return days[0], steps.pop(), ends


def _aligned(days: list[int]) -> tuple[bool, bool]:
    """Whether dates share a day of the month, and whether they all end one."""
    if min(days) < 61:
        return False, False
    parts = [_calendar(day) for day in days]
    return len({date for _, _, date in parts}) == 1, all(date == _month_length(year, month) for year, month, date in parts)


def _calendar(day: int) -> tuple[int, int, int]:
    if day < 61:
        raise VBAUnsupportedError("AutoFill by months or weekdays before 1 March 1900 is not implemented")
    return excel_date(day)


def _month_length(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (_dt.date(year, month + 1, 1) - _dt.date(year, month, 1)).days


def _months(day: int, months: int, *, end: bool) -> float:
    """The date ``months`` on from a serial day, on its day of the month or the month's last one."""
    year, month, date = _calendar(day)
    total = year * 12 + month - 1 + months
    year, month = divmod(total, 12)
    month += 1
    if not 1900 <= year <= 9999:
        raise VBAUnsupportedError("AutoFill past the calendar's first or last year is not implemented")
    length = _month_length(year, month)
    return float((_dt.date(year, month, length if end else min(date, length)) - _EPOCH).days)


def _weekdays(day: int, count: int) -> float:
    """The weekday ``count`` weekdays on from a serial day; a weekend day counts from the Monday or Friday."""
    _calendar(day)
    weekday = (_EPOCH + _dt.timedelta(days=day)).weekday()
    if count == 0:
        return float(day)
    if count > 0:
        if weekday >= 5:
            day += 7 - weekday
            count -= 1
            weekday = 0
        weeks, rest = divmod(count, 5)
        day += weeks * 7 + rest + (2 if weekday + rest >= 5 else 0)
        return float(day)
    if weekday >= 5:
        day -= weekday - 4
        count += 1
        weekday = 4
    weeks, rest = divmod(-count, 5)
    day -= weeks * 7 + rest + (2 if weekday - rest < 0 else 0)
    return float(day)


def _names(run: list[_Item], kind: int) -> Callable[[int], object] | None:
    first = run[0]
    days = first.kind == "day"
    if kind not in ((DEFAULT, VALUES, SERIES, DAYS, WEEKDAYS) if days else (DEFAULT, VALUES, SERIES, MONTHS)):
        return None
    size = 7 if days else 12
    listed = next((item.names for item in run if item.names), "long")
    names = (_DAYS if days else _MONTHS)[0 if listed == "short" else 1]
    places = [int(item.value) for item in run]
    if kind == WEEKDAYS:
        if len(run) > 1 or places[0] >= 5:
            raise VBAUnsupportedError("AutoFill of weekend day names or several day names by weekdays is not "
                                      "implemented")
        # Monday to Friday, round and round.
        return lambda along: _cased(first, names[(places[0] + along) % 5])
    steps = {(later - earlier) % size for earlier, later in zip(places, places[1:])}
    if len(steps) > 1:
        return None
    step = steps.pop() if steps else 1
    return lambda along: _cased(first, names[(places[0] + along * step) % size])


def _cased(first: _Item, name: str) -> str:
    shown = name.lower() if first.case == "lower" else name.upper() if first.case == "upper" else name.capitalize()
    return first.prefix + shown


def _texts(run: list[_Item], kind: int) -> Callable[[int], object] | None:
    if kind not in (DEFAULT, VALUES, SERIES):
        return None
    first = run[0]
    numbers = [int(item.value) for item in run]
    quarters = all(item.quarter for item in run)
    if quarters:
        steps = {(later - earlier) % 4 for earlier, later in zip(numbers, numbers[1:])}
    else:
        steps = {later - earlier for earlier, later in zip(numbers, numbers[1:])}
    if len(steps) > 1:
        return None
    step = steps.pop() if steps else 1
    width = max(item.width for item in run)

    def text(along: int) -> object:
        number = (numbers[0] - 1 + along * step) % 4 + 1 if quarters else abs(numbers[0] + along * step) % _WRAP
        if first.kind == "trailing":
            return first.prefix + str(number).zfill(width)
        if first.kind == "leading":
            return str(number) + first.rest
        return f"{number}{_suffix(number)}{first.rest}"
    return text
