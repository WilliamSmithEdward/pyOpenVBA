"""The M standard library.

The functions a query actually calls: the table transforms a step
records, the list and text families they lean on, and the sources that
can be read without leaving the machine.

A name this does not implement is told apart from a name M has never
had.  ``Table.Pivot`` is real, so asking for it raises
:class:`~pyopenvba.exceptions.VBAUnsupportedError` naming it; ``Table.Pivvot``
is not, so it is M's own "the name wasn't recognized".  The split is by
namespace, since the M library is a closed set of them.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from typing import Any, Callable, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.mlang._values import (
    is_list,
    is_mapping,
    Builtin,
    Duration,
    MError,
    MType,
    Record,
    Table,
    as_list,
    as_logical,
    as_number,
    as_record,
    as_table,
    as_text,
    Scaled,
    duration_text,
    number_out,
    places_of,
    type_name,
)

LIBRARY: dict[str, object] = {}

#: The namespaces the M library is made of.  A name in one of these
#: that is not implemented is a gap here; a name outside them is a name
#: M has never had.
NAMESPACES: Final = frozenset(
    """
    Access Binary BinaryFormat Byte Character Combiner Comparer Csv Cube Currency Date DateTime
    DateTimeZone Decimal Diagnostics Double Duration Embedded Error Excel Expression Extension
    File Folder Function Guid Html Int16 Int32 Int64 Int8 ItemExpression Json Lines List Logical
    Module Number Occurrence OleDb Odbc Order Percentage Record RelativePosition Replacer
    RowExpression Section Single Splitter Sql Table TableAction Text Time Type Uri Value Variable
    Web Xml
    """.split()
)


def m(name: str, minimum: int = 0, maximum: int | None = None) -> Callable[..., Any]:
    """Register a library function under its M name."""

    def register(function: Callable[..., object]) -> Callable[..., object]:
        LIBRARY[name] = Builtin(name=name, call=function, minimum=minimum, maximum=maximum)
        return function

    return register


def missing(name: str) -> Exception:
    """The right complaint for a name the library has not got."""
    from pyopenvba.mlang._inventory import has_name

    if has_name(name):
        return VBAUnsupportedError(
            f"{name} is a real Power Query function that pyOpenVBA does not implement"
        )
    return MError("Expression.Error", f"The name '{name}' wasn't recognized.")


def _call(function: object, args: list[object]) -> object:
    from pyopenvba.mlang._eval import call_value

    return call_value(function, args)


def _items(value: object) -> list[Any]:
    """An argument that may be one value or a list of them, as a list."""
    return as_list(value) if is_list(value) else [value]


def _optional(args: tuple[object, ...], at: int, default: object = None) -> object:
    if at >= len(args) or args[at] is None:
        return default
    return args[at]


# --- the constructors written with a hash ------------------------------------------------


@m("#table", 0, 2)
def m_hash_table(*args: object) -> object:
    """``#table(columns, rows)``: the literal table."""
    columns = _optional(args, 0)
    rows = _optional(args, 1, [])
    names: list[str] = []
    if is_list(columns):
        names = [as_text(one) for one in as_list(columns)]
    elif isinstance(columns, MType) and is_list(columns.detail):
        names = [as_text(one) for one in as_list(columns.detail)]
    elif isinstance(columns, (int, float)):
        names = [f"Column{index + 1}" for index in range(int(as_number(columns)))]
    out: list[list[object]] = []
    for row in as_list(rows):
        out.append(list(as_list(row)))
    if not names and out:
        names = [f"Column{index + 1}" for index in range(len(out[0]))]
    return Table(names, out)


@m("#date", 3, 3)
def m_hash_date(year: object, month: object, day: object) -> object:
    try:
        return _dt.date(int(as_number(year)), int(as_number(month)), int(as_number(day)))
    except ValueError as failure:
        raise MError("Expression.Error", str(failure)) from None


@m("#time", 3, 3)
def m_hash_time(hour: object, minute: object, second: object) -> object:
    whole = int(as_number(second))
    return _dt.time(int(as_number(hour)), int(as_number(minute)), whole)


@m("#datetime", 6, 6)
def m_hash_datetime(*args: object) -> object:
    numbers = [int(as_number(one)) for one in args]
    return _dt.datetime(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5])


@m("#datetimezone", 6, 8)
def m_hash_datetimezone(*args: object) -> object:
    numbers = [int(as_number(one)) for one in args[:6]]
    return _dt.datetime(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5])


@m("#duration", 4, 4)
def m_hash_duration(days: object, hours: object, minutes: object, seconds: object) -> object:
    return Duration(as_number(days), as_number(hours), as_number(minutes), as_number(seconds))


@m("#range", 2, 2)
def m_hash_range(first: object, last: object) -> object:
    """{1..3}, which M writes as a range inside a list.

    The ends can be characters as well as numbers: ``{"0".."9"}`` is the
    ten digits, which is how ``Text.Select`` is usually given what to
    keep.  A range that runs backwards is empty.
    """
    if isinstance(first, str) and isinstance(last, str):
        return [chr(one) for one in range(ord(first), ord(last) + 1)]
    start = int(as_number(first))
    stop = int(as_number(last))
    return [number_out(float(one)) for one in range(start, stop + 1)]


@m("#binary", 1, 1)
def m_hash_binary(value: object) -> object:
    """#binary takes base64 or a list of bytes, not plain text."""
    import base64
    import binascii

    if isinstance(value, bytes):
        return value
    if is_list(value):
        return bytes(int(as_number(one)) & 0xFF for one in value)
    try:
        return base64.b64decode(as_text(value), validate=True)
    except (binascii.Error, ValueError):
        raise MError("Expression.Error", "We cannot read that as binary.") from None


# --- types --------------------------------------------------------------------------------

for _name in (
    "Int8",
    "Int16",
    "Int32",
    "Int64",
    "Single",
    "Double",
    "Decimal",
    "Currency",
    "Percentage",
    "Byte",
    "Guid",
):
    LIBRARY[f"{_name}.Type"] = MType(name="number" if _name != "Guid" else "text")

LIBRARY["Text.Type"] = MType(name="text")
LIBRARY["Number.Type"] = MType(name="number")
LIBRARY["Logical.Type"] = MType(name="logical")
LIBRARY["Date.Type"] = MType(name="date")
LIBRARY["DateTime.Type"] = MType(name="datetime")
LIBRARY["Time.Type"] = MType(name="time")
LIBRARY["Duration.Type"] = MType(name="duration")
LIBRARY["Any.Type"] = MType(name="any")

#: The join kinds, as records of their own name.
for _kind in ("Inner", "LeftOuter", "RightOuter", "FullOuter", "LeftAnti", "RightAnti"):
    LIBRARY[f"JoinKind.{_kind}"] = _kind
for _kind in ("Global", "Local"):
    LIBRARY[f"GroupKind.{_kind}"] = _kind
LIBRARY["Order.Ascending"] = 0
LIBRARY["Order.Descending"] = 1
LIBRARY["Occurrence.First"] = 0
LIBRARY["Occurrence.Last"] = 1
LIBRARY["Occurrence.All"] = 2
LIBRARY["ExtraValues.List"] = 2
LIBRARY["ExtraValues.Error"] = 0
LIBRARY["ExtraValues.Ignore"] = 1
LIBRARY["MissingField.Error"] = 0
LIBRARY["MissingField.Ignore"] = 1
LIBRARY["MissingField.UseNull"] = 2
for _mode in ("ToEven", "AwayFromZero", "TowardZero", "Up", "Down"):
    LIBRARY[f"RoundingMode.{_mode}"] = _mode
LIBRARY["BinaryEncoding.Base64"] = "Base64"
LIBRARY["BinaryEncoding.Hex"] = "Hex"
LIBRARY["TextEncoding.Utf8"] = 65001
LIBRARY["TextEncoding.Windows"] = 1252


# --- conversion ------------------------------------------------------------------------------


def _to_number(value: object) -> object:
    if value is None:
        return None
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, (int, float)):
        return number_out(float(value))
    if isinstance(value, str):
        body = value.strip()
        if not body:
            return None
        try:
            return number_out(float(body))
        except ValueError:
            raise MError(
                "DataFormat.Error", f"We couldn't convert to Number: {value}"
            ) from None
    raise MError("DataFormat.Error", f"We cannot convert a {type_name(value)} to Number.")


@m("Number.From", 1, 2)
def m_number_from(value: object, *rest: object) -> object:
    return _to_number(value)


@m("Text.From", 1, 2)
def m_text_from(value: object, *rest: object) -> object:
    if value is None:
        return None
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, _dt.datetime):
        return value.strftime("%#m/%#d/%Y %#I:%M:%S %p")
    if isinstance(value, _dt.date):
        return value.strftime("%#m/%#d/%Y")
    if isinstance(value, _dt.time):
        return value.strftime("%#I:%M %p")
    if isinstance(value, Duration):
        return _duration_text(value)
    return as_text(value)


def _duration_text(value: Duration) -> str:
    """A duration as M writes one: days, then hours, minutes, seconds."""
    total = round(value.total_seconds)
    days, rest = divmod(int(total), 86400)
    hours, rest = divmod(rest, 3600)
    minutes, seconds = divmod(rest, 60)
    return f"{days}.{hours:02d}:{minutes:02d}:{seconds:02d}"


@m("Number.ToText", 1, 3)
def m_number_to_text(value: object, *rest: object) -> object:
    if value is None:
        return None
    pattern = _optional(rest, 0)
    number = as_number(value)
    if pattern is None:
        return as_text(number_out(number))
    from pyopenvba.access._format import format_value

    return format_value(number, as_text(pattern))


@m("Logical.From", 1, 1)
def m_logical_from(value: object) -> object:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        body = value.strip().lower()
        if body in ("true", "false"):
            return body == "true"
        raise MError("Expression.Error", f"We couldn't convert to Logical: {value}")
    return as_number(value) != 0


@m("Date.From", 1, 2)
def m_date_from(value: object, *rest: object) -> object:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, str):
        return _parse_date(value).date()
    number = as_number(value)
    return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=number)).date()


@m("DateTime.From", 1, 2)
def m_datetime_from(value: object, *rest: object) -> object:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    if isinstance(value, str):
        return _parse_date(value)
    return _dt.datetime(1899, 12, 30) + _dt.timedelta(days=as_number(value))


@m("Time.From", 1, 2)
def m_time_from(value: object, *rest: object) -> object:
    if value is None:
        return None
    if isinstance(value, _dt.datetime):
        return value.time()
    if isinstance(value, _dt.time):
        return value
    if isinstance(value, str):
        return _parse_date(value).time()
    fraction = as_number(value) % 1
    return (_dt.datetime(2000, 1, 1) + _dt.timedelta(days=fraction)).time()


def _parse_date(text: str) -> _dt.datetime:
    body = text.strip()
    for pattern in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%H:%M:%S",
        "%H:%M",
    ):
        try:
            return _dt.datetime.strptime(body, pattern)
        except ValueError:
            continue
    raise MError("Expression.Error", f"We couldn't parse the date: {text}")


@m("DateTime.LocalNow", 0, 0)
def m_localnow() -> object:
    return _dt.datetime.now()


@m("DateTime.FixedLocalNow", 0, 0)
def m_fixed_localnow() -> object:
    return _dt.datetime.now()


@m("Date.From.Text", 1, 2)
def m_date_from_text(value: object, *rest: object) -> object:
    return _parse_date(as_text(value)).date()


# --- Date ---------------------------------------------------------------------------------------


def _as_date(value: object) -> _dt.date:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, str):
        return _parse_date(value).date()
    raise MError("Expression.Error", f"We cannot read a {type_name(value)} as a date.")


@m("Date.Year", 1, 1)
def m_date_year(value: object) -> object:
    return None if value is None else _as_date(value).year


@m("Date.Month", 1, 1)
def m_date_month(value: object) -> object:
    return None if value is None else _as_date(value).month


@m("Date.Day", 1, 1)
def m_date_day(value: object) -> object:
    return None if value is None else _as_date(value).day


@m("Date.DayOfWeek", 1, 2)
def m_date_day_of_week(value: object, *rest: object) -> object:
    if value is None:
        return None
    start = _optional(rest, 0, 0)
    offset = int(as_number(start)) if isinstance(start, (int, float)) else 0
    return (_as_date(value).weekday() + 1 - offset) % 7


@m("Date.AddDays", 2, 2)
def m_date_add_days(value: object, count: object) -> object:
    if value is None:
        return None
    moved = _as_date(value) + _dt.timedelta(days=int(as_number(count)))
    return _dt.datetime.combine(moved, value.time()) if isinstance(value, _dt.datetime) else moved


@m("Date.AddMonths", 2, 2)
def m_date_add_months(value: object, count: object) -> object:
    if value is None:
        return None
    when = _as_date(value)
    total = when.year * 12 + when.month - 1 + int(as_number(count))
    year, month = divmod(total, 12)
    day = min(when.day, _days_in_month(year, month + 1))
    return _dt.date(year, month + 1, day)


@m("Date.AddYears", 2, 2)
def m_date_add_years(value: object, count: object) -> object:
    return m_date_add_months(value, int(as_number(count)) * 12)


@m("Date.StartOfMonth", 1, 1)
def m_date_start_of_month(value: object) -> object:
    when = _as_date(value)
    return _dt.date(when.year, when.month, 1)


@m("Date.EndOfMonth", 1, 1)
def m_date_end_of_month(value: object) -> object:
    when = _as_date(value)
    return _dt.date(when.year, when.month, _days_in_month(when.year, when.month))


@m("Date.ToText", 1, 3)
def m_date_to_text(value: object, *rest: object) -> object:
    if value is None:
        return None
    pattern = _optional(rest, 0)
    when = _as_date(value)
    if pattern is None:
        return when.isoformat()
    from pyopenvba.access._format import format_value

    return format_value(_dt.datetime(when.year, when.month, when.day), as_text(pattern))


def _days_in_month(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (_dt.date(year, month + 1, 1) - _dt.date(year, month, 1)).days


# --- Number -----------------------------------------------------------------------------------


@m("Number.Round", 1, 3)
def m_number_round(value: object, *rest: object) -> object:
    """M rounds a half to even, where Excel's own ROUND rounds away."""
    if value is None:
        return None
    from decimal import ROUND_HALF_EVEN, Decimal

    from decimal import ROUND_HALF_DOWN, ROUND_HALF_UP

    digits = int(as_number(_optional(rest, 0, 0)))
    mode = _optional(rest, 1, "ToEven")
    how = {
        "ToEven": ROUND_HALF_EVEN,
        "AwayFromZero": ROUND_HALF_UP,
        "TowardZero": ROUND_HALF_DOWN,
        "Up": ROUND_HALF_UP,
        "Down": ROUND_HALF_DOWN,
    }.get(as_text(mode), ROUND_HALF_EVEN)
    quantum = Decimal(1).scaleb(-digits)
    rounded = Decimal(repr(as_number(value))).quantize(quantum, rounding=how)
    return Scaled(float(rounded), places_of(value))


@m("Number.RoundDown", 1, 2)
def m_number_round_down(value: object, *rest: object) -> object:
    digits = int(as_number(_optional(rest, 0, 0)))
    scale = 10.0**digits
    return number_out(math.floor(as_number(value) * scale) / scale)


@m("Number.RoundUp", 1, 2)
def m_number_round_up(value: object, *rest: object) -> object:
    digits = int(as_number(_optional(rest, 0, 0)))
    scale = 10.0**digits
    return number_out(math.ceil(as_number(value) * scale) / scale)


@m("Number.Abs", 1, 1)
def m_number_abs(value: object) -> object:
    return None if value is None else number_out(abs(as_number(value)))


@m("Number.IntegerDivide", 2, 3)
def m_number_integer_divide(left: object, right: object, *rest: object) -> object:
    second = as_number(right)
    if second == 0:
        raise MError("Expression.Error", "cannot divide by zero")
    return number_out(math.trunc(as_number(left) / second))


@m("Number.Mod", 2, 3)
def m_number_mod(left: object, right: object, *rest: object) -> object:
    second = as_number(right)
    if second == 0:
        raise MError("Expression.Error", "cannot divide by zero")
    first = as_number(left)
    return number_out(first - second * math.trunc(first / second))


@m("Number.Power", 2, 2)
def m_number_power(left: object, right: object) -> object:
    return number_out(as_number(left) ** as_number(right))


@m("Number.Sqrt", 1, 1)
def m_number_sqrt(value: object) -> object:
    """The root of a negative number is NaN to M, not an error."""
    number = as_number(value)
    return number_out(math.nan if number < 0 else math.sqrt(number))


@m("Number.IsEven", 1, 1)
def m_number_is_even(value: object) -> object:
    return int(as_number(value)) % 2 == 0


@m("Number.IsOdd", 1, 1)
def m_number_is_odd(value: object) -> object:
    return int(as_number(value)) % 2 == 1


# --- Text ----------------------------------------------------------------------------------------


@m("Text.Length", 1, 1)
def m_text_length(value: object) -> object:
    return None if value is None else len(as_text(value))


@m("Text.Upper", 1, 2)
def m_text_upper(value: object, *rest: object) -> object:
    return None if value is None else as_text(value).upper()


@m("Text.Lower", 1, 2)
def m_text_lower(value: object, *rest: object) -> object:
    return None if value is None else as_text(value).lower()


@m("Text.Proper", 1, 2)
def m_text_proper(value: object, *rest: object) -> object:
    if value is None:
        return None
    return re.sub(
        r"[A-Za-z]+",
        lambda match: match.group(0)[0].upper() + match.group(0)[1:].lower(),
        as_text(value),
    )


@m("Text.Trim", 1, 2)
def m_text_trim(value: object, *rest: object) -> object:
    return None if value is None else as_text(value).strip()


@m("Text.TrimStart", 1, 2)
def m_text_trim_start(value: object, *rest: object) -> object:
    return None if value is None else as_text(value).lstrip()


@m("Text.TrimEnd", 1, 2)
def m_text_trim_end(value: object, *rest: object) -> object:
    return None if value is None else as_text(value).rstrip()


@m("Text.Start", 2, 2)
def m_text_start(value: object, count: object) -> object:
    return None if value is None else as_text(value)[: int(as_number(count))]


@m("Text.End", 2, 2)
def m_text_end(value: object, count: object) -> object:
    if value is None:
        return None
    body = as_text(value)
    wanted = int(as_number(count))
    return body[max(0, len(body) - wanted) :] if wanted else ""


@m("Text.Middle", 2, 3)
def m_text_middle(value: object, start: object, *rest: object) -> object:
    if value is None:
        return None
    body = as_text(value)
    at = int(as_number(start))
    count = _optional(rest, 0)
    if count is None:
        return body[at:]
    return body[at : at + int(as_number(count))]


@m("Text.Range", 2, 3)
def m_text_range(value: object, start: object, *rest: object) -> object:
    return m_text_middle(value, start, *rest)


@m("Text.Contains", 2, 3)
def m_text_contains(value: object, needle: object, *rest: object) -> object:
    return None if value is None else as_text(needle) in as_text(value)


@m("Text.StartsWith", 2, 3)
def m_text_starts_with(value: object, needle: object, *rest: object) -> object:
    return None if value is None else as_text(value).startswith(as_text(needle))


@m("Text.EndsWith", 2, 3)
def m_text_ends_with(value: object, needle: object, *rest: object) -> object:
    return None if value is None else as_text(value).endswith(as_text(needle))


@m("Text.Replace", 3, 3)
def m_text_replace(value: object, old: object, new: object) -> object:
    return None if value is None else as_text(value).replace(as_text(old), as_text(new))


@m("Text.PositionOf", 2, 4)
def m_text_position_of(value: object, needle: object, *rest: object) -> object:
    return as_text(value).find(as_text(needle))


@m("Text.Split", 2, 2)
def m_text_split(value: object, separator: object) -> object:
    return list(as_text(value).split(as_text(separator)))


@m("Text.Combine", 1, 2)
def m_text_combine(values: object, *rest: object) -> object:
    separator = as_text(_optional(rest, 0, ""))
    return separator.join(as_text(one) for one in as_list(values) if one is not None)


@m("Text.Repeat", 2, 2)
def m_text_repeat(value: object, count: object) -> object:
    return as_text(value) * int(as_number(count))


@m("Text.PadStart", 2, 3)
def m_text_pad_start(value: object, width: object, *rest: object) -> object:
    filler = as_text(_optional(rest, 0, " "))
    return as_text(value).rjust(int(as_number(width)), filler[0] if filler else " ")


@m("Text.PadEnd", 2, 3)
def m_text_pad_end(value: object, width: object, *rest: object) -> object:
    filler = as_text(_optional(rest, 0, " "))
    return as_text(value).ljust(int(as_number(width)), filler[0] if filler else " ")


@m("Text.Insert", 3, 3)
def m_text_insert(value: object, at: object, inserted: object) -> object:
    body = as_text(value)
    position = int(as_number(at))
    return body[:position] + as_text(inserted) + body[position:]


@m("Text.Remove", 2, 2)
def m_text_remove(value: object, removed: object) -> object:
    body = as_text(value)
    drop = removed if is_list(removed) else [removed]
    for one in drop:
        body = body.replace(as_text(one), "")
    return body


@m("Text.Select", 2, 2)
def m_text_select(value: object, kept: object) -> object:
    allowed = {as_text(one) for one in _items(kept)}
    expanded: set[str] = set()
    for piece in allowed:
        expanded.update(piece)
    return "".join(char for char in as_text(value) if char in expanded)


@m("Text.ToList", 1, 1)
def m_text_to_list(value: object) -> object:
    return list(as_text(value))


@m("Text.Format", 2, 3)
def m_text_format(pattern: object, values: object, *rest: object) -> object:
    body = as_text(pattern)
    items = values if is_list(values) else [values]
    for index, one in enumerate(items):
        body = body.replace(f"#{{{index}}}", as_text(one) if one is not None else "")
    if isinstance(values, Record):
        for name, one in values.fields.items():
            body = body.replace(f"#{{{name}}}", as_text(one) if one is not None else "")
    return body


# --- List -----------------------------------------------------------------------------------------


def _numbers_of(values: list[object]) -> list[float]:
    return [as_number(one) for one in values if one is not None and not isinstance(one, str)]


@m("List.Sum", 1, 2)
def m_list_sum(values: object, *rest: object) -> object:
    numbers = _numbers_of(as_list(values))
    return number_out(math.fsum(numbers)) if numbers else None


@m("List.Average", 1, 2)
def m_list_average(values: object, *rest: object) -> object:
    numbers = _numbers_of(as_list(values))
    return number_out(math.fsum(numbers) / len(numbers)) if numbers else None


@m("List.Count", 1, 1)
def m_list_count(values: object) -> object:
    return len(as_list(values))


@m("List.NonNullCount", 1, 1)
def m_list_non_null_count(values: object) -> object:
    return sum(1 for one in as_list(values) if one is not None)


@m("List.Max", 1, 4)
def m_list_max(values: object, *rest: object) -> object:
    items = [one for one in as_list(values) if one is not None]
    return max(items, key=_sort_key) if items else _optional(rest, 0)


@m("List.Min", 1, 4)
def m_list_min(values: object, *rest: object) -> object:
    items = [one for one in as_list(values) if one is not None]
    return min(items, key=_sort_key) if items else _optional(rest, 0)


@m("List.First", 1, 2)
def m_list_first(values: object, *rest: object) -> object:
    items = as_list(values)
    return items[0] if items else _optional(rest, 0)


@m("List.Last", 1, 2)
def m_list_last(values: object, *rest: object) -> object:
    items = as_list(values)
    return items[-1] if items else _optional(rest, 0)


@m("List.FirstN", 2, 2)
def m_list_first_n(values: object, count: object) -> object:
    items = as_list(values)
    if callable(getattr(count, "call", None)) or type_name(count) == "function":
        kept: list[object] = []
        for one in items:
            if not as_logical(_call(count, [one])):
                break
            kept.append(one)
        return kept
    return items[: int(as_number(count))]


@m("List.Skip", 1, 2)
def m_list_skip(values: object, *rest: object) -> object:
    items = as_list(values)
    count = _optional(rest, 0, 1)
    return items[int(as_number(count)) :]


@m("List.Range", 2, 3)
def m_list_range(values: object, start: object, *rest: object) -> object:
    items = as_list(values)
    at = int(as_number(start))
    count = _optional(rest, 0)
    return items[at:] if count is None else items[at : at + int(as_number(count))]


@m("List.Distinct", 1, 2)
def m_list_distinct(values: object, *rest: object) -> object:
    out: list[object] = []
    for one in as_list(values):
        if not any(_same(one, kept) for kept in out):
            out.append(one)
    return out


@m("List.Contains", 2, 3)
def m_list_contains(values: object, wanted: object, *rest: object) -> object:
    return any(_same(one, wanted) for one in as_list(values))


@m("List.PositionOf", 2, 4)
def m_list_position_of(values: object, wanted: object, *rest: object) -> object:
    for index, one in enumerate(as_list(values)):
        if _same(one, wanted):
            return index
    return -1


@m("List.Transform", 2, 2)
def m_list_transform(values: object, function: object) -> object:
    return [_call(function, [one]) for one in as_list(values)]


@m("List.Select", 2, 2)
def m_list_select(values: object, condition: object) -> object:
    return [one for one in as_list(values) if _kept(condition, one)]


@m("List.RemoveNulls", 1, 1)
def m_list_remove_nulls(values: object) -> object:
    return [one for one in as_list(values) if one is not None]


@m("List.Sort", 1, 2)
def m_list_sort(values: object, *rest: object) -> object:
    items = list(as_list(values))
    order = _optional(rest, 0, 0)
    if type_name(order) == "function":
        return sorted(items, key=lambda one: _sort_key(_call(order, [one])))
    return sorted(items, key=_sort_key, reverse=bool(order == 1))


@m("List.Reverse", 1, 1)
def m_list_reverse(values: object) -> object:
    return list(reversed(as_list(values)))


@m("List.Combine", 1, 1)
def m_list_combine(values: object) -> object:
    out: list[object] = []
    for one in as_list(values):
        out.extend(as_list(one))
    return out


@m("List.Numbers", 2, 3)
def m_list_numbers(start: object, count: object, *rest: object) -> object:
    step = as_number(_optional(rest, 0, 1))
    first = as_number(start)
    return [number_out(first + step * index) for index in range(int(as_number(count)))]


@m("List.Zip", 1, 1)
def m_list_zip(values: object) -> object:
    lists = [as_list(one) for one in as_list(values)]
    if not lists:
        return []
    longest = max(len(one) for one in lists)
    return [
        [one[index] if index < len(one) else None for one in lists] for index in range(longest)
    ]


@m("List.Accumulate", 3, 3)
def m_list_accumulate(values: object, seed: object, function: object) -> object:
    total = seed
    for one in as_list(values):
        total = _call(function, [total, one])
    return total


@m("List.IsEmpty", 1, 1)
def m_list_is_empty(values: object) -> object:
    return not as_list(values)


@m("List.ReplaceValue", 4, 4)
def m_list_replace_value(values: object, old: object, new: object, replacer: object) -> object:
    return [new if _same(one, old) else one for one in as_list(values)]


def _same(left: object, right: object) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return left is right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return float(left) == float(right)
    return left == right


def _sort_key(value: object) -> tuple[int, Any]:
    """An ordering that puts the types in M's order and never raises."""
    if value is None:
        return (0, 0)
    if isinstance(value, bool):
        return (1, int(value))
    if isinstance(value, (int, float)):
        return (2, float(value))
    if isinstance(value, _dt.datetime):
        return (3, value.timestamp())
    if isinstance(value, _dt.date):
        return (3, _dt.datetime(value.year, value.month, value.day).timestamp())
    if isinstance(value, str):
        return (4, value)
    return (5, str(value))


# --- Record ---------------------------------------------------------------------------------------


@m("Record.Field", 2, 2)
def m_record_field(record: object, name: object) -> object:
    return as_record(record)[as_text(name)]


@m("Record.FieldOrDefault", 2, 3)
def m_record_field_or_default(record: object, name: object, *rest: object) -> object:
    return as_record(record).get(as_text(name), _optional(rest, 0))


@m("Record.FieldNames", 1, 1)
def m_record_field_names(record: object) -> object:
    return as_record(record).names


@m("Record.FieldValues", 1, 1)
def m_record_field_values(record: object) -> object:
    return list(as_record(record).fields.values())


@m("Record.AddField", 3, 4)
def m_record_add_field(record: object, name: object, value: object, *rest: object) -> object:
    fields = dict(as_record(record).fields)
    fields[as_text(name)] = value
    return Record(fields)


@m("Record.RemoveFields", 2, 3)
def m_record_remove_fields(record: object, names: object, *rest: object) -> object:
    drop = {as_text(one) for one in _items(names)}
    return Record({name: value for name, value in as_record(record).fields.items() if name not in drop})


@m("Record.SelectFields", 2, 3)
def m_record_select_fields(record: object, names: object, *rest: object) -> object:
    wanted = [as_text(one) for one in _items(names)]
    source = as_record(record)
    return Record({name: source.get(name) for name in wanted})


@m("Record.RenameFields", 2, 3)
def m_record_rename_fields(record: object, renames: object, *rest: object) -> object:
    pairs = renames if is_list(renames) else []
    mapping = {as_text(one[0]): as_text(one[1]) for one in (as_list(pair) for pair in pairs)}
    return Record(
        {mapping.get(name, name): value for name, value in as_record(record).fields.items()}
    )


@m("Record.HasFields", 2, 2)
def m_record_has_fields(record: object, names: object) -> object:
    wanted = [as_text(one) for one in _items(names)]
    source = as_record(record)
    return all(source.has(name) for name in wanted)


@m("Record.Combine", 1, 1)
def m_record_combine(records: object) -> object:
    fields: dict[str, object] = {}
    for one in as_list(records):
        fields.update(as_record(one).fields)
    return Record(fields)


@m("Record.FromList", 2, 2)
def m_record_from_list(values: object, names: object) -> object:
    keys = [as_text(one) for one in as_list(names)]
    return Record(dict(zip(keys, as_list(values))))


@m("Record.ToList", 1, 1)
def m_record_to_list(record: object) -> object:
    return list(as_record(record).fields.values())


@m("Record.ToTable", 1, 1)
def m_record_to_table(record: object) -> object:
    source = as_record(record)
    return Table(["Name", "Value"], [[name, value] for name, value in source.fields.items()])


# --- Table ------------------------------------------------------------------------------------------


@m("Table.FromRows", 1, 2)
def m_table_from_rows(rows: object, *rest: object) -> object:
    columns = _optional(rest, 0)
    names: list[str] = []
    if is_list(columns):
        names = [as_text(one) for one in as_list(columns)]
    elif isinstance(columns, MType) and is_list(columns.detail):
        names = [as_text(one) for one in as_list(columns.detail)]
    body = [list(as_list(row)) for row in as_list(rows)]
    if not names:
        width = max((len(row) for row in body), default=0)
        names = [f"Column{index + 1}" for index in range(width)]
    return Table(names, body)


@m("Table.FromColumns", 1, 2)
def m_table_from_columns(columns: object, *rest: object) -> object:
    lists = [as_list(one) for one in as_list(columns)]
    names = [f"Column{index + 1}" for index in range(len(lists))]
    given = _optional(rest, 0)
    if is_list(given):
        names = [as_text(one) for one in given]
    height = max((len(one) for one in lists), default=0)
    rows: list[list[object]] = [
        [one[index] if index < len(one) else None for one in lists] for index in range(height)
    ]
    return Table(names, rows)


@m("Table.FromRecords", 1, 3)
def m_table_from_records(records: object, *rest: object) -> object:
    given = _optional(rest, 0)
    columns = [as_text(one) for one in given] if is_list(given) else None
    return Table.from_records([as_record(one) for one in as_list(records)], columns)


@m("Table.FromList", 1, 4)
def m_table_from_list(values: object, *rest: object) -> object:
    items = as_list(values)
    splitter = _optional(rest, 0)
    if splitter is None:
        return Table(["Column1"], [[one] for one in items])
    rows = [as_list(_call(splitter, [one])) for one in items]
    width = max((len(row) for row in rows), default=1)
    names = [f"Column{index + 1}" for index in range(width)]
    return Table(names, [list(row) + [None] * (width - len(row)) for row in rows])


@m("Table.ToRows", 1, 1)
def m_table_to_rows(table: object) -> object:
    return [list(row) for row in as_table(table).rows]


@m("Table.ToColumns", 1, 1)
def m_table_to_columns(table: object) -> object:
    source = as_table(table)
    return [source.column(name) for name in source.columns]


@m("Table.ToRecords", 1, 1)
def m_table_to_records(table: object) -> object:
    return list(as_table(table).records())


@m("Table.ColumnNames", 1, 1)
def m_table_column_names(table: object) -> object:
    return list(as_table(table).columns)


@m("Table.Column", 2, 2)
def m_table_column(table: object, name: object) -> object:
    return as_table(table).column(as_text(name))


@m("Table.RowCount", 1, 1)
def m_table_row_count(table: object) -> object:
    return as_table(table).height


@m("Table.ColumnCount", 1, 1)
def m_table_column_count(table: object) -> object:
    return as_table(table).width


@m("Table.IsEmpty", 1, 1)
def m_table_is_empty(table: object) -> object:
    return as_table(table).height == 0


@m("Table.Buffer", 1, 1)
def m_table_buffer(table: object) -> object:
    return table


@m("Table.PromoteHeaders", 1, 2)
def m_table_promote_headers(table: object, *rest: object) -> object:
    source = as_table(table)
    if not source.rows:
        return source
    head = source.rows[0]
    names: list[str] = []
    for index, value in enumerate(head):
        name = as_text(value) if value is not None and value != "" else f"Column{index + 1}"
        while name in names:
            name = f"{name}_1"
        names.append(name)
    return Table(names, [list(row) for row in source.rows[1:]])


@m("Table.DemoteHeaders", 1, 1)
def m_table_demote_headers(table: object) -> object:
    source = as_table(table)
    names = [f"Column{index + 1}" for index in range(source.width)]
    return Table(names, [list(source.columns), *[list(row) for row in source.rows]])


@m("Table.SelectColumns", 2, 3)
def m_table_select_columns(table: object, names: object, *rest: object) -> object:
    source = as_table(table)
    wanted = [as_text(one) for one in _items(names)]
    missing_ok = _optional(rest, 0, 0)
    indexes: list[int | None] = []
    for name in wanted:
        if name in source.columns:
            indexes.append(source.columns.index(name))
        elif missing_ok in (1, 2):
            indexes.append(None)
        else:
            source.column_index(name)
            indexes.append(None)
    rows = [[row[at] if at is not None else None for at in indexes] for row in source.rows]
    return Table(wanted, rows)


@m("Table.RemoveColumns", 2, 3)
def m_table_remove_columns(table: object, names: object, *rest: object) -> object:
    source = as_table(table)
    drop = {as_text(one) for one in _items(names)}
    kept = [name for name in source.columns if name not in drop]
    indexes = [source.columns.index(name) for name in kept]
    return Table(kept, [[row[at] for at in indexes] for row in source.rows])


@m("Table.RenameColumns", 2, 3)
def m_table_rename_columns(table: object, renames: object, *rest: object) -> object:
    source = as_table(table)
    pairs = as_list(renames)
    mapping: dict[str, str] = {}
    if pairs and not isinstance(pairs[0], list):
        mapping[as_text(pairs[0])] = as_text(pairs[1])
    else:
        for pair in pairs:
            both = as_list(pair)
            mapping[as_text(both[0])] = as_text(both[1])
    return Table([mapping.get(name, name) for name in source.columns], [list(row) for row in source.rows])


@m("Table.ReorderColumns", 2, 3)
def m_table_reorder_columns(table: object, order: object, *rest: object) -> object:
    source = as_table(table)
    wanted = [as_text(one) for one in as_list(order)]
    rest_names = [name for name in source.columns if name not in wanted]
    names = wanted + rest_names
    indexes = [source.column_index(name) for name in names]
    return Table(names, [[row[at] for at in indexes] for row in source.rows])


def _kept(condition: object, row: object) -> bool:
    """Whether a row passes: a test that comes out null keeps nothing."""
    answer = _call(condition, [row])
    return answer is not None and as_logical(answer)


@m("Table.SelectRows", 2, 2)
def m_table_select_rows(table: object, condition: object) -> object:
    source = as_table(table)
    kept = [row for row in source.records() if _kept(condition, row)]
    return Table(list(source.columns), [[row.get(name) for name in source.columns] for row in kept])


@m("Table.RemoveRows", 2, 3)
def m_table_remove_rows(table: object, start: object, *rest: object) -> object:
    source = as_table(table)
    at = int(as_number(start))
    count = int(as_number(_optional(rest, 0, 1)))
    rows = [row for index, row in enumerate(source.rows) if not at <= index < at + count]
    return Table(list(source.columns), rows)


@m("Table.FirstN", 2, 2)
def m_table_first_n(table: object, count: object) -> object:
    source = as_table(table)
    if type_name(count) == "function":
        kept: list[list[object]] = []
        for row in source.records():
            if not as_logical(_call(count, [row])):
                break
            kept.append([row.get(name) for name in source.columns])
        return Table(list(source.columns), kept)
    return Table(list(source.columns), [list(row) for row in source.rows[: int(as_number(count))]])


@m("Table.LastN", 2, 2)
def m_table_last_n(table: object, count: object) -> object:
    source = as_table(table)
    wanted = int(as_number(count))
    return Table(list(source.columns), [list(row) for row in source.rows[max(0, source.height - wanted) :]])


@m("Table.Skip", 1, 2)
def m_table_skip(table: object, *rest: object) -> object:
    source = as_table(table)
    count = int(as_number(_optional(rest, 0, 1)))
    return Table(list(source.columns), [list(row) for row in source.rows[count:]])


@m("Table.Range", 2, 3)
def m_table_range(table: object, start: object, *rest: object) -> object:
    source = as_table(table)
    at = int(as_number(start))
    count = _optional(rest, 0)
    rows = source.rows[at:] if count is None else source.rows[at : at + int(as_number(count))]
    return Table(list(source.columns), [list(row) for row in rows])


@m("Table.First", 1, 2)
def m_table_first(table: object, *rest: object) -> object:
    source = as_table(table)
    return source.record_at(0) if source.height else _optional(rest, 0)


@m("Table.Last", 1, 2)
def m_table_last(table: object, *rest: object) -> object:
    source = as_table(table)
    return source.record_at(source.height - 1) if source.height else _optional(rest, 0)


@m("Table.Distinct", 1, 2)
def m_table_distinct(table: object, *rest: object) -> object:
    source = as_table(table)
    columns = _optional(rest, 0)
    if columns is None:
        keys = list(range(source.width))
    else:
        wanted = [as_text(one) for one in _items(columns)]
        keys = [source.column_index(name) for name in wanted]
    seen: list[tuple[object, ...]] = []
    rows: list[list[object]] = []
    for row in source.rows:
        key = tuple(row[at] for at in keys)
        if key not in seen:
            seen.append(key)
            rows.append(list(row))
    return Table(list(source.columns), rows)


@m("Table.Sort", 2, 2)
def m_table_sort(table: object, order: object) -> object:
    source = as_table(table)
    clauses: list[tuple[str, int]] = []
    items = _items(order)
    for one in items:
        if is_list(one):
            pair = as_list(one)
            clauses.append((as_text(pair[0]), int(as_number(pair[1])) if len(pair) > 1 else 0))
        else:
            clauses.append((as_text(one), 0))
    rows = [list(row) for row in source.rows]
    for name, direction in reversed(clauses):
        at = source.column_index(name)
        rows.sort(key=lambda row, at=at: _sort_key(row[at]), reverse=direction == 1)
    return Table(list(source.columns), rows)


@m("Table.AddColumn", 3, 4)
def m_table_add_column(table: object, name: object, generator: object, *rest: object) -> object:
    source = as_table(table)
    column = as_text(name)
    rows = [[*row, _call(generator, [record])] for row, record in zip(source.rows, source.records())]
    return Table([*source.columns, column], rows)


@m("Table.AddIndexColumn", 2, 5)
def m_table_add_index_column(table: object, name: object, *rest: object) -> object:
    source = as_table(table)
    start = int(as_number(_optional(rest, 0, 0)))
    step = int(as_number(_optional(rest, 1, 1)))
    rows = [[*row, start + index * step] for index, row in enumerate(source.rows)]
    return Table([*source.columns, as_text(name)], rows)


@m("Table.TransformColumns", 2, 3)
def m_table_transform_columns(table: object, transforms: object, *rest: object) -> object:
    source = as_table(table)
    pairs = _items(transforms)
    if pairs and not isinstance(pairs[0], list):
        pairs = [pairs]
    rows = [list(row) for row in source.rows]
    for pair in pairs:
        both = as_list(pair)
        at = source.column_index(as_text(both[0]))
        function = both[1]
        for row in rows:
            row[at] = _call(function, [row[at]])
    return Table(list(source.columns), rows)


@m("Table.TransformColumnTypes", 2, 3)
def m_table_transform_column_types(table: object, types: object, *rest: object) -> object:
    """The step the query editor writes when you set a column's type."""
    source = as_table(table)
    pairs = _items(types)
    if pairs and not isinstance(pairs[0], list):
        pairs = [pairs]
    rows = [list(row) for row in source.rows]
    for pair in pairs:
        both = as_list(pair)
        name = as_text(both[0])
        if name not in source.columns:
            continue
        at = source.columns.index(name)
        wanted = both[1]
        for row in rows:
            row[at] = _converted(row[at], wanted)
    return Table(list(source.columns), rows)


def _converted(value: object, wanted: object) -> object:
    if value is None:
        return None
    kind = wanted.name if isinstance(wanted, MType) else as_text(wanted)
    if kind.startswith("nullable "):
        kind = kind[9:]
    if kind == "text":
        return m_text_from(value)
    if kind == "number":
        return _to_number(value)
    if kind == "logical":
        return m_logical_from(value)
    if kind == "date":
        return m_date_from(value)
    if kind == "datetime":
        return m_datetime_from(value)
    if kind == "time":
        return m_time_from(value)
    return value


@m("Table.ReplaceValue", 5, 5)
def m_table_replace_value(
    table: object, old: object, new: object, replacer: object, columns: object
) -> object:
    """The replacer decides whether a cell is matched whole or searched."""
    source = as_table(table)
    wanted = [as_text(one) for one in _items(columns)]
    indexes = [source.column_index(name) for name in wanted]
    rows = [list(row) for row in source.rows]
    for row in rows:
        for at in indexes:
            row[at] = _call(replacer, [row[at], old, new]) if replacer is not None else row[at]
    return Table(list(source.columns), rows)


@m("Table.Combine", 1, 2)
def m_table_combine(tables: object, *rest: object) -> object:
    return combine_tables([as_table(one) for one in as_list(tables)])


def combine_tables(tables: list[Table]) -> Table:
    """Several tables stacked, sharing whatever columns they have."""
    names: list[str] = []
    for one in tables:
        for name in one.columns:
            if name not in names:
                names.append(name)
    rows: list[list[object]] = []
    for one in tables:
        for record in one.records():
            rows.append([record.get(name) for name in names])
    return Table(names, rows)


@m("Table.Group", 3, 5)
def m_table_group(table: object, keys: object, aggregations: object, *rest: object) -> object:
    source = as_table(table)
    wanted = [as_text(one) for one in _items(keys)]
    indexes = [source.column_index(name) for name in wanted]
    groups: list[tuple[tuple[object, ...], list[list[object]]]] = []
    for row in source.rows:
        key = tuple(row[at] for at in indexes)
        for held, rows in groups:
            if held == key:
                rows.append(list(row))
                break
        else:
            groups.append((key, [list(row)]))
    specs = _items(aggregations)
    if specs and not isinstance(specs[0], list):
        specs = [specs]
    names = wanted + [as_text(as_list(one)[0]) for one in specs]
    out: list[list[object]] = []
    for key, rows in groups:
        made = list(key)
        piece = Table(list(source.columns), rows)
        for one in specs:
            both = as_list(one)
            made.append(_call(both[1], [piece]))
        out.append(made)
    return Table(names, out)


@m("Table.Join", 4, 7)
def m_table_join(
    left: object, left_keys: object, right: object, right_keys: object, *rest: object
) -> object:
    kind = _optional(rest, 0, "Inner")
    return _join(as_table(left), left_keys, as_table(right), right_keys, as_text(kind), None)


@m("Table.NestedJoin", 5, 7)
def m_table_nested_join(
    left: object, left_keys: object, right: object, right_keys: object, into: object, *rest: object
) -> object:
    kind = _optional(rest, 0, "LeftOuter")
    return _join(as_table(left), left_keys, as_table(right), right_keys, as_text(kind), as_text(into))


def _join(
    left: Table, left_keys: object, right: Table, right_keys: object, kind: str, into: str | None
) -> Table:
    first = [as_text(one) for one in _items(left_keys)]
    second = [as_text(one) for one in _items(right_keys)]
    left_at = [left.column_index(name) for name in first]
    right_at = [right.column_index(name) for name in second]
    keep_right = [index for index, name in enumerate(right.columns) if name not in second]
    if into is not None:
        names = [*left.columns, into]
    else:
        names = [*left.columns, *[name for name in right.columns if name not in second]]
    rows: list[list[object]] = []
    matched_right: set[int] = set()
    for row in left.rows:
        key = tuple(row[at] for at in left_at)
        hits = [
            index
            for index, other in enumerate(right.rows)
            if tuple(other[at] for at in right_at) == key
        ]
        matched_right.update(hits)
        if into is not None:
            nested = Table(list(right.columns), [list(right.rows[index]) for index in hits])
            if kind == "Inner" and not hits:
                continue
            if kind == "LeftAnti" and hits:
                continue
            rows.append([*row, nested])
            continue
        if not hits:
            if kind in ("LeftOuter", "FullOuter", "LeftAnti"):
                rows.append([*row, *[None] * len(keep_right)])
            continue
        if kind == "LeftAnti":
            continue
        for index in hits:
            rows.append([*row, *[right.rows[index][at] for at in keep_right]])
    if into is None and kind in ("RightOuter", "FullOuter"):
        for index, other in enumerate(right.rows):
            if index not in matched_right:
                made: list[object] = [None] * len(left.columns)
                for position, at in enumerate(left_at):
                    made[at] = other[right_at[position]]
                rows.append([*made, *[other[at] for at in keep_right]])
    return Table(names, rows)


@m("Table.ExpandTableColumn", 3, 4)
def m_table_expand_table_column(table: object, name: object, columns: object, *rest: object) -> object:
    source = as_table(table)
    at = source.column_index(as_text(name))
    wanted = [as_text(one) for one in as_list(columns)]
    given = _optional(rest, 0)
    new_names = [as_text(one) for one in as_list(given)] if is_list(given) else list(wanted)
    names = [one for index, one in enumerate(source.columns) if index != at]
    names = source.columns[:at] + new_names + source.columns[at + 1 :]
    rows: list[list[object]] = []
    for row in source.rows:
        nested = row[at]
        head = list(row[:at])
        tail = list(row[at + 1 :])
        if isinstance(nested, Table) and nested.height:
            for record in nested.records():
                rows.append([*head, *[record.get(one) for one in wanted], *tail])
        elif isinstance(nested, Record):
            rows.append([*head, *[nested.get(one) for one in wanted], *tail])
        else:
            rows.append([*head, *[None] * len(wanted), *tail])
    return Table(names, rows)


@m("Table.ExpandRecordColumn", 3, 4)
def m_table_expand_record_column(table: object, name: object, columns: object, *rest: object) -> object:
    return m_table_expand_table_column(table, name, columns, *rest)


@m("Table.SplitColumn", 3, 6)
def m_table_split_column(table: object, name: object, splitter: object, *rest: object) -> object:
    source = as_table(table)
    at = source.column_index(as_text(name))
    given = _optional(rest, 0)
    pieces = [as_list(_call(splitter, [row[at]])) for row in source.rows]
    width = max((len(one) for one in pieces), default=1)
    new_names = (
        [as_text(one) for one in as_list(given)]
        if is_list(given)
        else [f"{as_text(name)}.{index + 1}" for index in range(width)]
    )
    width = max(width, len(new_names))
    names = source.columns[:at] + new_names + source.columns[at + 1 :]
    rows: list[list[object]] = []
    for row, made in zip(source.rows, pieces):
        padded = list(made) + [None] * (len(new_names) - len(made))
        rows.append([*row[:at], *padded[: len(new_names)], *row[at + 1 :]])
    return Table(names, rows)


@m("Table.Schema", 1, 1)
def m_table_schema(table: object) -> object:
    source = as_table(table)
    return Table(
        ["Name", "Position"], [[name, index] for index, name in enumerate(source.columns)]
    )


# --- splitters and sources -------------------------------------------------------------------------


@m("Splitter.SplitTextByDelimiter", 1, 2)
def m_splitter_by_delimiter(delimiter: object, *rest: object) -> object:
    separator = as_text(delimiter)

    def split(value: object) -> object:
        return list(as_text(value).split(separator)) if value is not None else [None]

    return Builtin(name="Splitter", call=split, minimum=1, maximum=1)


@m("Splitter.SplitTextByEachDelimiter", 1, 3)
def m_splitter_by_each(delimiters: object, *rest: object) -> object:
    parts = [as_text(one) for one in as_list(delimiters)]

    def split(value: object) -> object:
        body = as_text(value)
        out: list[object] = []
        for piece in parts:
            head, _, body = body.partition(piece)
            out.append(head)
        out.append(body)
        return out

    return Builtin(name="Splitter", call=split, minimum=1, maximum=1)


@m("Csv.Document", 1, 5)
def m_csv_document(source: object, *rest: object) -> object:
    """A CSV read into a table, with the delimiter the options name."""
    text = source.decode("utf-8-sig") if isinstance(source, bytes) else as_text(source)
    delimiter = ","
    columns: list[str] | None = None
    options = _optional(rest, 0)
    if isinstance(options, Record):
        if options.has("Delimiter"):
            delimiter = as_text(options["Delimiter"])
        if options.has("Columns") and isinstance(options["Columns"], list):
            columns = [as_text(one) for one in as_list(options["Columns"])]
    elif isinstance(options, str):
        delimiter = options
    elif is_list(options):
        columns = [as_text(one) for one in options]
    second = _optional(rest, 1)
    if isinstance(second, str):
        delimiter = second
    rows = _read_csv(text, delimiter)
    width = max((len(row) for row in rows), default=0)
    if columns is None:
        columns = [f"Column{index + 1}" for index in range(width)]
    width = max(width, len(columns))
    padded = [list(row) + [None] * (width - len(row)) for row in rows]
    return Table(columns, padded)


def _read_csv(text: str, delimiter: str) -> list[list[object]]:
    """A CSV reader that handles the quoting a file actually carries."""
    rows: list[list[object]] = []
    row: list[object] = []
    field_text: list[str] = []
    quoted = False
    index = 0
    while index < len(text):
        char = text[index]
        if quoted:
            if char == '"':
                if index + 1 < len(text) and text[index + 1] == '"':
                    field_text.append('"')
                    index += 2
                    continue
                quoted = False
                index += 1
                continue
            field_text.append(char)
            index += 1
            continue
        if char == '"' and not field_text:
            quoted = True
            index += 1
            continue
        if text.startswith(delimiter, index):
            row.append("".join(field_text))
            field_text = []
            index += len(delimiter)
            continue
        if char in "\r\n":
            row.append("".join(field_text))
            rows.append(row)
            row = []
            field_text = []
            index += 2 if text.startswith("\r\n", index) else 1
            continue
        field_text.append(char)
        index += 1
    if field_text or row:
        row.append("".join(field_text))
        rows.append(row)
    return [one for one in rows if any(piece != "" for piece in one)]


@m("Json.Document", 1, 2)
def m_json_document(source: object, *rest: object) -> object:
    import json

    text = source.decode("utf-8-sig") if isinstance(source, bytes) else as_text(source)
    return _from_json(json.loads(text))


def _from_json(value: object) -> object:
    if is_mapping(value):
        return Record({str(name): _from_json(one) for name, one in value.items()})
    if is_list(value):
        return [_from_json(one) for one in value]
    if isinstance(value, bool) or value is None:
        return value
    if isinstance(value, (int, float)):
        return number_out(float(value))
    return value


@m("Lines.FromBinary", 1, 4)
def m_lines_from_binary(source: object, *rest: object) -> object:
    text = source.decode("utf-8-sig") if isinstance(source, bytes) else as_text(source)
    return text.splitlines()


@m("Text.FromBinary", 1, 3)
def m_text_from_binary(source: object, *rest: object) -> object:
    return source.decode("utf-8-sig") if isinstance(source, bytes) else as_text(source)


@m("Binary.FromText", 1, 2)
def m_binary_from_text(source: object, *rest: object) -> object:
    return as_text(source).encode("utf-8")


@m("Replacer.ReplaceText", 3, 3)
def m_replacer_replace_text(value: object, old: object, new: object) -> object:
    if value is None:
        return None
    return as_text(value).replace(as_text(old), as_text(new))


@m("Replacer.ReplaceValue", 3, 3)
def m_replacer_replace_value(value: object, old: object, new: object) -> object:
    return new if _same(value, old) else value


@m("Binary.ToText", 1, 2)
def m_binary_to_text(source: object, *rest: object) -> object:
    """A binary written out, base64 unless asked for hex."""
    import base64

    raw = source if isinstance(source, bytes) else as_text(source).encode("utf-8")
    if as_text(_optional(rest, 0, "Base64")) == "Hex":
        return raw.hex()
    return base64.b64encode(raw).decode("ascii")


# --- durations ---------------------------------------------------------------------------


def _as_duration(value: object) -> Duration:
    if isinstance(value, Duration):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return Duration(seconds=as_number(value) * 86400)
    raise MError(
        "Expression.Error",
        f"We cannot convert a value of type {type_name(value)} to type Duration.",
    )


@m("Duration.From", 1, 2)
def m_duration_from(value: object, *rest: object) -> object:
    """A number of days as a duration, which is how M counts one."""
    return None if value is None else _as_duration(value)


@m("Duration.Days", 1, 1)
def m_duration_days(value: object) -> object:
    return None if value is None else number_out(float(int(_as_duration(value).days)))


@m("Duration.Hours", 1, 1)
def m_duration_hours(value: object) -> object:
    return None if value is None else number_out(float(int(_as_duration(value).hours)))


@m("Duration.Minutes", 1, 1)
def m_duration_minutes(value: object) -> object:
    return None if value is None else number_out(float(int(_as_duration(value).minutes)))


@m("Duration.Seconds", 1, 1)
def m_duration_seconds(value: object) -> object:
    return None if value is None else number_out(float(int(_as_duration(value).seconds)))


@m("Duration.TotalDays", 1, 1)
def m_duration_total_days(value: object) -> object:
    return None if value is None else number_out(_as_duration(value).total_seconds / 86400)


@m("Duration.TotalHours", 1, 1)
def m_duration_total_hours(value: object) -> object:
    return None if value is None else number_out(_as_duration(value).total_seconds / 3600)


@m("Duration.TotalMinutes", 1, 1)
def m_duration_total_minutes(value: object) -> object:
    return None if value is None else number_out(_as_duration(value).total_seconds / 60)


@m("Duration.TotalSeconds", 1, 1)
def m_duration_total_seconds(value: object) -> object:
    return None if value is None else number_out(_as_duration(value).total_seconds)


@m("Duration.ToText", 1, 2)
def m_duration_to_text(value: object, *rest: object) -> object:
    return None if value is None else duration_text(_as_duration(value))


# --- characters and whole numbers --------------------------------------------------------


@m("Character.FromNumber", 1, 1)
def m_character_from_number(value: object) -> object:
    return None if value is None else chr(int(as_number(value)))


@m("Character.ToNumber", 1, 1)
def m_character_to_number(value: object) -> object:
    if value is None:
        return None
    text = as_text(value)
    if not text:
        raise MError("Expression.Error", "We cannot take the number of an empty character.")
    return ord(text[0])


@m("Int64.From", 1, 3)
def m_int64_from(value: object, *rest: object) -> object:
    """A whole number, rounded to even the way Number.Round is."""
    if value is None:
        return None
    rounded = m_number_round(_to_number(value), 0, _optional(rest, 1, "ToEven"))
    return int(as_number(rounded))


@m("Text.BetweenDelimiters", 3, 5)
def m_text_between_delimiters(value: object, start: object, end: object, *rest: object) -> object:
    """What lies between two delimiters, by default the first pair.

    The fourth argument counts start delimiters from zero, so 1 in
    "a[b]c[d]e" is the second bracket and answers "d".
    """
    if value is None:
        return None
    text = as_text(value)
    opener, closer = as_text(start), as_text(end)
    which = int(as_number(_first_of(_optional(rest, 0, 0))))
    which_end = int(as_number(_first_of(_optional(rest, 1, 0))))
    at = -len(opener)
    for _ in range(which + 1):
        at = text.find(opener, at + len(opener))
        if at < 0:
            return ""
    at += len(opener)
    stop = at - len(closer)
    for _ in range(which_end + 1):
        stop = text.find(closer, stop + len(closer))
        if stop < 0:
            return text[at:]
    return text[at:stop]


def _first_of(value: object) -> object:
    """A position argument, which M lets be a number or {number, from}."""
    return as_list(value)[0] if is_list(value) else value


@m("Uri.EscapeDataString", 1, 1)
def m_uri_escape_data_string(value: object) -> object:
    from urllib.parse import quote

    return None if value is None else quote(as_text(value), safe="-._~")


# M has Uri.EscapeDataString and no Uri.UnescapeDataString to go with
# it: asking for one is the name error any other misspelling gives.


# --- more of the list library -------------------------------------------------------------


@m("List.Generate", 3, 4)
def m_list_generate(initial: object, condition: object, next_one: object, *rest: object) -> object:
    """The loop M writes as a list: seed, keep going, step, take."""
    selector = _optional(rest, 0)
    out: list[object] = []
    state = _call(initial, [])
    while as_logical(_call(condition, [state])):
        out.append(state if selector is None else _call(selector, [state]))
        state = _call(next_one, [state])
        if len(out) > 1_000_000:
            raise MError("Expression.Error", "The list kept generating and did not stop.")
    return out


@m("List.Median", 1, 2)
def m_list_median(values: object, *rest: object) -> object:
    numbers = sorted(_numbers_of(as_list(values)))
    if not numbers:
        return None
    middle = len(numbers) // 2
    if len(numbers) % 2:
        return number_out(numbers[middle])
    return number_out((numbers[middle - 1] + numbers[middle]) / 2)


@m("List.StandardDeviation", 1, 1)
def m_list_standard_deviation(values: object) -> object:
    """The sample deviation, which is the one M answers with."""
    numbers = _numbers_of(as_list(values))
    if not numbers:
        return None
    if len(numbers) == 1:
        return 0
    mean = math.fsum(numbers) / len(numbers)
    spread = math.fsum((one - mean) ** 2 for one in numbers) / (len(numbers) - 1)
    return number_out(math.sqrt(spread))


# --- more of the table library ------------------------------------------------------------


def _filled(source: Table, names: object, *, upward: bool) -> Table:
    wanted = [source.column_index(as_text(one)) for one in _items(names)]
    rows = [list(row) for row in source.rows]
    order = range(len(rows) - 1, -1, -1) if upward else range(len(rows))
    carried: dict[int, object] = {}
    for index in order:
        for at in wanted:
            if rows[index][at] is None:
                rows[index][at] = carried.get(at)
            else:
                carried[at] = rows[index][at]
    return Table(list(source.columns), rows)


@m("Table.FillDown", 2, 2)
def m_table_fill_down(table: object, names: object) -> object:
    return _filled(as_table(table), names, upward=False)


@m("Table.FillUp", 2, 2)
def m_table_fill_up(table: object, names: object) -> object:
    return _filled(as_table(table), names, upward=True)


@m("Table.ReverseRows", 1, 1)
def m_table_reverse_rows(table: object) -> object:
    source = as_table(table)
    return Table(list(source.columns), [list(row) for row in reversed(source.rows)])


@m("Table.Repeat", 2, 2)
def m_table_repeat(table: object, count: object) -> object:
    source = as_table(table)
    times = int(as_number(count))
    return Table(list(source.columns), [list(row) for _ in range(times) for row in source.rows])


def _extreme(table: object, comparison: object, rest: tuple[object, ...], *, biggest: bool) -> object:
    source = as_table(table)
    if not source.rows:
        return _optional(rest, 0)
    if isinstance(comparison, str):
        at = source.column_index(as_text(comparison))
        keys = [_sort_key(row[at]) for row in source.rows]
    elif is_list(comparison):
        indexes = [source.column_index(as_text(one)) for one in as_list(comparison)]
        keys = [tuple(_sort_key(row[at]) for at in indexes) for row in source.rows]
    else:
        keys = [_sort_key(_call(comparison, [record])) for record in source.records()]
    best = 0
    for index in range(1, len(keys)):
        if (keys[index] > keys[best]) if biggest else (keys[index] < keys[best]):
            best = index
    return source.record_at(best)


@m("Table.Max", 2, 3)
def m_table_max(table: object, comparison: object, *rest: object) -> object:
    return _extreme(table, comparison, rest, biggest=True)


@m("Table.Min", 2, 3)
def m_table_min(table: object, comparison: object, *rest: object) -> object:
    return _extreme(table, comparison, rest, biggest=False)


@m("Table.Unpivot", 4, 4)
def m_table_unpivot(table: object, names: object, attribute: object, value: object) -> object:
    source = as_table(table)
    moving = [as_text(one) for one in _items(names)]
    kept = [name for name in source.columns if name not in moving]
    indexes = [source.columns.index(name) for name in kept]
    rows: list[list[object]] = []
    for row in source.rows:
        front = [row[at] for at in indexes]
        for name in moving:
            rows.append([*front, name, row[source.column_index(name)]])
    return Table([*kept, as_text(attribute), as_text(value)], rows)


@m("Table.UnpivotOtherColumns", 4, 4)
def m_table_unpivot_other_columns(
    table: object, names: object, attribute: object, value: object
) -> object:
    source = as_table(table)
    kept = {as_text(one) for one in _items(names)}
    moving = [name for name in source.columns if name not in kept]
    return m_table_unpivot(source, moving, attribute, value)


@m("Table.Pivot", 4, 5)
def m_table_pivot(
    table: object, values: object, attribute: object, value: object, *rest: object
) -> object:
    """The unpivot run backwards: one column per named attribute."""
    source = as_table(table)
    wanted = [as_text(one) for one in _items(values)]
    attribute_at = source.column_index(as_text(attribute))
    value_at = source.column_index(as_text(value))
    kept = [
        name
        for index, name in enumerate(source.columns)
        if index not in (attribute_at, value_at)
    ]
    keep_at = [source.columns.index(name) for name in kept]
    aggregate = _optional(rest, 0)
    groups: dict[tuple[object, ...], dict[str, list[object]]] = {}
    order: list[tuple[object, ...]] = []
    for row in source.rows:
        key = tuple(row[at] for at in keep_at)
        if key not in groups:
            groups[key] = {}
            order.append(key)
        name = as_text(row[attribute_at])
        groups[key].setdefault(name, []).append(row[value_at])
    rows: list[list[object]] = []
    for key in order:
        cells: list[object] = list(key)
        for name in wanted:
            found = groups[key].get(name, [])
            if aggregate is not None:
                cells.append(_call(aggregate, [found]))
            elif not found:
                cells.append(None)
            elif len(found) == 1:
                cells.append(found[0])
            else:
                raise MError("Expression.Error", "There were too many elements to pivot.")
        rows.append(cells)
    return Table([*kept, *wanted], rows)


@m("Table.TransformColumnNames", 2, 3)
def m_table_transform_column_names(table: object, rename: object, *rest: object) -> object:
    source = as_table(table)
    names = [as_text(_call(rename, [name])) for name in source.columns]
    return Table(names, [list(row) for row in source.rows])


@m("Record.TransformFields", 2, 3)
def m_record_transform_fields(record: object, transforms: object, *rest: object) -> object:
    source = as_record(record)
    pairs = _items(transforms)
    if pairs and not is_list(pairs[0]):
        pairs = [pairs]
    fields = dict(source.fields)
    for pair in pairs:
        both = as_list(pair)
        name = as_text(both[0])
        if name in fields:
            fields[name] = _call(both[1], [fields[name]])
    return Record(fields)


@m("Value.Is", 2, 2)
def m_value_is(value: object, wanted: object) -> object:
    from pyopenvba.mlang._eval import is_type

    return is_type(value, wanted.name if isinstance(wanted, MType) else as_text(wanted))


@m("Type.Is", 2, 2)
def m_type_is(value: object, wanted: object) -> object:
    return m_value_is(value, wanted)
