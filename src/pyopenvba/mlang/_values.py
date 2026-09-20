"""What Power Query's M computes with.

M has its own set of values, and the two that carry the work are the
record and the table.  A record is an ordered set of named fields; a
table is a list of records that all share their field names, which is
why a column can be added to a whole table at once.

Nothing here is an Excel value or a VBA value.  A refresh converts at
the boundary, where the result lands on a sheet.
"""

from __future__ import annotations

import datetime as _dt
import decimal as _dec
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Callable, Iterator, TypeGuard

if TYPE_CHECKING:
    from pyopenvba.mlang._eval import Scope
    from pyopenvba.mlang._parse import Node

from pyopenvba.exceptions import PyOpenVBAError


class MError(PyOpenVBAError):
    """An error raised while evaluating M, carrying M's own record.

    M errors are values: ``try`` catches one and hands back a record
    with Reason, Message and Detail, so the error travels with those
    three fields rather than as a bare message.
    """

    def __init__(self, reason: str, message: str = "", detail: object = None) -> None:
        super().__init__(f"{reason}: {message}" if message else reason)
        self.reason = reason
        self.message = message
        self.detail = detail

    def as_record(self) -> Record:
        return Record({"Reason": self.reason, "Message": self.message, "Detail": self.detail})


def expression_error(message: str, detail: object = None) -> MError:
    return MError("Expression.Error", message, detail)


#: M's null.  Python's None is used for it everywhere except where a
#: missing value and a null have to be told apart.
NULL = None


@dataclass(slots=True)
class Record:
    """An ordered set of named fields."""

    fields: dict[str, object] = field(default_factory=lambda: {})

    def __getitem__(self, name: str) -> object:
        if name not in self.fields:
            raise MError("Expression.Error", f"The field '{name}' of the record wasn't found.")
        return self.fields[name]

    def get(self, name: str, default: object = None) -> object:
        return self.fields.get(name, default)

    def has(self, name: str) -> bool:
        return name in self.fields

    @property
    def names(self) -> list[str]:
        return list(self.fields)

    def __len__(self) -> int:
        return len(self.fields)

    def __repr__(self) -> str:
        inner = ", ".join(f"{name} = {short(value)}" for name, value in self.fields.items())
        return f"[{inner}]"


@dataclass(slots=True)
class Table:
    """Rows of values under named columns."""

    columns: list[str] = field(default_factory=lambda: [])
    rows: list[list[object]] = field(default_factory=lambda: [])

    @classmethod
    def from_records(cls, records: list[Record], columns: list[str] | None = None) -> Table:
        names = list(columns) if columns is not None else []
        if not names:
            for record in records:
                for name in record.names:
                    if name not in names:
                        names.append(name)
        rows = [[record.get(name) for name in names] for record in records]
        return cls(names, rows)

    def record_at(self, index: int) -> Record:
        if not 0 <= index < len(self.rows):
            raise MError("Expression.Error", "There weren't enough elements in the enumeration.")
        return Record(dict(zip(self.columns, self.rows[index])))

    def records(self) -> Iterator[Record]:
        for row in self.rows:
            yield Record(dict(zip(self.columns, row)))

    def column_index(self, name: str) -> int:
        try:
            return self.columns.index(name)
        except ValueError:
            raise MError(
                "Expression.Error", f"The column '{name}' of the table wasn't found."
            ) from None

    def column(self, name: str) -> list[object]:
        at = self.column_index(name)
        return [row[at] for row in self.rows]

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return len(self.columns)

    def __repr__(self) -> str:
        return f"Table({self.width} columns x {self.height} rows)"


@dataclass(slots=True)
class Function:
    """A function value: its parameters, its body, and where it was written."""

    parameters: list[Parameter]
    body: Node | None
    closure: Scope | None
    name: str = ""

    def __repr__(self) -> str:
        return f"Function({', '.join(one.name for one in self.parameters)})"


@dataclass(slots=True)
class Parameter:
    name: str
    optional: bool = False
    declared: str = ""


@dataclass(slots=True)
class Builtin:
    """A library function, with the arity M would report."""

    name: str
    call: Callable[..., object]
    minimum: int = 0
    maximum: int | None = None

    def __repr__(self) -> str:
        return f"Function({self.name})"


@dataclass(slots=True)
class MType:
    """A type value: what ``type text`` and ``Int64.Type`` come to."""

    name: str
    detail: object = None

    def __repr__(self) -> str:
        return f"type {self.name}"


@dataclass(slots=True)
class Duration:
    """M's duration: days, hours, minutes and seconds.

    Written the way it was given but held normalized, because
    ``#duration(0, 0, 0, 90)`` is a minute and a half to the engine and
    writes itself as ``00:01:30``.
    """

    days: float = 0.0
    hours: float = 0.0
    minutes: float = 0.0
    seconds: float = 0.0

    def __post_init__(self) -> None:
        whole = self.total_seconds
        sign = -1.0 if whole < 0 else 1.0
        rest = abs(whole)
        days, rest = divmod(rest, 86400.0)
        hours, rest = divmod(rest, 3600.0)
        minutes, seconds = divmod(rest, 60.0)
        self.days, self.hours = sign * days, sign * hours
        self.minutes, self.seconds = sign * minutes, sign * seconds

    @property
    def total_seconds(self) -> float:
        return self.days * 86400 + self.hours * 3600 + self.minutes * 60 + self.seconds

    def __repr__(self) -> str:
        return f"#duration({int(self.days)}, {int(self.hours)}, {int(self.minutes)}, {self.seconds:g})"


def duration_text(value: Duration) -> str:
    """A duration as ``Duration.ToText`` writes it.

    The day count is only there when there is one: three hours is
    ``03:00:00``, and a day and two hours is ``1.02:00:00``.
    """
    whole = value.total_seconds
    sign = "-" if whole < 0 else ""
    rest = abs(whole)
    days, rest = divmod(rest, 86400.0)
    hours, rest = divmod(rest, 3600.0)
    minutes, seconds = divmod(rest, 60.0)
    clock = f"{int(hours):02d}:{int(minutes):02d}:{seconds:02.0f}"
    return f"{sign}{int(days)}.{clock}" if days else f"{sign}{clock}"


def is_list(value: object) -> TypeGuard[list[Any]]:
    """Whether the value is an M list.

    A type guard rather than a bare isinstance: narrowing an object to
    list leaves the element type unknown, and every call that passes
    the narrowed list on then carries that unknown with it.
    """
    return isinstance(value, list)


def is_mapping(value: object) -> TypeGuard[dict[Any, Any]]:
    """Whether the value is a mapping, for the same reason as is_list."""
    return isinstance(value, dict)


def short(value: object) -> str:
    """One value, short enough to sit inside another's repr."""
    if isinstance(value, str):
        return f'"{value}"'
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    return repr(value)


# --- what a value is ---------------------------------------------------------------


def type_name(value: object) -> str:
    """The name M gives a value's type."""
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "logical"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str):
        return "text"
    if is_list(value):
        return "list"
    if isinstance(value, Record):
        return "record"
    if isinstance(value, Table):
        return "table"
    if isinstance(value, (Function, Builtin)):
        return "function"
    if isinstance(value, MType):
        return "type"
    if isinstance(value, bytes):
        return "binary"
    if isinstance(value, Duration):
        return "duration"
    if isinstance(value, _dt.datetime):
        return "datetime"
    if isinstance(value, _dt.date):
        return "date"
    if isinstance(value, _dt.time):
        return "time"
    return "any"


def as_number(value: object, *, where: str = "") -> float:
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        raise MError("Expression.Error", "We cannot convert the value null to type Number.")
    raise MError(
        "Expression.Error",
        f"We cannot convert a value of type {type_name(value)} to type Number.{where}",
    )


def as_text(value: object) -> str:
    if isinstance(value, str):
        return value
    if value is None:
        raise MError("Expression.Error", "We cannot convert the value null to type Text.")
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, Scaled) and value.places:
        return f"{float(value):.{value.places}f}"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (int, float)):
        return repr(value)
    if isinstance(value, _dt.datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, _dt.date):
        return value.isoformat()
    raise MError(
        "Expression.Error", f"We cannot convert a value of type {type_name(value)} to type Text."
    )


def as_logical(value: object) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        raise MError("Expression.Error", "We cannot convert the value null to type Logical.")
    raise MError(
        "Expression.Error",
        f"We cannot convert a value of type {type_name(value)} to type Logical.",
    )


def as_list(value: object) -> list[Any]:
    if is_list(value):
        return value
    raise MError(
        "Expression.Error", f"We cannot convert a value of type {type_name(value)} to type List."
    )


def as_table(value: object) -> Table:
    if isinstance(value, Table):
        return value
    raise MError(
        "Expression.Error", f"We cannot convert a value of type {type_name(value)} to type Table."
    )


def as_record(value: object) -> Record:
    if isinstance(value, Record):
        return value
    raise MError(
        "Expression.Error", f"We cannot convert a value of type {type_name(value)} to type Record."
    )


def number_out(value: float) -> object:
    """A number as M hands it back, whole where it is whole."""
    if isinstance(value, bool):
        return value
    if isinstance(value, Scaled):
        return value
    if float(value).is_integer() and abs(value) < 1e15:
        return int(value)
    return value


class Scaled(float):
    """A number that remembers the decimal places it was rounded at.

    ``Number.Round(2.345, 2)`` answers 2.34, and Power Query writes that
    as "2.340": the engine rounds in decimal and hands back a value that
    still carries the places the number arrived with.  Carrying them on
    the float itself means every other operator goes on seeing a plain
    number, and only the writing out has to know.
    """

    __slots__ = ("places",)

    places: int

    def __new__(cls, value: float, places: int) -> Scaled:
        # Rounding -0.5 to even gives a negative zero, which M writes
        # as 0 rather than -0.
        out = super().__new__(cls, 0.0 if value == 0 else value)
        out.places = max(places, 0)
        return out


def places_of(value: object) -> int:
    """How many decimal places a number was written with."""
    if isinstance(value, Scaled):
        return value.places
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return 0
    exponent = _dec.Decimal(repr(float(value))).normalize().as_tuple().exponent
    return -exponent if isinstance(exponent, int) and exponent < 0 else 0
