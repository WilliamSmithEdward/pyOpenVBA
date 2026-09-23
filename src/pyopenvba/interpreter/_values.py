"""VBA's value model: what a Variant holds, and how VBA combines two of them.

A Variant is represented by an ordinary Python object wherever Python has
one that behaves the same, and by a class here wherever it does not:

    Empty            EMPTY              an unset variable
    Null             NULL               the database "no value"
    Nothing          NOTHING            an object variable pointing nowhere
    Missing          MISSING            an Optional argument nobody passed
    Boolean          bool
    Byte             VBAInt(kind="Byte")
    Integer          VBAInt(kind="Integer")
    Long             VBAInt(kind="Long"), or a plain int
    LongLong         VBAInt(kind="LongLong")
    Single           VBASingle
    Double           float
    Currency         VBACurrency
    Decimal          VBADecimal
    Date             VBADate
    String           str
    Object           a VBAObject, or NOTHING
    Error            VBAErrorValue
    array            VBAArray

The integer width is carried rather than inferred because VBA's arithmetic
depends on it: ``30000 * 30000`` is an overflow, not 900000000, since both
operands are Integers and VBA does not widen the result.  That gotcha is
one of the reasons to run a macro here at all, so the model keeps it.

Every conversion and operator in this module raises
:class:`~pyopenvba.exceptions.VBARuntimeError` with the number VBA uses,
so ``On Error`` traps them exactly as it would in Office.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
import struct
from decimal import ROUND_HALF_EVEN, Decimal, InvalidOperation, localcontext
from typing import Any, Final

from pyopenvba.exceptions import VBARuntimeError

# --- error numbers ---------------------------------------------------------------

ERR_RETURN_WITHOUT_GOSUB: Final = 3
ERR_INVALID_PROCEDURE_CALL: Final = 5
ERR_OVERFLOW: Final = 6
ERR_OUT_OF_MEMORY: Final = 7
ERR_SUBSCRIPT_OUT_OF_RANGE: Final = 9
ERR_ARRAY_FIXED: Final = 10
ERR_DIVISION_BY_ZERO: Final = 11
ERR_TYPE_MISMATCH: Final = 13
ERR_OUT_OF_STRING_SPACE: Final = 14
ERR_BAD_FILE_NAME_OR_NUMBER: Final = 52
ERR_FILE_NOT_FOUND: Final = 53
ERR_DEVICE_IO: Final = 57
ERR_PERMISSION_DENIED: Final = 70
ERR_PATH_NOT_FOUND: Final = 76
ERR_OBJECT_VARIABLE_NOT_SET: Final = 91
ERR_INVALID_USE_OF_NULL: Final = 94
ERR_CANNOT_CREATE_OBJECT: Final = 429
ERR_MEMBER_NOT_FOUND: Final = 438
ERR_ARGUMENT_NOT_OPTIONAL: Final = 449
ERR_NOT_A_COLLECTION: Final = 451
ERR_APPLICATION_DEFINED: Final = 1004

#: The text the VBA IDE shows beside each of the error numbers it raises.
ERROR_TEXT: Final[dict[int, str]] = {
    3: "Return without GoSub",
    5: "Invalid procedure call or argument",
    6: "Overflow",
    7: "Out of memory",
    9: "Subscript out of range",
    10: "This array is fixed or temporarily locked",
    11: "Division by zero",
    13: "Type mismatch",
    14: "Out of string space",
    16: "Expression too complex",
    17: "Can't perform requested operation",
    18: "User interrupt occurred",
    20: "Resume without error",
    28: "Out of stack space",
    35: "Sub or Function not defined",
    48: "Error in loading DLL",
    49: "Bad DLL calling convention",
    51: "Internal error",
    52: "Bad file name or number",
    53: "File not found",
    54: "Bad file mode",
    55: "File already open",
    57: "Device I/O error",
    58: "File already exists",
    59: "Bad record length",
    61: "Disk full",
    62: "Input past end of file",
    63: "Bad record number",
    67: "Too many files",
    68: "Device unavailable",
    70: "Permission denied",
    71: "Disk not ready",
    74: "Can't rename with different drive",
    75: "Path/File access error",
    76: "Path not found",
    91: "Object variable or With block variable not set",
    92: "For loop not initialized",
    93: "Invalid pattern string",
    94: "Invalid use of Null",
    322: "Can't create necessary temporary file",
    424: "Object required",
    429: "ActiveX component can't create object",
    430: "Class doesn't support Automation",
    432: "File name or class name not found during Automation operation",
    438: "Object doesn't support this property or method",
    440: "Automation error",
    442: "Connection to type library or object library for remote process has been lost",
    443: "Automation object does not have a default value",
    445: "Object doesn't support this action",
    446: "Object doesn't support named arguments",
    447: "Object doesn't support current locale setting",
    448: "Named argument not found",
    449: "Argument not optional",
    450: "Wrong number of arguments or invalid property assignment",
    451: "Property let procedure not defined and property get procedure did not return an object",
    452: "Invalid ordinal",
    453: "Specified DLL function not found",
    454: "Code resource not found",
    455: "Code resource lock error",
    457: "This key is already associated with an element of this collection",
    458: "Variable uses an Automation type not supported in Visual Basic",
    459: "Object or class does not support the set of events",
    460: "Invalid Clipboard format",
    461: "Method or data member not found",
    462: "The remote server machine does not exist or is unavailable",
    463: "Class not registered on local machine",
    1004: "Application-defined or object-defined error",
}


def error(number: int, description: str = "", *, source: str = "", where: str = "") -> VBARuntimeError:
    """The :class:`VBARuntimeError` VBA raises for ``number``."""
    return VBARuntimeError(
        number,
        description or ERROR_TEXT.get(number, "Application-defined or object-defined error"),
        source=source,
        where=where,
    )


# --- the values with no Python equivalent -----------------------------------------


class _Singleton:
    """A value that is only ever itself: Empty, Null, Nothing, Missing."""

    __slots__ = ("_name",)

    def __init__(self, name: str) -> None:
        self._name = name

    def __repr__(self) -> str:
        return self._name

    def __bool__(self) -> bool:
        return False


EMPTY: Final = _Singleton("Empty")
NULL: Final = _Singleton("Null")
NOTHING: Final = _Singleton("Nothing")
MISSING: Final = _Singleton("Missing")

#: Print's two separators, which travel with its arguments: a comma
#: moves to the next 14-column zone and a trailing semicolon holds the
#: line open for the next Print.
PRINT_ZONE: Final = _Singleton("PrintZone")
PRINT_CONTINUE: Final = _Singleton("PrintContinue")


class VBAInt(int):
    """An integer that remembers how wide VBA declared it.

    No ``__slots__``: CPython refuses them on an int subclass, and the
    width has to live somewhere.
    """

    kind: str

    def __new__(cls, value: int, kind: str = "Long") -> VBAInt:
        made = super().__new__(cls, value)
        made.kind = kind
        return made

    def __repr__(self) -> str:
        return f"{self.kind}({int(self)})"

    def __str__(self) -> str:
        # int's own, because str() falls back to __repr__ otherwise and
        # every Decimal(str(value)) in the codebase would see "Long(1)".
        return int.__repr__(self)


class VBASingle(float):
    """A Double narrowed to the 32-bit precision VBA stores in a Single."""

    __slots__ = ()

    def __new__(cls, value: float) -> VBASingle:
        try:
            narrowed = struct.unpack("<f", struct.pack("<f", value))[0]
        except OverflowError:
            raise error(ERR_OVERFLOW) from None
        return super().__new__(cls, narrowed)


class VBACurrency(Decimal):
    """Currency: a 64-bit integer of ten-thousandths, so always 4 decimals."""

    __slots__ = ()

    def __new__(cls, value: object) -> VBACurrency:
        try:
            scaled = Decimal(str(value)).quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
        except (InvalidOperation, ValueError, ArithmeticError):
            raise error(ERR_TYPE_MISMATCH) from None
        if not -922337203685477.5808 <= scaled <= 922337203685477.5807:
            raise error(ERR_OVERFLOW)
        return super().__new__(cls, scaled)


class VBADecimal(Decimal):
    """The Variant-only Decimal subtype: 28 significant digits, no more."""

    __slots__ = ()


class VBADate:
    """A VBA Date: a Double whose whole part is days from 1899-12-30."""

    __slots__ = ("serial",)

    #: Day zero.  A Date of 0 is midnight on it, and 1 is 1899-12-31.
    EPOCH: Final = _dt.datetime(1899, 12, 30)

    def __init__(self, serial: float) -> None:
        if not -657434.0 <= serial < 2958466.0:
            raise error(ERR_OVERFLOW)
        self.serial = float(serial)

    @classmethod
    def from_datetime(cls, when: _dt.datetime | _dt.date | _dt.time) -> VBADate:
        if isinstance(when, _dt.datetime):
            delta = when - cls.EPOCH
            return cls(delta.days + delta.seconds / 86400.0 + delta.microseconds / 86400e6)
        if isinstance(when, _dt.date):
            return cls((when - cls.EPOCH.date()).days)
        seconds = when.hour * 3600 + when.minute * 60 + when.second
        return cls(seconds / 86400.0 + when.microsecond / 86400e6)

    def to_datetime(self) -> _dt.datetime:
        """The moment this serial names, to the second VBA stores."""
        days = math.floor(self.serial)
        fraction = self.serial - days
        if self.serial < 0:
            # A negative serial counts the day backwards but the time forwards.
            days = -math.floor(-self.serial)
            fraction = -self.serial - math.floor(-self.serial)
        seconds = round(fraction * 86400.0)
        if seconds >= 86400:
            days, seconds = days + 1, 0
        return self.EPOCH + _dt.timedelta(days=days, seconds=seconds)

    def __repr__(self) -> str:
        return f"VBADate({self.serial!r})"

    def __eq__(self, other: object) -> bool:
        return isinstance(other, VBADate) and other.serial == self.serial

    def __hash__(self) -> int:
        return hash(("VBADate", self.serial))


class VBAErrorValue:
    """The value ``CVErr`` makes, held in a Variant until someone reads it."""

    __slots__ = ("number",)

    def __init__(self, number: int) -> None:
        self.number = number

    def __repr__(self) -> str:
        return f"Error {self.number}"

    def __eq__(self, other: object) -> bool:
        return isinstance(other, VBAErrorValue) and other.number == self.number

    def __hash__(self) -> int:
        return hash(("VBAErrorValue", self.number))


class VBAArray:
    """A VBA array: bounds per dimension, and a flat store in VBA's order.

    VBA stores an array column-major, which matters only where somebody
    reads it as a block (a Range's Value, say), so the flat index is
    computed the way VBA lays it out rather than the way Python would.
    """

    __slots__ = ("bounds", "items", "element_type", "fixed")

    def __init__(
        self,
        bounds: list[tuple[int, int]],
        *,
        element_type: str = "Variant",
        fixed: bool = False,
        items: list[object] | None = None,
    ) -> None:
        for lower, upper in bounds:
            if upper < lower - 1:
                raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        self.bounds = bounds
        self.element_type = element_type
        self.fixed = fixed
        self.items = items if items is not None else [default_for(element_type) for _ in range(self.size)]

    @property
    def size(self) -> int:
        total = 1
        for lower, upper in self.bounds:
            total *= max(0, upper - lower + 1)
        return total

    @property
    def dimensions(self) -> int:
        return len(self.bounds)

    def offset(self, subscripts: list[int]) -> int:
        if len(subscripts) != len(self.bounds):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        index = 0
        stride = 1
        for value, (lower, upper) in zip(subscripts, self.bounds):
            if value < lower or value > upper:
                raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
            index += (value - lower) * stride
            stride *= upper - lower + 1
        return index

    def get(self, subscripts: list[int]) -> object:
        return self.items[self.offset(subscripts)]

    def set(self, subscripts: list[int], value: object) -> None:
        self.items[self.offset(subscripts)] = coerce(value, self.element_type)

    def resized(self, bounds: list[tuple[int, int]], *, preserve: bool) -> VBAArray:
        """A new array of ``bounds``, carrying the old contents if asked.

        VBA's Preserve keeps elements by subscript, and only the last
        dimension may change size; anything else is error 9.
        """
        fresh = VBAArray(bounds, element_type=self.element_type)
        if not preserve:
            return fresh
        if len(bounds) != len(self.bounds):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        if any(old != new for old, new in zip(self.bounds[:-1], bounds[:-1])):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        for flat in range(self.size):
            subscripts = self._subscripts(flat)
            if all(lower <= value <= upper for value, (lower, upper) in zip(subscripts, bounds)):
                fresh.items[fresh.offset(subscripts)] = self.items[flat]
        return fresh

    def _subscripts(self, flat: int) -> list[int]:
        out: list[int] = []
        for lower, upper in self.bounds:
            width = upper - lower + 1
            out.append(lower + flat % width)
            flat //= width
        return out

    def elements(self) -> list[object]:
        """Every element, in the order ``For Each`` walks them."""
        return list(self.items)

    def __repr__(self) -> str:
        shape = ", ".join(f"{lower} To {upper}" for lower, upper in self.bounds)
        return f"VBAArray({shape})"


# --- type names -------------------------------------------------------------------

#: What VarType() answers for each of the types.
VAR_TYPES: Final[dict[str, int]] = {
    "Empty": 0,
    "Null": 1,
    "Integer": 2,
    "Long": 3,
    "Single": 4,
    "Double": 5,
    "Currency": 6,
    "Date": 7,
    "String": 8,
    "Object": 9,
    "Error": 10,
    "Boolean": 11,
    "Variant": 12,
    "Decimal": 14,
    "Byte": 17,
    "LongLong": 20,
    "Nothing": 9,
}

NUMERIC_TYPES: Final = frozenset(
    {"Byte", "Integer", "Long", "LongLong", "Single", "Double", "Currency", "Decimal"}
)

#: The inclusive range each integer width holds.
INTEGER_LIMITS: Final[dict[str, tuple[int, int]]] = {
    "Byte": (0, 255),
    "Integer": (-32768, 32767),
    "Long": (-2147483648, 2147483647),
    "LongLong": (-9223372036854775808, 9223372036854775807),
}

_WIDENING: Final = ["Byte", "Integer", "Long", "LongLong", "Single", "Currency", "Decimal", "Double"]

#: How many bytes a declared type takes, which is what Len answers for
#: anything that is not a String.
STORAGE_WIDTH: Final[dict[str, int]] = {
    "Byte": 1,
    "Boolean": 2,
    "Integer": 2,
    "Long": 4,
    "LongLong": 8,
    "Single": 4,
    "Double": 8,
    "Currency": 8,
    "Date": 8,
    "Decimal": 14,
}


def type_name(value: object) -> str:
    """The name ``TypeName`` gives ``value``."""
    if value is EMPTY:
        return "Empty"
    if value is NULL:
        return "Null"
    if value is NOTHING:
        return "Nothing"
    if value is MISSING:
        return "Error"
    if isinstance(value, bool):
        return "Boolean"
    if isinstance(value, VBAInt):
        return value.kind
    if isinstance(value, int):
        return "Long"
    if isinstance(value, VBASingle):
        return "Single"
    if isinstance(value, float):
        return "Double"
    if isinstance(value, VBACurrency):
        return "Currency"
    if isinstance(value, Decimal):
        return "Decimal"
    if isinstance(value, VBADate):
        return "Date"
    if isinstance(value, str):
        return "String"
    if isinstance(value, VBAErrorValue):
        return "Error"
    if isinstance(value, VBAArray):
        return f"{value.element_type}()"
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(value, VBAObject):
        return value.vba_type_name
    return "Variant"


def default_for(declared: str) -> object:
    """What a variable of ``declared`` holds before anything is assigned."""
    base = declared.rstrip("()")
    if base in ("Variant", ""):
        return EMPTY
    if base == "String":
        return ""
    if base == "Boolean":
        return False
    if base == "Date":
        return VBADate(0.0)
    if base == "Object" or base not in VAR_TYPES:
        return NOTHING
    if base in ("Single", "Double"):
        return VBASingle(0.0) if base == "Single" else 0.0
    if base == "Currency":
        return VBACurrency(0)
    if base == "Decimal":
        return VBADecimal(0)
    return VBAInt(0, base)


def is_object(value: object) -> bool:
    from pyopenvba.interpreter._objects import VBAObject

    return isinstance(value, VBAObject) or value is NOTHING


# --- conversions ------------------------------------------------------------------

_NUMBER_TEXT = re.compile(
    r"""^[ \t]*(?P<sign>[-+]?)[ \t]*
        (?: &[hH](?P<hex>[0-9A-Fa-f]+)
          | &[oO](?P<octal>[0-7]+)
          | (?P<plain>(?:\d+\.?\d*|\.\d+)(?:[eEdD][-+]?\d+)?)
        )[ \t]*$""",
    re.VERBOSE,
)


def to_number(value: object, *, context: str = "") -> int | float | Decimal:
    """``value`` as a number, the way VBA reads one from a Variant."""
    if isinstance(value, bool):
        return VBAInt(-1 if value else 0, "Integer")
    if isinstance(value, (VBAInt, int)):
        return value
    if isinstance(value, (float, Decimal)):
        return value
    if value is EMPTY or value is MISSING:
        return VBAInt(0, "Integer")
    if value is NULL:
        raise error(ERR_INVALID_USE_OF_NULL)
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, str):
        return text_to_number(value)
    if isinstance(value, VBAErrorValue):
        raise error(ERR_TYPE_MISMATCH)
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(value, VBAObject):
        return to_number(value.vba_value(), context=context)
    raise error(ERR_TYPE_MISMATCH, f"cannot read {type_name(value)} as a number{context}")


def text_to_number(text: str) -> int | float:
    """A string as VBA reads it in an arithmetic context.

    Stricter than ``Val``: the whole string has to be a number, or VBA
    raises 13 rather than taking the leading digits.
    """
    match = _NUMBER_TEXT.match(text)
    if not match:
        raise error(ERR_TYPE_MISMATCH)
    sign = -1 if match.group("sign") == "-" else 1
    if match.group("hex") is not None:
        return VBAInt(sign * int(match.group("hex"), 16), "Long")
    if match.group("octal") is not None:
        return VBAInt(sign * int(match.group("octal"), 8), "Long")
    body = match.group("plain").lower().replace("d", "e")
    if "." not in body and "e" not in body:
        return VBAInt(sign * int(body), "Long" if abs(int(body)) > 32767 else "Integer")
    return sign * float(body)


def to_bool(value: object) -> bool:
    """``value`` as a condition: anything non-zero is True."""
    if isinstance(value, bool):
        return value
    if value is NULL:
        raise error(ERR_INVALID_USE_OF_NULL)
    if value is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET)
    if isinstance(value, str) and value.strip().lower() in ("true", "false"):
        # CBool("True") is True in VBA: the words are read as well as
        # the numbers.
        return value.strip().lower() == "true"
    return to_number(value) != 0


def number_text(value: int | float | Decimal) -> str:
    """A number as VBA renders it into a string.

    The rule, measured against Excel rather than recalled: round to the
    type's significant digits (15 for a Double, 7 for a Single), then
    write it out in full if every digit fits within that many character
    positions, counting from the first digit written to the last.
    Otherwise use an exponent.  So 999999999999999 is written in full
    and 1E+15 is not, 0.000000000000001 is and 1E-16 is not, and
    1.23456789012345E-11 is not although 1E-11 is.
    """
    if isinstance(value, bool):
        return "True" if value else "False"
    if isinstance(value, int):
        return str(int(value))
    if isinstance(value, Decimal):
        plain = format(value.normalize(), "f")
        return plain if plain != "-0" else "0"
    if value != value:  # NaN
        return "-1.#IND"
    if value in (float("inf"), float("-inf")):
        return "1.#INF" if value > 0 else "-1.#INF"
    if value == 0.0:
        # VBA keeps the sign: Round(-0.5) prints -0, not 0.
        return "-0" if math.copysign(1.0, value) < 0 else "0"
    precision = 7 if isinstance(value, VBASingle) else 15
    with localcontext() as context:
        context.prec = precision
        rounded = (+Decimal(float(value))).normalize()
    sign, digits, exponent = rounded.as_tuple()
    if not isinstance(exponent, int):  # pragma: no cover - only for NaN
        return str(rounded)
    count = len(digits)
    highest = exponent + count - 1
    positions = max(highest + 1, count) if highest >= 0 else count - highest - 1
    if positions <= precision:
        return format(rounded, "f")
    mantissa = "".join(str(digit) for digit in digits).rstrip("0") or "0"
    body = mantissa[0] + ("." + mantissa[1:] if len(mantissa) > 1 else "")
    return f"{'-' if sign else ''}{body}E{'+' if highest >= 0 else '-'}{abs(highest):02d}"


def to_text(value: object, *, context: str = "") -> str:
    """``value`` as a String, the way an implicit conversion makes one."""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "True" if value else "False"
    if value is EMPTY or value is MISSING:
        return ""
    if value is NULL:
        raise error(ERR_INVALID_USE_OF_NULL)
    if value is NOTHING:
        raise error(ERR_OBJECT_VARIABLE_NOT_SET)
    if isinstance(value, VBADate):
        return date_text(value)
    if isinstance(value, (int, float, Decimal)):
        return number_text(value)
    if isinstance(value, VBAErrorValue):
        return f"Error {value.number}"
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(value, VBAObject):
        return to_text(value.vba_value(), context=context)
    raise error(ERR_TYPE_MISMATCH, f"cannot read {type_name(value)} as a string{context}")


def date_text(value: VBADate) -> str:
    """A Date as VBA's default conversion prints it.

    Midnight prints as the date alone and a serial under a day as the
    time alone, which is why a cell holding 0.5 shows only a clock.
    """
    when = value.to_datetime()
    if value.serial == math.floor(value.serial):
        return f"{when.month}/{when.day}/{when.year}"
    if -1 < value.serial < 1:
        return _time_text(when)
    return f"{when.month}/{when.day}/{when.year} {_time_text(when)}"


def _time_text(when: _dt.datetime) -> str:
    hour = when.hour % 12 or 12
    suffix = "AM" if when.hour < 12 else "PM"
    return f"{hour}:{when.minute:02d}:{when.second:02d} {suffix}"


_DATE_PATTERNS: Final = (
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %I:%M:%S %p",
    "%m/%d/%Y %I:%M %p",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
    "%m/%d/%y",
    "%m-%d-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%B %d, %Y",
    "%d %B %Y",
    "%b %d, %Y",
    "%H:%M:%S",
    "%I:%M:%S %p",
    "%I:%M %p",
    "%H:%M",
)


def parse_date_text(text: str) -> VBADate | None:
    """``text`` as a Date, or None where VBA would not read one.

    The list is the US order Office's own date literals use; a locale's
    day-first order is a separate question this does not answer.
    """
    stripped = text.strip()
    if not stripped:
        return None
    for pattern in _DATE_PATTERNS:
        try:
            when = _dt.datetime.strptime(stripped, pattern)
        except ValueError:
            continue
        if "%Y" not in pattern and "%y" not in pattern:
            return VBADate.from_datetime(when.time())
        return VBADate.from_datetime(when)
    return None


def to_date(value: object) -> VBADate:
    """``value`` as a Date, as CDate reads one."""
    if isinstance(value, VBADate):
        return value
    if value is NULL:
        raise error(ERR_INVALID_USE_OF_NULL)
    if value is EMPTY:
        return VBADate(0.0)
    if isinstance(value, str):
        parsed = parse_date_text(value)
        if parsed is None:
            raise error(ERR_TYPE_MISMATCH, f"{value!r} is not a date")
        return parsed
    number = to_number(value)
    return VBADate(float(number))


def round_half_even(value: float | Decimal, places: int = 0) -> float | Decimal:
    """VBA rounds half to even, so 0.5 goes to 0 and 1.5 goes to 2."""
    if isinstance(value, Decimal):
        with localcontext() as context:
            context.prec = 40
            return value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_EVEN)
    if value != value or value in (float("inf"), float("-inf")):
        raise error(ERR_OVERFLOW)
    try:
        quantized = Decimal(repr(value)).quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_EVEN)
    except (InvalidOperation, ArithmeticError):
        raise error(ERR_OVERFLOW) from None
    return float(quantized)


def to_integer(value: object, kind: str = "Long") -> VBAInt:
    """``value`` as one of the integer widths, rounding half to even."""
    number = to_number(value)
    if isinstance(number, int) and not isinstance(number, bool):
        whole = int(number)
    else:
        whole = int(round_half_even(number))
    low, high = INTEGER_LIMITS[kind]
    if not low <= whole <= high:
        raise error(ERR_OVERFLOW)
    return VBAInt(whole, kind)


def coerce(value: object, declared: str) -> object:
    """``value`` stored in a variable declared ``As declared``.

    This is where a typed variable does its work: assigning 1.7 to an
    Integer stores 2, and assigning 40000 raises 6 rather than widening.
    """
    base = declared.rstrip("()")
    if base in ("Variant", "", "Any"):
        return value
    if value is NULL:
        if base == "String":
            raise error(ERR_INVALID_USE_OF_NULL)
        raise error(ERR_INVALID_USE_OF_NULL)
    if base == "String":
        return to_text(value)
    if base == "Boolean":
        return to_bool(value) if not isinstance(value, bool) else value
    if base == "Date":
        return to_date(value)
    if base in INTEGER_LIMITS:
        return to_integer(value, base)
    if base == "Single":
        return VBASingle(float(to_number(value)))
    if base == "Double":
        return float(to_number(value))
    if base == "Currency":
        return VBACurrency(to_number(value))
    if base == "Decimal":
        return VBADecimal(str(to_number(value)))
    if base == "Object" or base not in VAR_TYPES:
        if not is_object(value):
            raise error(ERR_TYPE_MISMATCH, f"cannot store {type_name(value)} in a {base}")
        return value
    return value


# --- operators --------------------------------------------------------------------


def _wider(left: object, right: object) -> str:
    """The type VBA computes in when the two operands differ."""
    kinds: list[str] = []
    for value in (left, right):
        name = type_name(value)
        if name in ("Boolean", "Empty", "Error"):
            name = "Integer"
        elif name == "Date":
            name = "Double"
        elif name == "String":
            name = "Double"
        kinds.append(name)
    return max(kinds, key=lambda name: _WIDENING.index(name) if name in _WIDENING else len(_WIDENING))


def _as_kind(value: int | float | Decimal, kind: str) -> object:
    if kind in INTEGER_LIMITS:
        low, high = INTEGER_LIMITS[kind]
        if not low <= value <= high:
            raise error(ERR_OVERFLOW)
        return VBAInt(int(value), kind)
    if kind == "Single":
        return VBASingle(float(value))
    if kind == "Currency":
        return VBACurrency(value)
    if kind == "Decimal":
        return VBADecimal(str(value))
    result = float(value)
    if result in (float("inf"), float("-inf")):
        raise error(ERR_OVERFLOW)
    return result


def _both_numbers(left: object, right: object) -> tuple[Any, Any, str]:
    kind = _wider(left, right)
    first = to_number(left)
    second = to_number(right)
    if kind in ("Currency", "Decimal"):
        return Decimal(str(first)), Decimal(str(second)), kind
    if kind in INTEGER_LIMITS:
        return int(first), int(second), kind
    return float(first), float(second), kind


def add(left: object, right: object) -> object:
    """``+``, which concatenates only when both sides are already strings."""
    if left is NULL or right is NULL:
        return NULL
    if isinstance(left, str) and isinstance(right, str):
        return left + right
    if left is EMPTY and isinstance(right, str):
        return right
    if right is EMPTY and isinstance(left, str):
        return left
    if isinstance(left, VBADate) or isinstance(right, VBADate):
        return VBADate(float(to_number(left)) + float(to_number(right)))
    first, second, kind = _both_numbers(left, right)
    return _as_kind(first + second, kind)


def subtract(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    if isinstance(left, VBADate) and isinstance(right, VBADate):
        return float(left.serial) - float(right.serial)
    if isinstance(left, VBADate):
        return VBADate(left.serial - float(to_number(right)))
    first, second, kind = _both_numbers(left, right)
    return _as_kind(first - second, kind)


def multiply(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    first, second, kind = _both_numbers(left, right)
    return _as_kind(first * second, kind)


def divide(left: object, right: object) -> object:
    """``/``, which is always a floating division even between Integers."""
    if left is NULL or right is NULL:
        return NULL
    first, second, kind = _both_numbers(left, right)
    if second == 0:
        raise error(ERR_DIVISION_BY_ZERO)
    if kind in ("Currency", "Decimal"):
        with localcontext() as context:
            context.prec = 30
            return _as_kind(Decimal(first) / Decimal(second), kind)
    return float(first) / float(second)


def _whole_width(left: object, right: object) -> str:
    """The width ``\\``, Mod and the bitwise operators come out in.

    Both sides are rounded to a whole number first; the result is an
    Integer only where both sides already were one.
    """
    kinds = {type_name(left), type_name(right)}
    if kinds <= {"Byte", "Integer", "Boolean", "Empty"}:
        return "Integer"
    if "LongLong" in kinds:
        return "LongLong"
    return "Long"


def int_divide(left: object, right: object) -> object:
    """``\\``, which rounds both sides to a whole number and truncates after."""
    if left is NULL or right is NULL:
        return NULL
    kind = _whole_width(left, right)
    first = to_integer(left, "LongLong")
    second = to_integer(right, "LongLong")
    if second == 0:
        raise error(ERR_DIVISION_BY_ZERO)
    quotient = abs(int(first)) // abs(int(second))
    signed = -quotient if (first < 0) != (second < 0) else quotient
    return _as_kind(signed, kind)


def modulo(left: object, right: object) -> object:
    """``Mod``, whose sign follows the left operand, as VBA's does."""
    if left is NULL or right is NULL:
        return NULL
    kind = _whole_width(left, right)
    first = to_integer(left, "LongLong")
    second = to_integer(right, "LongLong")
    if second == 0:
        raise error(ERR_DIVISION_BY_ZERO)
    remainder = abs(int(first)) % abs(int(second))
    return _as_kind(-remainder if first < 0 else remainder, kind)


def power(left: object, right: object) -> object:
    """``^``, which is a Double even when both sides are whole numbers."""
    if left is NULL or right is NULL:
        return NULL
    base = float(to_number(left))
    exponent = float(to_number(right))
    try:
        result = base**exponent
    except (OverflowError, ValueError):
        raise error(ERR_OVERFLOW) from None
    if isinstance(result, complex):
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return float(result)


def concat(left: object, right: object) -> object:
    """``&``, which reads Null as an empty string unless both sides are Null."""
    if left is NULL and right is NULL:
        return NULL
    first = "" if left is NULL else to_text(left)
    second = "" if right is NULL else to_text(right)
    return first + second


def negate(value: object) -> object:
    if value is NULL:
        return NULL
    number = to_number(value)
    kind = type_name(value)
    if kind in ("Boolean", "Empty"):
        kind = "Integer"
    elif kind in ("String", "Date"):
        kind = type_name(number)
    return _as_kind(-number, kind if kind in _WIDENING else "Double")


def logical_not(value: object) -> object:
    """``Not``, which is a bitwise complement on the integer value."""
    if value is NULL:
        return NULL
    if isinstance(value, bool):
        return not value
    return _as_kind(~int(to_integer(value, "LongLong")), _whole_width(value, value))


def _bitwise(left: object, right: object, op: str) -> object:
    booleans = isinstance(left, bool) and isinstance(right, bool)
    kind = _whole_width(left, right)
    first = int(to_integer(left, "LongLong"))
    second = int(to_integer(right, "LongLong"))
    if op == "and":
        result = first & second
    elif op == "or":
        result = first | second
    elif op == "xor":
        result = first ^ second
    elif op == "eqv":
        result = ~(first ^ second)
    else:  # imp
        result = ~first | second
    if booleans:
        return result != 0
    return _as_kind(result, kind)


def logical_and(left: object, right: object) -> object:
    """``And``, whose Null rules are the ones a WHERE clause uses."""
    if left is NULL or right is NULL:
        other = right if left is NULL else left
        if other is not NULL and not _truth_is_unknown(other):
            return NULL if to_bool(other) else False
        return NULL
    return _bitwise(left, right, "and")


def logical_or(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        other = right if left is NULL else left
        if other is not NULL and not _truth_is_unknown(other):
            return True if to_bool(other) else NULL
        return NULL
    return _bitwise(left, right, "or")


def _truth_is_unknown(value: object) -> bool:
    return value is NULL


def logical_xor(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    return _bitwise(left, right, "xor")


def logical_eqv(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    return _bitwise(left, right, "eqv")


def logical_imp(left: object, right: object) -> object:
    if left is NULL and right is NULL:
        return NULL
    if left is NULL:
        return True if right is not NULL and to_bool(right) else NULL
    if right is NULL:
        return NULL if to_bool(left) else True
    return _bitwise(left, right, "imp")


def compare(op: str, left: object, right: object, *, text_compare: bool = False) -> object:
    """A comparison, with Null answering Null rather than True or False."""
    if left is NULL or right is NULL:
        return NULL
    if op == "is":
        return left is right
    order = _order(left, right, text_compare=text_compare)
    if op == "=":
        return order == 0
    if op == "<>":
        return order != 0
    if op == "<":
        return order < 0
    if op == ">":
        return order > 0
    if op == "<=":
        return order <= 0
    if op == ">=":
        return order >= 0
    raise error(ERR_TYPE_MISMATCH, f"unknown comparison {op!r}")


def _order(left: object, right: object, *, text_compare: bool) -> int:
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(left, VBAObject):
        left = left.vba_value()
    if isinstance(right, VBAObject):
        right = right.vba_value()
    if isinstance(left, str) and isinstance(right, str):
        first, second = (left.upper(), right.upper()) if text_compare else (left, right)
        return (first > second) - (first < second)
    if isinstance(left, str) != isinstance(right, str):
        # A number against a string reads the string as a number, so
        # 1 < "1" is False rather than True.  Measured against Excel;
        # a string that is not a number is a type mismatch.
        if left is EMPTY or right is EMPTY:
            left = "" if left is EMPTY else left
            right = "" if right is EMPTY else right
            return _order(left, right, text_compare=text_compare)
    first_number = to_number(left)
    second_number = to_number(right)
    if isinstance(first_number, Decimal) or isinstance(second_number, Decimal):
        first_number = Decimal(str(first_number))
        second_number = Decimal(str(second_number))
    return (first_number > second_number) - (first_number < second_number)


_LIKE_CLASS = re.compile(r"\[(!?)((?:[^\]])*)\]")


def like_match(text: str, pattern: str, *, text_compare: bool = False) -> bool:
    """VBA's ``Like``: ``?`` one character, ``*`` any run, ``#`` a digit."""
    regex: list[str] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "?":
            regex.append(".")
        elif char == "*":
            regex.append(".*")
        elif char == "#":
            regex.append("[0-9]")
        elif char == "[":
            match = _LIKE_CLASS.match(pattern, index)
            if not match:
                raise error(93)
            body = match.group(2).replace("\\", "\\\\").replace("^", "\\^")
            regex.append(f"[{'^' if match.group(1) else ''}{body}]")
            index = match.end()
            continue
        else:
            regex.append(re.escape(char))
        index += 1
    flags = re.DOTALL | (re.IGNORECASE if text_compare else 0)
    return re.fullmatch("".join(regex), text, flags) is not None
