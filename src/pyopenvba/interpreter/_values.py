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
import unicodedata
from contextlib import AbstractContextManager
from decimal import ROUND_HALF_EVEN, ROUND_HALF_UP, Context, Decimal, InvalidOperation, localcontext
from typing import Any, Callable, Final
from weakref import ReferenceType

from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter import _locale

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

#: The text Err.Description and Error() give for each error number VBA names, measured
#: in live Excel for 1 to 1000 (scripts/measure_vba_error_text.py); any other is
#: "Application-defined or object-defined error".
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
    47: "Too many DLL application clients",
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
    96: ("Unable to sink events of object because the object is already firing events to the maximum number of "
         "event receivers that it supports"),
    97: "Can not call friend function on object which is not an instance of defining class",
    98: ("A property or method call cannot include a reference to a private object, either as an argument or as a "
         "return value"),
    321: "Invalid file format",
    322: "Can't create necessary temporary file",
    325: "Invalid format in resource file",
    380: "Invalid property value",
    381: "Invalid property array index",
    382: "Set not supported at runtime",
    383: "Set not supported (read-only property)",
    385: "Need property array index",
    387: "Set not permitted",
    393: "Get not supported at runtime",
    394: "Get not supported (write-only property)",
    422: "Property not found",
    423: "Property or method not found",
    424: "Object required",
    429: "ActiveX component can't create object",
    430: "Class does not support Automation or does not support expected interface",
    432: "File name or class name not found during Automation operation",
    438: "Object doesn't support this property or method",
    440: "Automation error",
    442: ("Connection to type library or object library for remote process has been lost. Press OK for dialog to "
          "remove reference."),
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
    460: "Invalid clipboard format",
    461: "Method or data member not found",
    462: "The remote server machine does not exist or is unavailable",
    463: "Class not registered on local machine",
    481: "Invalid picture",
    482: "Printer error",
    735: "Can't save file to TEMP",
    744: "Search text not found",
    746: "Replacements too long",
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
    """The Variant-only Decimal subtype: a 96-bit integer and a scale of 0 to 28.

    So at most 28 decimal places, and a mantissa no larger than
    79228162514264337593543950335: CDec(1) / 3 has 28 threes, and the largest
    Decimal keeps all 29 digits (native probes). A value is rounded half to
    even to fit; one too large for any scale is error 6.
    """

    __slots__ = ()

    LARGEST: Final = 79228162514264337593543950335

    def __new__(cls, value: str | int | float | Decimal = "0") -> VBADecimal:
        with localcontext() as context:
            context.prec = 60
            exact = Decimal(repr(value)) if isinstance(value, float) else Decimal(value)
            if exact.is_finite():
                if abs(exact) > cls.LARGEST:
                    # Too large at any scale, whatever exponent it is written with:
                    # CDec(CSng(3.4E+38)) is error 6 (native probe).
                    raise error(ERR_OVERFLOW)
                if _places(exact) > 28:
                    exact = exact.quantize(Decimal(1).scaleb(-28), rounding=ROUND_HALF_EVEN)
                while abs(exact.scaleb(_places(exact))) > cls.LARGEST:
                    if _places(exact) <= 0:
                        raise error(ERR_OVERFLOW)
                    exact = exact.quantize(Decimal(1).scaleb(1 - _places(exact)), rounding=ROUND_HALF_EVEN)
        return super().__new__(cls, exact)


def decimal_of(number: object) -> Decimal:
    """A number as the Decimal VBA makes of it: a Double through its 15
    significant digits and a Single through its 7, so CDec(1 / 3) is
    0.333333333333333 and CDec(CSng(0.1)) is 0.1 (native probes)."""
    if isinstance(number, VBASingle):
        return Decimal(format(float(number), ".7g"))
    if isinstance(number, float):
        return Decimal(format(number, ".15g"))
    return Decimal(str(number))


def _places(value: Decimal) -> int:
    """The decimal places a finite Decimal is written with: its negated exponent."""
    exponent = value.as_tuple().exponent
    return -exponent if isinstance(exponent, int) else 0


#: Half a second, as a part of a day: what a Date's moment is rounded by.
_HALF_SECOND: Final = 0.5 / 86400.0
#: 9999-12-31, the last day a Date holds.
_LAST_DAY: Final = 2958465.0


class VBADate:
    """A VBA Date: a Double whose whole part is days from 1899-12-30."""

    __slots__ = ("serial",)

    #: Day zero.  A Date of 0 is midnight on it, and 1 is 1899-12-31.
    EPOCH: Final = _dt.datetime(1899, 12, 30)

    def __init__(self, serial: float) -> None:
        # 100-01-01 runs to just short of -657435, its time counting forwards (native probe: -657434.99999 is
        # 1/1/100 11:59:59 PM, -657435 is error 6).
        if not -657435.0 < serial < 2958466.0:
            raise error(ERR_OVERFLOW)
        self.serial = float(serial)

    @classmethod
    def from_datetime(cls, when: _dt.datetime | _dt.date | _dt.time) -> VBADate:
        if isinstance(when, _dt.datetime):
            days = (when.date() - cls.EPOCH.date()).days
            fraction = (when.hour * 3600 + when.minute * 60 + when.second) / 86400.0 + when.microsecond / 86400e6
            # Before day zero the day counts back and the time forward, so
            # 1899-12-29 06:00 is -1.25, not -0.75 (native probes).
            return cls(days + fraction if days >= 0 else days - fraction)
        if isinstance(when, _dt.date):
            return cls((when - cls.EPOCH.date()).days)
        seconds = when.hour * 3600 + when.minute * 60 + when.second
        return cls(seconds / 86400.0 + when.microsecond / 86400e6)

    def to_datetime(self) -> _dt.datetime:
        """The moment this serial names, to the second VBA stores.

        Rounded as oleaut32's VariantTimeToSystemTime rounds it, which VBA's
        CStr, Format, Hour, Minute and Second go through: half a second is
        added to the serial away from zero, the days are cut toward zero and
        the seconds cut from the part of a day left over. A negative serial
        counts the day backwards but the time forwards, so one whose time
        rounds past its own day takes the day after it. A serial the half
        second would carry past the start of 9999-12-31 is cut instead, so
        the last day's times never round up. That matches the system's own
        conversion on 760,000 serials, ties and both ends of the Date range
        among them.
        """
        moved = self.serial + _HALF_SECOND if self.serial >= 0 else self.serial - _HALF_SECOND
        if moved > _LAST_DAY:
            moved = self.serial
        days = math.floor(moved) if moved >= 0 else -math.floor(-moved)
        seconds = int(abs(moved - days) * 86400.0)
        if self.serial < 0 and days != -math.floor(-self.serial):
            days = -math.floor(-self.serial) + 1
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


def release_in_order(items: list[object]) -> None:
    """Drop each item's reference first to last, so objects are freed in that order."""
    for index in range(len(items)):
        items[index] = None


class VBAArray:
    """A VBA array: bounds per dimension, and a flat store in VBA's order.

    VBA stores an array column-major, which matters only where somebody
    reads it as a block (a Range's Value, say), so the flat index is
    computed the way VBA lays it out rather than the way Python would.
    """

    __slots__ = ("bounds", "items", "element_type", "fixed", "borrowed_elements", "element_factory", "parent_array",
                 "auto_new", "__weakref__")

    def __init__(
        self,
        bounds: list[tuple[int, int]],
        *,
        element_type: str = "Variant",
        fixed: bool = False,
        items: list[object] | None = None,
        element_factory: Callable[[], object] | None = None,
    ) -> None:
        for lower, upper in bounds:
            if upper < lower - 1:
                raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        self.borrowed_elements = 0
        self.element_factory = element_factory
        #: For an array declared As New: what makes an element the first time it is read while Nothing.
        self.auto_new: Callable[[], object] | None = None
        self.parent_array: ReferenceType[VBAArray] | None = None
        self.bounds = bounds
        self.element_type = element_type
        self.fixed = fixed
        self.items = items if items is not None else [element_factory() if element_factory is not None else default_for(element_type) for _ in range(self.size)]
        for item in self.items:
            binder = getattr(item, "vba_bind_array", None)
            if callable(binder):
                binder(self)

    def __del__(self) -> None:
        # An array that can hold objects releases them first to last, as VBA's
        # Erase and end of scope do (native probes); a list would go last to first.
        if self.element_type not in STORAGE_WIDTH and not self.element_type.startswith("String"):
            release_in_order(self.items)

    @property
    def size(self) -> int:
        if not self.bounds:
            # A dynamic array no ReDim has sized: unallocated, so it holds nothing.
            return 0
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
        offset = self.offset(subscripts)
        item = self.items[offset]
        if item is NOTHING and self.auto_new is not None:
            # Dim c(1) As New Collection makes c(1) where it is first used (native probe).
            item = self.items[offset] = self.auto_new()
        return item

    def set(self, subscripts: list[int], value: object) -> None:
        offset = self.offset(subscripts)
        value = coerce(value, self.element_type)
        assigner = getattr(self.items[offset], "vba_assign_value", None)
        if callable(assigner):
            assigner(value)
        else:
            self.items[offset] = copy_value(value)
            binder = getattr(self.items[offset], "vba_bind_array", None)
            if callable(binder):
                binder(self)

    def copy_from(self, source: VBAArray) -> None:
        """Copy fixed record-field storage without replacing element addresses."""
        if self.bounds != source.bounds or self.element_type != source.element_type:
            raise error(ERR_TYPE_MISMATCH)
        for index, item in enumerate(source.items):
            self.set(self._subscripts(index), item)

    def borrow_arrays(self) -> list[VBAArray]:
        arrays: list[VBAArray] = [self]
        parent = self.parent_array() if self.parent_array is not None else None
        while parent is not None and parent not in arrays:
            arrays.append(parent)
            parent = parent.parent_array() if parent.parent_array is not None else None
        return arrays

    def resized(self, bounds: list[tuple[int, int]], *, preserve: bool) -> VBAArray:
        """A new array of ``bounds``, carrying the old contents if asked.

        VBA's Preserve keeps elements by subscript, and only the last
        dimension may change size; anything else is error 9.
        """
        fresh = VBAArray(bounds, element_type=self.element_type, element_factory=self.element_factory)
        fresh.auto_new = self.auto_new
        if not preserve or not self.bounds:
            # Preserving an unallocated array keeps nothing (native probe).
            return fresh
        if len(bounds) != len(self.bounds):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        if any(old != new for old, new in zip(self.bounds[:-1], bounds[:-1])):
            raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
        if bounds[-1][0] != self.bounds[-1][0]:
            # Only the last dimension's upper bound may move: (1 To 3) to (0 To 3) is error 9 (native probe).
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
    "LongPtr": 8,
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
        return "Object()" if holds_objects(value) else f"{value.element_type}()"
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(value, VBAObject):
        return value.vba_type_name
    return "Variant"


def holds_objects(array: VBAArray) -> bool:
    """Whether an array is of an object type, class, Collection or Range alike,
    which TypeName calls "Object()" and VarType 8201 (native probes)."""
    return array.element_factory is None and default_for(array.element_type) is NOTHING


def default_for(declared: str) -> object:
    """What a variable of ``declared`` holds before anything is assigned."""
    base = declared.rstrip("()")
    if base == 'LongPtr':
        base = 'LongLong'
    if base.startswith('String*'):
        return '\x00' * int(base.partition('*')[2])
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
    re.VERBOSE | re.ASCII,
)


#: The digits of a number a US-English Office reads: commas anywhere after
#: the first digit, even after the point, then an optional exponent. Digits are
#: ASCII: en-US VBA refuses CDbl("１２") (live probe, tests/fixtures/vba_locale_live.json).
_DECORATED_BODY = re.compile(r"(?:\d[\d,]*(?:\.[\d,]*)?|\.\d[\d,]*)(?:[eEdD][-+]?\d+)?", re.ASCII)


def _decorated_number(text: str) -> int | float | None:
    """A number written as a US-English Office reads one, beyond the plain
    form: ``1,000``, ``$5``, ``5$``, ``5-``, ``(5)``, ``($5)``, ``$(5)``.

    Measured over 52 strings, with IsNumeric, CDbl, CLng and arithmetic
    agreeing on every one: a currency sign once, before or after, inside or
    outside parentheses; a sign once, before or after; parentheses for a
    negative, never with a sign; spaces and tabs between the parts but not
    inside the digits. ``,5``, ``$$5``, ``--5``, ``(-5)``, ``5%`` are not
    numbers. Returns None for text that is not one.

    The currency sign is the locale's: ja-JP takes the yen sign, its
    full-width form and the backslash, and not ``$``
    (tests/fixtures/vba_locale.json).
    """
    locale = _locale.current()
    signs = locale.currency_signs
    rest = text.strip(" \t")
    if locale is _locale.JA_JP:
        # ja-JP takes the sign repeated at the end: 12¥¥ is 12.
        while len(rest) >= 2 and rest[-1] in signs and rest[-2] in signs:
            rest = rest[:-1]
    currency = sign = 0
    negative = parens = False
    if rest[:1] and rest[:1] in signs:
        currency, rest = 1, rest[1:].strip(" \t")
    elif rest[-1:] and rest[-1:] in signs and rest[:1] == "(":
        currency, rest = 1, rest[:-1].strip(" \t")
    if len(rest) >= 2 and rest[0] == "(" and rest[-1] == ")":
        parens = negative = True
        rest = rest[1:-1].strip(" \t")
    for end in (0, -1):
        while rest and (rest[end] in "+-" or rest[end] in signs):
            mark = rest[end]
            rest = (rest[1:] if end == 0 else rest[:-1]).strip(" \t")
            if mark in signs:
                currency += 1
            else:
                sign += 1
                negative = mark == "-"
    if currency > 1 or sign > 1 or (parens and sign) or not _DECORATED_BODY.fullmatch(rest):
        return None
    body = rest.replace(",", "").lower().replace("d", "e")
    if "." not in body and "e" not in body:
        whole = -int(body) if negative else int(body)
        return VBAInt(whole, "Long" if abs(whole) > 32767 else "Integer")
    value = float(body)
    return -value if negative else value


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
        default = value.vba_value()
        if default is value:
            raise error(ERR_TYPE_MISMATCH)
        return to_number(default, context=context)
    raise error(ERR_TYPE_MISMATCH, f"cannot read {type_name(value)} as a number{context}")


def text_to_number(text: str) -> int | float:
    """A string as VBA reads it in an arithmetic context.

    Stricter than ``Val``: the whole string has to be a number, or VBA
    raises 13 rather than taking the leading digits. Under ja-JP, full-width
    digits and signs read as their ASCII forms (``１２`` is 12).
    """
    if _locale.current().folds_width:
        text = _locale.fold_width(text)
    match = _NUMBER_TEXT.match(text)
    if not match:
        decorated = _decorated_number(text)
        if decorated is None:
            raise error(ERR_TYPE_MISMATCH)
        return decorated
    sign = -1 if match.group("sign") == "-" else 1
    if match.group("hex") is not None:
        return VBAInt(sign * int(match.group("hex"), 16), "Long")
    if match.group("octal") is not None:
        return VBAInt(sign * int(match.group("octal"), 8), "Long")
    body = match.group("plain").lower().replace("d", "e")
    if "." not in body and "e" not in body:
        return VBAInt(sign * int(body), "Long" if abs(int(body)) > 32767 else "Integer")
    return sign * float(body)


def condition(value: object) -> bool:
    """A value as an If, ElseIf, While, Do, IIf or Case reads it: Null is False
    there, not error 94, so If Null takes the Else and Do Until Null loops on
    (native probes). Assigning Null to a Boolean is still error 94."""
    return False if value is NULL else to_bool(value)


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
        with localcontext() as context:
            # A Decimal holds 29 digits; the default context would round the last away.
            context.prec = 60
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
        # An exact tie goes away from zero, a Single's as a Double's (tests/fixtures/vba_semantics).
        context.prec, context.rounding = precision, ROUND_HALF_UP
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
    if isinstance(value, VBAArray) and value.element_type == "Byte" and value.dimensions <= 1:
        return bytes_text(value)
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(value, VBAObject):
        default = value.vba_value()
        if default is value:
            raise error(ERR_TYPE_MISMATCH)
        return to_text(default, context=context)
    raise error(ERR_TYPE_MISMATCH, f"cannot read {type_name(value)} as a string{context}")


def bytes_text(array: VBAArray) -> str:
    """A Byte array as a String: its bytes read as UTF-16LE, two to a character (native probe)."""
    data = bytes(int(to_integer(item, "Byte")) for item in array.items)
    return data[: len(data) // 2 * 2].decode("utf-16-le", errors="surrogatepass")


def text_bytes(text: str) -> VBAArray:
    """A String as a Byte array: two UTF-16LE bytes per character (native probe)."""
    data = text.encode("utf-16-le", errors="surrogatepass")
    return VBAArray([(0, len(data) - 1)] if data else [(0, -1)], element_type="Byte",
                    items=[VBAInt(byte, "Byte") for byte in data])


def date_text(value: VBADate) -> str:
    """A Date as VBA's default conversion prints it.

    The moment is rounded to the second first. On 30 December 1899, day
    0, it prints as the time alone, midnight included, which is why a cell
    holding 0.5 shows only a clock; at midnight on any other day, as the
    date alone, so a time a hair short of midnight prints the next day.
    Under ja-JP the date is year/MM/dd and the clock 24-hour, as
    2026/01/05 9:07:07 (oleaut32's VarBstrFromDate, LCID 1041).
    """
    when = value.to_datetime()
    japanese = _locale.current() is _locale.JA_JP
    if japanese:
        day = f"{when.year}/{when.month:02d}/{when.day:02d}"
        clock = f"{when.hour}:{when.minute:02d}:{when.second:02d}"
    else:
        day = f"{when.month}/{when.day}/{when.year}"
        clock = _time_text(when)
    if when.date() == VBADate.EPOCH.date():
        return clock
    if when.time() == _dt.time():
        return day
    return f"{day} {clock}"


def _time_text(when: _dt.datetime) -> str:
    hour = when.hour % 12 or 12
    suffix = "AM" if when.hour < 12 else "PM"
    return f"{hour}:{when.minute:02d}:{when.second:02d} {suffix}"


def parse_date_text(text: str) -> VBADate | None:
    """``text`` as a Date, or None where VBA would not read one.

    Reads text as a US-English Office does (CDate, IsDate, DateValue),
    measured over 182 strings in ``tests/fixtures/date_text.json``:

    - Numbers, month names (any prefix of three letters or more, a dot
      after one allowed), AM/PM (``AM``, ``PM``, ``A``, ``P``), separated by
      ``/``, ``-``, ``,`` or spaces. Weekday names are not read.
    - Three numbers are month-day-year, else year-month-day, else
      day-month-year: the first that is a real date. Two are month-day,
      day-month, month-year or year-month, the year this one when there is
      none. With a month name, the first number is the day unless it is
      past 31; one number past 31 is the year, the day then the 1st.
    - A year under 100 is 1950 to 2049; 100 to 9999 is itself.
    - A time is ``h:mm[:ss]``, or an hour beside AM/PM, before or after the
      date and set off by spaces; AM/PM may come before it.
    """
    tokens = _date_tokens(text)
    if tokens is None:
        return None
    return _date_from_tokens(tokens)


_MONTH_NAMES: Final = ("january", "february", "march", "april", "may", "june", "july",
                       "august", "september", "october", "november", "december")
_AM_PM: Final = {"am": "am", "a": "am", "pm": "pm", "p": "pm"}

#: A token of date text: (kind, value, spaced). kind is "num" (value the digits),
#: "month" (1-12), "ampm", "sep" (/ - ,) or ":"; spaced says whitespace came before it.
_DateToken = tuple[str, object, bool]


def _date_tokens(text: str, japanese: bool = False) -> list[_DateToken] | None:
    tokens: list[_DateToken] = []
    at, spaced = 0, False
    while at < len(text):
        char = text[at]
        if char in " \t":
            at, spaced = at + 1, True
            continue
        if japanese and char in _UNITS:
            tokens.append(("unit", char, spaced))
            end = at + 1
        elif japanese and text.startswith(("午前", "午後"), at):
            tokens.append(("ampm", "am" if text[at + 1] == "前" else "pm", spaced))
            end = at + 2
        elif "0" <= char <= "9":
            end = at
            while end < len(text) and "0" <= text[end] <= "9":
                end += 1
            tokens.append(("num", text[at:end], spaced))
        elif char.isalpha():
            end = at
            while end < len(text) and text[end].isalpha():
                end += 1
            word = text[at:end].lower()
            if japanese and (word in ("a", "p") or (word[:2] in ("am", "pm") and word not in _AM_PM)):
                # ja-JP takes A and P after a number or a clock, and before one only with a space between,
                # and reads PMOct as PM and Oct.
                raise VBAUnsupportedError("reading A, P or a word joined to AM or PM in Japanese date text is "
                                          "not implemented")
            if word in _AM_PM:
                tokens.append(("ampm", _AM_PM[word], spaced))
            else:
                month = next((index for index, name in enumerate(_MONTH_NAMES, 1)
                              if len(word) >= 3 and name.startswith(word)), None)
                if month is None:
                    return None
                tokens.append(("month", month, spaced))
                if end < len(text) and text[end] == ".":
                    end += 1
        elif char in "/-,":
            tokens.append(("sep", char, spaced))
            end = at + 1
        elif char in ":.":
            # A dot separates a time's parts as a colon does: CDate("1.5") is 1:05 AM,
            # and "1. 5" too (native probes).
            tokens.append((":", char, spaced))
            end = at + 1
        else:
            return None
        at, spaced = end, False
    return tokens


def _date_from_tokens(tokens: list[_DateToken], japanese: bool = False) -> VBADate | None:
    # The time: numbers joined by colons, or one number beside AM/PM.
    times: list[tuple[int, int]] = []  # (start, end) token spans
    index = 0
    while index < len(tokens):
        if tokens[index][0] == "num" and index + 1 < len(tokens) and tokens[index + 1][0] == ":":
            end = index + 1
            while end + 1 < len(tokens) and tokens[end][0] == ":" and tokens[end + 1][0] == "num" and (not tokens[end + 1][2] or tokens[end][1] == ".") and not tokens[end][2]:
                end += 2
            if tokens[end - 1][0] != "num" or (end < len(tokens) and tokens[end][0] == ":"):
                return None
            times.append((index, end))
            index = end
        else:
            index += 1
    marks = [at for at, token in enumerate(tokens) if token[0] == "ampm"]
    if len(marks) > 1 or len(times) > 1:
        return None
    if marks and japanese and not times and not (marks[0] > 0 and tokens[marks[0] - 1][0] == "num"):
        # ja-JP passes over AM or PM with no clock and no hour before it: 午後2019/10/28 is the date,
        # 午後9 nothing; 9 午後 is 9 PM.
        tokens = tokens[:marks[0]] + tokens[marks[0] + 1:]
        marks = []
    if marks:
        mark = marks[0]
        if times and times[0][1] == mark:
            times = [(times[0][0], mark + 1)]
        elif times and times[0][0] == mark + 1:
            times = [(mark, times[0][1])]
        elif not times and mark > 0 and tokens[mark - 1][0] == "num":
            times = [(mark - 1, mark + 1)]
        elif not times and mark + 1 < len(tokens) and tokens[mark + 1][0] == "num":
            times = [(mark, mark + 2)]
        else:
            return None
    clock: _dt.time | None = None
    rest = tokens
    if times:
        start, end = times[0]
        part = tokens[start:end]
        # Set off by spaces from the date, never by a separator; ja-JP needs no space.
        if (start > 0 and (tokens[start - 1][0] == "sep" or not (part[0][2] or japanese))) or \
                (end < len(tokens) and (tokens[end][0] == "sep" or not (tokens[end][2] or japanese))):
            return None
        clock = _clock(part)
        if clock is None:
            return None
        rest = tokens[:start] + tokens[end:]
        if japanese and start == 0 and rest:
            raise VBAUnsupportedError("reading Japanese date text with the time before the date is not implemented")
    items = _date_items(rest)
    if items is None:
        return None
    if not items:
        return VBADate.from_datetime(clock) if clock is not None else None
    day = _calendar_day(items, japanese)
    if day is None:
        return None
    if clock is None:
        return VBADate.from_datetime(day)
    return VBADate.from_datetime(_dt.datetime.combine(day, clock))


def _clock(part: list[_DateToken]) -> _dt.time | None:
    mark = next((str(token[1]) for token in part if token[0] == "ampm"), None)
    numbers = [int(str(token[1])) for token in part if token[0] == "num"]
    if len(numbers) > 3:
        return None
    hour, minute, second = (numbers + [0, 0])[:3]
    if not 0 <= hour <= 23 or not 0 <= minute <= 59 or not 0 <= second <= 59:
        return None
    if mark == "pm" and hour < 12:
        hour += 12
    elif mark == "am" and hour == 12:
        hour = 0
    return _dt.time(hour, minute, second)


def _date_items(tokens: list[_DateToken]) -> list[_DateToken] | None:
    """The numbers and month names of the date, separators checked and dropped."""
    items: list[_DateToken] = []
    expect_item = True
    for token in tokens:
        if token[0] == "sep":
            if expect_item:
                return None
            expect_item = True
        elif token[0] in ("num", "month"):
            items.append(token)
            expect_item = False
        else:
            return None
    if items and expect_item:
        return None
    return items


def _year(value: int) -> int:
    """A year as written: under 100 is 1950 to 2049 (native probes)."""
    return value + (2000 if value < 50 else 1900) if value < 100 else value


def _real_day(year: int, month: int, day: int) -> _dt.date | None:
    if not 100 <= year <= 9999:
        return None
    try:
        return _dt.date(year, month, day)
    except ValueError:
        return None


def _this_year() -> int:
    """The year text with none is read in."""
    return _dt.date.today().year


def _calendar_day(items: list[_DateToken], japanese: bool = False) -> _dt.date | None:
    months = [int(str(token[1])) for token in items if token[0] == "month"]
    numbers = [int(str(token[1])) for token in items if token[0] == "num"]
    this_year = _this_year()
    if len(months) > 1 or len(items) > 3:
        return None
    if months:
        month = months[0]
        if len(numbers) == 1:
            value = numbers[0]
            return _real_day(this_year, month, value) if value <= 31 else _real_day(_year(value), month, 1)
        if len(numbers) == 2:
            first, second = numbers
            if japanese:
                # ja-JP reads the year first, "Mar 4 5" as 5 March 2004, then the day: "Oct 28 2019".
                return _real_day(_year(first), month, second) or _real_day(_year(second), month, first)
            if first > 31:
                return _real_day(_year(first), month, second)
            return _real_day(_year(second), month, first)
        return None
    if len(numbers) == 3:
        a, b, c = numbers
        # en-US tries month-day-year first, ja-JP year-month-day.
        orders = ((a, b, c), (c, a, b), (c, b, a)) if japanese else ((c, a, b), (a, b, c), (c, b, a))
        for year, month, day in orders:
            found = _real_day(_year(year), month, day)
            if found is not None:
                return found
        return None
    if len(numbers) == 2:
        a, b = numbers
        for year, month, day in ((this_year, a, b), (this_year, b, a), (_year(b), a, 1), (_year(a), b, 1)):
            found = _real_day(year, month, day)
            if found is not None:
                return found
        return None
    return None


#: The marks that say which number of Japanese date text is which: year, month, day, hour, minute, second.
_UNITS: Final = "年月日時分秒"
#: An era before the year it counts: its name, its first character or its letter, then the year, or 元 (the
#: first year) before 年.
_ERA_YEAR: Final = re.compile(r"(?<![A-Za-z])(令和|平成|昭和|大正|明治|[令平昭大明RHSTMrhstm])[ \t]*(\d+|元(?=年))")


def read_date_text(text: str) -> VBADate | None:
    """``text`` as a Date as CDate, IsDate and DateValue read it under the
    locale in force, or None. A date literal in code reads as en-US whatever
    the locale (``parse_date_text``)."""
    if _locale.current() is _locale.JA_JP:
        return _japanese_date(text)
    return parse_date_text(text)


def _japanese_date(text: str) -> VBADate | None:
    """Date text as ja-JP reads it, checked against oleaut32 with LCID 1041
    over the strings in tests/fixtures/vba_locale.json.

    Full-width letters and digits read as ASCII. Three numbers are
    year-month-day first, then month-day-year, then day-month-year, and a
    month name's first number is the year. An era before the first number
    (令和8, 令8, R8, 平成元年) makes it the era's year, past the era's end too.
    Numbers marked 年, 月 and 日 are the year, month and day in any order,
    which needs the month and the year or the day; 時, 分 and 秒 mark the
    hour, minute and second. 午前 and 午後 are AM and PM.
    """
    text = _locale.fold_width(text)
    found = list(_ERA_YEAR.finditer(text))
    if len(found) > 1:
        return None
    if found:
        match = found[0]
        name = match.group(1)
        era = next(era for era in _locale.JA_JP.eras
                   if name in (era.name, era.short) or name.upper() == era.letter)
        count = 1 if match.group(2) == "元" else int(match.group(2))
        # A space keeps the year apart from a number before it.
        text = f"{text[:match.start()]} {era.start.year + count - 1}{text[match.end():]}"
    tokens = _date_tokens(text, japanese=True)
    if tokens is None:
        return None
    if any(token[0] == "unit" for token in tokens):
        return _marked_date(tokens)
    return _date_from_tokens(tokens, japanese=True)


def _marked_date(tokens: list[_DateToken]) -> VBADate | None:
    """Japanese date text whose numbers carry 年月日 or 時分秒."""
    marked: dict[str, int] = {}
    rest: list[_DateToken] = []
    places: list[int] = []  # where each token of rest stood
    first_date = first_time = len(tokens)
    at = 0
    while at < len(tokens):
        kind, value, _spaced = tokens[at]
        if kind == "num" and at + 1 < len(tokens) and tokens[at + 1][0] == "unit":
            unit = str(tokens[at + 1][1])
            if unit in marked:
                return None
            marked[unit] = int(str(value))
            if unit in "年月日":
                first_date = min(first_date, at)
            else:
                first_time = min(first_time, at)
            at += 2
            if unit in "時分" and at < len(tokens) and tokens[at][0] == "num" and \
                    not (at + 1 < len(tokens) and tokens[at + 1][0] == "unit" and str(tokens[at + 1][1]) in "時分秒"):
                # A number after 時 or 分 is another part of the time: 13時2019年 is nothing.
                return None
            continue
        if kind == "unit":
            return None
        rest.append(tokens[at])
        places.append(at)
        at += 1
    timed = any(unit in marked for unit in "時分秒")
    dated = any(unit in marked for unit in "年月日")
    marks = [index for index, token in enumerate(rest) if token[0] == "ampm"]
    mark = str(rest[marks[0]][1]) if marks else None
    if dated and not timed and len(rest) == 1 and marks:
        # AM or PM with no time is passed over before the date and fails after it.
        if places[0] > first_date:
            return None
        rest = []
    plain = [place for place, token in zip(places, rest) if token[0] != "ampm"]
    if (timed and plain and plain[0] > first_time) or (dated and plain and plain[0] < first_date) or \
            (timed and dated and first_time < first_date):
        raise VBAUnsupportedError("reading Japanese date text with the time before the date is not implemented")
    day: _dt.date | None = None
    clock: _dt.time | None = None
    if dated:
        if "月" not in marked or not ("年" in marked or "日" in marked):
            return None
        year = _year(marked["年"]) if "年" in marked else _this_year()
        day = _real_day(year, marked["月"], marked.get("日", 1))
        if day is None:
            return None
    if timed:
        hour, minute, second = marked.get("時", 0), marked.get("分", 0), marked.get("秒", 0)
        if not 0 <= hour <= 23 or not 0 <= minute <= 59 or not 0 <= second <= 59:
            return None
        if mark == "pm" and hour < 12:
            hour += 12
        elif mark == "am" and hour == 12:
            hour = 0
        clock = _dt.time(hour, minute, second)
        rest = [token for token in rest if token[0] != "ampm"]
    if rest:
        # What is left is the other half: a plain date beside a marked time, a clock beside a marked date.
        if dated and timed:
            return None
        found = _date_from_tokens(rest, japanese=True)
        if found is None:
            return None
        if dated:
            if not 0 <= found.serial < 1:
                return None
            clock = found.to_datetime().time()
        else:
            if found.serial != int(found.serial) or found.serial == 0:
                return None
            day = found.to_datetime().date()
    if day is None:
        return VBADate.from_datetime(clock) if clock is not None else None
    if clock is None:
        return VBADate.from_datetime(day)
    return VBADate.from_datetime(_dt.datetime.combine(day, clock))


def to_date(value: object) -> VBADate:
    """``value`` as a Date, as CDate reads one."""
    if isinstance(value, VBADate):
        return value
    if value is NULL:
        raise error(ERR_INVALID_USE_OF_NULL)
    if value is EMPTY:
        return VBADate(0.0)
    if isinstance(value, str):
        parsed = read_date_text(value)
        if parsed is None:
            try:
                # CDate("2024") is the serial 2024, though IsDate says False (native probe).
                return VBADate(float(text_to_number(value)))
            except VBARuntimeError:
                raise error(ERR_TYPE_MISMATCH, f"{value!r} is not a date") from None
        return parsed
    number = to_number(value)
    return VBADate(float(number))


def round_half_even(value: float | Decimal, places: int = 0) -> float | Decimal:
    """VBA rounds half to even, so 0.5 goes to 0 and 1.5 goes to 2."""
    if isinstance(value, Decimal):
        with localcontext() as context:
            context.prec = 40
            return value.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_EVEN)
    # A Double is scaled by 10^places in floating point, rounded half to even
    # there and scaled back, so Round(2.345, 2) is 2.35 (2.345 * 100 is
    # 234.50000000000003) and Round(1.255, 2) is 1.25 (native probes). A zero
    # keeps a minus sign only from a tie: Round(-0.5) is -0, Round(-0.4) and
    # Round(-0#) are 0.
    try:
        scale = 10.0 ** places
    except OverflowError:
        return value
    scaled = value * scale
    if scaled != scaled or scaled in (float("inf"), float("-inf")):
        raise error(ERR_OVERFLOW)
    whole = round(scaled)
    if scaled - math.floor(scaled) == 0.5 and whole == 0:
        return math.copysign(0.0, scaled) / scale
    return whole / scale or 0.0


_SIGNED_WIDTH: Final = {"Integer": 16, "Long": 32, "LongLong": 64}


def _hex_text_integer(text: str, kind: str) -> VBAInt | None:
    """Hex or octal text read at the target's width: CInt("&HFFFF") is -1 and
    CLng("&HFFFFFFFF") is -1, while CLng("&HFFFF") is 65535 (native probes).
    More bits than the width hold is error 6. None for other text."""
    match = _NUMBER_TEXT.match(text)
    if match is None or (match.group("hex") is None and match.group("octal") is None):
        return None
    raw = int(match.group("hex"), 16) if match.group("hex") is not None else int(match.group("octal"), 8)
    width = _SIGNED_WIDTH.get(kind)
    if width is None:
        return None
    if raw >= 1 << width:
        raise error(ERR_OVERFLOW)
    signed = raw - (1 << width) if raw >= 1 << (width - 1) else raw
    return VBAInt(-signed if match.group("sign") == "-" else signed, kind)


def to_integer(value: object, kind: str = "Long") -> VBAInt:
    """``value`` as one of the integer widths, rounding half to even."""
    if isinstance(value, str):
        wrapped = _hex_text_integer(value, kind)
        if wrapped is not None:
            return wrapped
    if value is True and kind == "Byte":
        # True is all bits set, which a Byte reads as 255: CByte(True) and b = True (native probes).
        return VBAInt(255, kind)
    number = to_number(value)
    if isinstance(number, int) and not isinstance(number, bool):
        whole = int(number)
    else:
        whole = int(round_half_even(number))
    low, high = INTEGER_LIMITS[kind]
    if not low <= whole <= high:
        raise error(ERR_OVERFLOW)
    return VBAInt(whole, kind)


def copy_value(value: object) -> object:
    """Copy VBA value containers while retaining contained object references."""
    if isinstance(value, VBAArray):
        return VBAArray(list(value.bounds), element_type=value.element_type, fixed=value.fixed,
                        items=[copy_value(item) for item in value.items], element_factory=value.element_factory)
    copier = getattr(value, "vba_copy_value", None)
    return copier() if callable(copier) else value


def coerce(value: object, declared: str) -> object:
    """``value`` stored in a variable declared ``As declared``.

    This is where a typed variable does its work: assigning 1.7 to an
    Integer stores 2, and assigning 40000 raises 6 rather than widening.
    """
    base = declared.rstrip("()")
    if base == 'LongPtr':
        base = 'LongLong'
    if base.startswith('String*'):
        width = int(base.partition('*')[2])
        return to_text(value)[:width].ljust(width)
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


_WHOLE_ORDER: Final = ("Byte", "Integer", "Long", "LongLong")
#: What divides to a Single rather than a Double, once a Single is in it.
_SINGLE_DIVISION: Final = frozenset({"Single", "Byte", "Integer", "Boolean", "Empty"})


def _result_kind(left: object, right: object, op: str) -> str:
    """The type ``+``, ``-``, ``*`` or ``/`` computes in, as a matrix of every
    pair of types measured in live Excel has it (tests/fixtures/vba_runtime_edges.json):
    Decimal wins; ``/`` is a Double but for Singles among small whole numbers;
    a Date stays a Date through ``+`` and ``-`` (Date - Date is a Double) and
    is a Double to ``*``; a Currency wins over a Double, except that ``*`` with
    a Single, Double or String is a Double; a Single beside a Long is a Double;
    and Empty takes the other side's whole-number type."""
    first, second = (type_name(value) for value in (left, right))
    kinds = {first, second}
    if "Decimal" in kinds:
        return "Decimal"
    if op == "/":
        return "Single" if "Single" in kinds and kinds <= _SINGLE_DIVISION else "Double"
    if "Date" in kinds:
        return "Double" if op == "*" or (op == "-" and first == second) else "Date"
    if "Currency" in kinds:
        return "Double" if op == "*" and kinds & {"Single", "Double", "String"} else "Currency"
    if kinds & {"String", "Double"}:
        return "Double"
    if "Single" in kinds:
        return "Double" if kinds & {"Long", "LongLong"} else "Single"
    if op == "-" and first == "Empty" and second == "Byte":
        return "Integer"
    whole = ["Integer" if kind in ("Boolean", "Error") else kind for kind in (first, second) if kind != "Empty"]
    return max(whole, key=_WHOLE_ORDER.index) if whole else "Integer"


def _decimal_context() -> AbstractContextManager[Context]:
    """Decimal arithmetic at the 29 digits a VBA Decimal holds, where the
    default context would round the last away (native probe: max - 1 + 1)."""
    context = Context(prec=29, rounding=ROUND_HALF_EVEN)
    return localcontext(context)


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
    if kind == "Date":
        return VBADate(float(value))
    result = float(value)
    if result in (float("inf"), float("-inf")):
        raise error(ERR_OVERFLOW)
    return result


def _both_numbers(left: object, right: object, op: str) -> tuple[Any, Any, str]:
    kind = _result_kind(left, right, op)
    first = to_number(left)
    second = to_number(right)
    if kind == "Decimal":
        return decimal_of(first), decimal_of(second), kind
    if kind == "Currency":
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
    first, second, kind = _both_numbers(left, right, "+")
    with _decimal_context():
        return _as_kind(first + second, kind)


def subtract(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    first, second, kind = _both_numbers(left, right, "-")
    with _decimal_context():
        return _as_kind(first - second, kind)


def multiply(left: object, right: object) -> object:
    if left is NULL or right is NULL:
        return NULL
    first, second, kind = _both_numbers(left, right, "*")
    with _decimal_context():
        return _as_kind(first * second, kind)


def divide(left: object, right: object) -> object:
    """``/``, which is always a floating division even between Integers."""
    if left is NULL or right is NULL:
        return NULL
    first, second, kind = _both_numbers(left, right, "/")
    if second == 0:
        # 0 / 0 is an overflow, as the floating-point NaN it makes, except between
        # Decimals; any other number over zero is division by zero (native probes).
        raise error(ERR_OVERFLOW if first == 0 and kind != "Decimal" else ERR_DIVISION_BY_ZERO)
    if kind == "Decimal":
        with localcontext() as context:
            context.prec = 30
            return _as_kind(Decimal(first) / Decimal(second), kind)
    return _as_kind(float(first) / float(second), kind)


def _whole_width(left: object, right: object) -> str:
    """The width ``\\``, Mod and the bitwise operators come out in.

    Both sides are rounded to a whole number first; the result is an
    Integer only where both sides already were one.
    """
    kinds = {type_name(left), type_name(right)}
    if kinds == {"Byte"}:
        return "Byte"
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
    except ZeroDivisionError:
        # 0 ^ -1 is an invalid procedure call (native probes).
        raise error(ERR_INVALID_PROCEDURE_CALL) from None
    except (OverflowError, ValueError):
        raise error(ERR_OVERFLOW) from None
    if isinstance(result, complex):
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return float(result)


def concat(left: object, right: object) -> object:
    """``&``, which reads Null as an empty string unless both sides are Null."""
    if isinstance(left, VBAErrorValue) or isinstance(right, VBAErrorValue):
        # "x" & CVErr(2042) is a type mismatch, though CStr spells one (native probes).
        raise error(ERR_TYPE_MISMATCH)
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
    if kind in ("Boolean", "Empty", "Byte"):
        # -CByte(2) is an Integer, and a negated Date stays a Date (native probes).
        kind = "Integer"
    elif kind == "Date":
        return VBADate(-float(number))
    elif kind == "String":
        kind = "Double"
    return _as_kind(-number, kind if kind in _WIDENING else "Double")


def logical_not(value: object) -> object:
    """``Not``, which is a bitwise complement on the integer value."""
    if value is NULL:
        return NULL
    if isinstance(value, bool):
        return not value
    if isinstance(value, VBAArray):
        # Not of an array complements its descriptor's address, so the idiom
        # (Not a) = -1 finds an unallocated one (native probe); an allocated
        # one answers 0 here, which keeps Not Not a nonzero.
        return VBAInt(-1 if not value.bounds else 0, "Long")
    kind = _whole_width(value, value)
    result = ~int(to_integer(value, "LongLong"))
    # A Byte has no sign bit: Not CByte(12) is 243 (native probes).
    return _as_kind(result & 0xFF if kind == "Byte" else result, kind)


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
    return _as_kind(result & 0xFF if kind == "Byte" else result, kind)


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


def compare(op: str, left: object, right: object, *, text_compare: bool = False,
            kinds: tuple[str, str] | None = None) -> object:
    """A comparison, with Null answering Null rather than True or False.

    ``kinds`` gives each operand's static type, "string", "number" or
    "variant", where the caller knows it; it decides a string against a
    number (see :func:`_mixed_order`).
    """
    if op == "is":
        from pyopenvba.interpreter._objects import VBAObject

        # Both sides have to be objects: Empty, Null or a number is error 424 (native probes).
        if not all(isinstance(side, VBAObject) or side is NOTHING for side in (left, right)):
            raise error(424)
        # An interface view and its object are one object, as COM's identity makes them.
        return getattr(left, "vba_identity", left) is getattr(right, "vba_identity", right)
    if isinstance(left, VBAErrorValue) or isinstance(right, VBAErrorValue):
        # Two Error values compare by their numbers; an Error beside anything
        # else is a type mismatch, Null included (native probes).
        if not (isinstance(left, VBAErrorValue) and isinstance(right, VBAErrorValue)):
            raise error(ERR_TYPE_MISMATCH)
        left, right = left.number, right.number
    if left is NULL or right is NULL:
        return NULL
    order = _order(left, right, text_compare=text_compare, kinds=kinds)
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


def utf16_text(text: str) -> str:
    """One Python spelling for equivalent UTF-16 strings, retaining unpaired units."""
    return text if text.isascii() else text.encode('utf-16-le', 'surrogatepass').decode('utf-16-le', 'surrogatepass')


#: Characters a text comparison reads as their narrow or katakana forms: the full-width and half-width forms block and
#: the ideographic space.
_WIDE_OR_HALF = re.compile(r"[　＀-￯]+")


def _folded(text: str) -> str:
    """Text as a text comparison weighs it, besides case: full-width forms as ASCII, half-width katakana as full,
    a half-width voiced mark joined to its kana, and hiragana as katakana (tests/fixtures/locale_report.json,
    text_compare_names.json)."""
    narrowed = _WIDE_OR_HALF.sub(lambda found: unicodedata.normalize("NFKC", found.group()), text)
    return "".join(chr(ord(one) + 0x60) if "ぁ" <= one <= "ゖ" or one in "ゝゞ" else one
                   for one in narrowed)


def string_order(first: str, second: str, *, text_compare: bool = False) -> int:
    if text_compare:
        first, second = _folded(utf16_text(first)).upper(), _folded(utf16_text(second)).upper()
    if first.isascii() and second.isascii():
        return (first > second) - (first < second)
    left = first.encode('utf-16-be', 'surrogatepass')
    right = second.encode('utf-16-be', 'surrogatepass')
    return (left > right) - (left < right)


def _order(left: object, right: object, *, text_compare: bool, kinds: tuple[str, str] | None = None) -> int:
    from pyopenvba.interpreter._objects import VBAObject

    if isinstance(left, VBAObject):
        left = left.vba_value()
    if isinstance(right, VBAObject):
        right = right.vba_value()
    if isinstance(left, str) and isinstance(right, str):
        return string_order(left, right, text_compare=text_compare)
    if isinstance(left, str) != isinstance(right, str):
        # A number against a string reads the string as a number, so
        # 1 < "1" is False rather than True.  Measured against Excel;
        # a string that is not a number is a type mismatch.
        if left is EMPTY or right is EMPTY:
            left = "" if left is EMPTY else left
            right = "" if right is EMPTY else right
            return _order(left, right, text_compare=text_compare)
        if kinds is not None and "variant" in kinds:
            return _mixed_order(left, right, kinds)
    first_number = to_number(left)
    second_number = to_number(right)
    if isinstance(first_number, Decimal) or isinstance(second_number, Decimal):
        first_number = Decimal(str(first_number))
        second_number = Decimal(str(second_number))
    return (first_number > second_number) - (first_number < second_number)


def _mixed_order(left: object, right: object, kinds: tuple[str, str]) -> int:
    """A string against a number where a Variant holds one of them (19 native probes).

    Both in Variants: the number is less than the string, unread. One in a
    Variant: the string is read as a number when it is one ("10" = 10, "True"
    = True), and otherwise the number is less. Neither: the caller's type
    mismatch rule applies.
    """
    text, number, text_first = (left, right, True) if isinstance(left, str) else (right, left, False)
    text_kind, number_kind = kinds if text_first else (kinds[1], kinds[0])
    read: object = None
    if not (text_kind == "variant" and number_kind == "variant"):
        try:
            read = to_number(text)
        except VBARuntimeError:
            if str(text).strip().lower() in ("true", "false"):
                read = VBAInt(-1 if str(text).strip().lower() == "true" else 0, "Integer")
    if read is None:
        return 1 if text_first else -1
    if text_first:
        return _order(read, number, text_compare=False)
    return _order(number, read, text_compare=False)


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
            # Escape what a Python set reads as its own syntax: "[[]" is a set holding "[".
            body = re.sub(r"([\\^\[&~|])", r"\\\1", match.group(2))
            regex.append(f"[{'^' if match.group(1) else ''}{body}]")
            index = match.end()
            continue
        else:
            regex.append(re.escape(char))
        index += 1
    flags = re.DOTALL | (re.IGNORECASE if text_compare else 0)
    return re.fullmatch("".join(regex), text, flags) is not None
