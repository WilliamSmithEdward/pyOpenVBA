"""The VBA runtime library: the functions every project has without a reference.

The inventory is the one the VBA language reference lists.  A function
here either behaves as VBA's does or raises
:class:`~pyopenvba.exceptions.VBAUnsupportedError` saying what it would
have needed; nothing returns a plausible-looking wrong answer.

What is deliberately not implemented, and why:

* the file and folder verbs (Dir, FreeFile, GetAttr, Kill, ...) reach
  outside the model into the real machine;
* the registry verbs (GetSetting, SaveSetting) do the same;
* VarPtr, StrPtr and ObjPtr hand out addresses that mean nothing here;
* Shell, SendKeys and AppActivate start or drive other programs.

MsgBox and InputBox do not block.  A message goes into the dialog
transcript and the answer comes from the queue the caller primed, so a
macro that would have stopped for a prompt runs to the end and the
prompt is visible afterwards.
"""

from __future__ import annotations

import datetime as _dt
import math
import time
from collections.abc import Callable
from decimal import Decimal
from typing import TYPE_CHECKING, Any, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (
    EMPTY,
    ERR_INVALID_PROCEDURE_CALL,
    ERR_SUBSCRIPT_OUT_OF_RANGE,
    ERR_TYPE_MISMATCH,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBADecimal,
    VBAErrorValue,
    VBAInt,
    VBASingle,
    error,
    number_text,
    parse_date_text,
    round_half_even,
    to_bool,
    to_date,
    to_integer,
    to_number,
    to_text,
    type_name,
)

if TYPE_CHECKING:
    from pyopenvba.interpreter._runtime import Interpreter

Intrinsic = Callable[["Interpreter", list[object], dict[str, object]], object]

INTRINSICS: dict[str, Intrinsic] = {}

#: The names VBA defines that this does not implement, and the reason.
UNIMPLEMENTED: Final[dict[str, str]] = {
    "dir": "reads the real file system",
    "curdir": "reads the real file system",
    "filelen": "reads the real file system",
    "filedatetime": "reads the real file system",
    "getattr": "reads the real file system",
    "freefile": "opens real files",
    "eof": "reads real files",
    "lof": "reads real files",
    "loc": "reads real files",
    "seek": "reads real files",
    "fileattr": "reads real files",
    "getsetting": "reads the Windows registry",
    "savesetting": "writes the Windows registry",
    "deletesetting": "writes the Windows registry",
    "getallsettings": "reads the Windows registry",
    "varptr": "hands out a memory address",
    "strptr": "hands out a memory address",
    "objptr": "hands out a memory address",
    "shell": "starts another program",
    "sendkeys": "drives another program's window",
    "appactivate": "drives another program's window",
    "macscript": "runs AppleScript",
    "macid": "is a Mac file type",
    "imestatus": "reports the input method editor",
    "command": "reads the command line Office was started with",
    "environ": "reads the machine's environment",
    "doevents": "yields to a message pump this does not have",
    "callbyname": "calls a member by a name computed at run time",
    "getobject": "attaches to a running application",
    "erl": "reports the last numbered line",
    "rate": "is an iterative financial solver",
    "irr": "is an iterative financial solver",
    "mirr": "is an iterative financial solver",
    "ipmt": "is not implemented yet",
    "ppmt": "is not implemented yet",
    "npv": "is not implemented yet",
    "qbcolor": "is not implemented yet",
    "strconv": "is not implemented yet",
    "partition": "is not implemented yet",
}


def intrinsic(name: str, *parameters: str, minimum: int = 0) -> Callable[..., Any]:
    """Register a VBA function under ``name`` with VBA's parameter names."""

    def register(function: Callable[..., object]) -> Callable[..., object]:
        def dispatch(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
            laid: list[object] = list(args) + [MISSING] * max(0, len(parameters) - len(args))
            if len(args) > len(parameters):
                raise error(450, f"{name} takes {len(parameters)} arguments")
            for key, value in named.items():
                for index, parameter in enumerate(parameters):
                    if parameter.lower() == key.lower():
                        laid[index] = value
                        break
                else:
                    raise error(448, f"{name} has no argument named {key}")
            # How many were written, not how many are non-Missing: an
            # argument whose value is Missing was still passed, which is
            # the whole point of IsMissing.
            supplied = len(args) + len(named)
            if supplied < minimum:
                raise error(449, f"{name} needs {minimum} argument{'s' if minimum != 1 else ''}")
            return function(interpreter, *laid[: len(parameters)])

        INTRINSICS[name.lower()] = dispatch
        return function

    return register


def _unimplemented(name: str, reason: str) -> Intrinsic:
    def refuse(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
        raise VBAUnsupportedError(f"{name} {reason}, so pyOpenVBA does not implement it")

    return refuse


for _name, _reason in UNIMPLEMENTED.items():
    INTRINSICS[_name] = _unimplemented(_name, _reason)


# --- conversion ---------------------------------------------------------------------


@intrinsic("CBool", "Expression", minimum=1)
def vba_cbool(interpreter: Interpreter, value: object) -> object:
    return to_bool(value)


def _convertible(value: object) -> object:
    """An Error as the number conversion functions read it: its number.

    CLng(CVErr(2042)) is 2042, and so are CInt, CByte, CSng, CDbl and CCur
    of it, measured in Excel's VBA; arithmetic on the same Variant, Int
    and Val are still a type mismatch.
    """
    return VBAInt(value.number, "Long") if isinstance(value, VBAErrorValue) else value


@intrinsic("CByte", "Expression", minimum=1)
def vba_cbyte(interpreter: Interpreter, value: object) -> object:
    return to_integer(_convertible(value), "Byte")


@intrinsic("CInt", "Expression", minimum=1)
def vba_cint(interpreter: Interpreter, value: object) -> object:
    return to_integer(_convertible(value), "Integer")


@intrinsic("CLng", "Expression", minimum=1)
def vba_clng(interpreter: Interpreter, value: object) -> object:
    return to_integer(_convertible(value), "Long")


@intrinsic("CLngLng", "Expression", minimum=1)
def vba_clnglng(interpreter: Interpreter, value: object) -> object:
    return to_integer(value, "LongLong")


@intrinsic("CSng", "Expression", minimum=1)
def vba_csng(interpreter: Interpreter, value: object) -> object:
    return VBASingle(float(to_number(_convertible(value))))


@intrinsic("CDbl", "Expression", minimum=1)
def vba_cdbl(interpreter: Interpreter, value: object) -> object:
    return float(to_number(_convertible(value)))


@intrinsic("CCur", "Expression", minimum=1)
def vba_ccur(interpreter: Interpreter, value: object) -> object:
    return VBACurrency(to_number(_convertible(value)))


@intrinsic("CDec", "Expression", minimum=1)
def vba_cdec(interpreter: Interpreter, value: object) -> object:
    return VBADecimal(str(to_number(value)))


@intrinsic("CStr", "Expression", minimum=1)
def vba_cstr(interpreter: Interpreter, value: object) -> object:
    return to_text(value)


@intrinsic("CDate", "Date", minimum=1)
def vba_cdate(interpreter: Interpreter, value: object) -> object:
    return to_date(value)


@intrinsic("CVDate", "Expression", minimum=1)
def vba_cvdate(interpreter: Interpreter, value: object) -> object:
    return to_date(value)


@intrinsic("CVar", "Expression", minimum=1)
def vba_cvar(interpreter: Interpreter, value: object) -> object:
    return value


@intrinsic("CVErr", "ErrorNumber", minimum=1)
def vba_cverr(interpreter: Interpreter, value: object) -> object:
    return VBAErrorValue(int(to_integer(value, "Long")))


@intrinsic("Val", "String", minimum=1)
def vba_val(interpreter: Interpreter, value: object) -> object:
    """Val takes the leading number and stops, and always gives a Double."""
    if isinstance(value, VBAErrorValue):
        # CStr spells an Error as "Error 2042"; Val refuses one outright.
        raise error(ERR_TYPE_MISMATCH)
    text = to_text(value).replace(" ", "").replace("\t", "")
    if text[:2].lower() in ("&h", "&o"):
        for stop in range(len(text), 2, -1):
            try:
                return float(int(text[2:stop], 16 if text[1] in "hH" else 8))
            except ValueError:
                continue
        return 0.0
    for stop in range(len(text), 0, -1):
        try:
            return float(text[:stop])
        except ValueError:
            continue
    return 0.0


@intrinsic("Str", "Number", minimum=1)
def vba_str(interpreter: Interpreter, value: object) -> object:
    """Str keeps a leading space where the number is not negative."""
    text = number_text(to_number(value))
    return text if text.startswith("-") else f" {text}"


def _unsigned(value: object) -> int:
    """A negative number as the bits of its own width, which is what
    Hex and Oct print: Hex(-1) is FFFF for an Integer and FFFFFFFF for
    a Long."""
    whole = int(to_integer(value, "LongLong"))
    if whole >= 0:
        return whole
    width = {"Byte": 8, "Integer": 16, "Long": 32}.get(type_name(value), 0)
    if not width:
        width = 16 if -32768 <= whole else (32 if whole >= -(1 << 31) else 64)
    return whole + (1 << width)


@intrinsic("Hex", "Number", minimum=1)
def vba_hex(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    return format(_unsigned(value), "X")


@intrinsic("Oct", "Number", minimum=1)
def vba_oct(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    return format(_unsigned(value), "o")


# --- maths ----------------------------------------------------------------------------


@intrinsic("Abs", "Number", minimum=1)
def vba_abs(interpreter: Interpreter, value: object) -> object:
    """Abs keeps the width it was given: Abs of an Integer is an Integer."""
    if value is NULL:
        return NULL
    number = to_number(value)
    if number >= 0:
        return number
    if isinstance(number, VBAInt):
        return VBAInt(-int(number), number.kind)
    if isinstance(number, VBASingle):
        return VBASingle(-float(number))
    return -number


@intrinsic("Sgn", "Number", minimum=1)
def vba_sgn(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    number = to_number(value)
    return VBAInt((number > 0) - (number < 0), "Integer")


@intrinsic("Int", "Number", minimum=1)
def vba_int(interpreter: Interpreter, value: object) -> object:
    """Int rounds toward minus infinity; Fix rounds toward zero."""
    if value is NULL:
        return NULL
    number = to_number(value)
    if isinstance(number, int):
        return number
    if isinstance(number, Decimal):
        return VBADecimal(math.floor(number))
    return float(math.floor(number))


@intrinsic("Fix", "Number", minimum=1)
def vba_fix(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    number = to_number(value)
    if isinstance(number, int):
        return number
    if isinstance(number, Decimal):
        return VBADecimal(int(number))
    return float(int(number))


@intrinsic("Round", "Number", "NumDigitsAfterDecimal", minimum=1)
def vba_round(interpreter: Interpreter, value: object, digits: object) -> object:
    if value is NULL:
        return NULL
    places = 0 if digits is MISSING else int(to_integer(digits, "Long"))
    number = to_number(value)
    if isinstance(number, int):
        return number
    return round_half_even(number, places)


@intrinsic("Sqr", "Number", minimum=1)
def vba_sqr(interpreter: Interpreter, value: object) -> object:
    number = float(to_number(value))
    if number < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return math.sqrt(number)


@intrinsic("Exp", "Number", minimum=1)
def vba_exp(interpreter: Interpreter, value: object) -> object:
    try:
        return math.exp(float(to_number(value)))
    except OverflowError:
        raise error(6) from None


@intrinsic("Log", "Number", minimum=1)
def vba_log(interpreter: Interpreter, value: object) -> object:
    number = float(to_number(value))
    if number <= 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return math.log(number)


@intrinsic("Sin", "Number", minimum=1)
def vba_sin(interpreter: Interpreter, value: object) -> object:
    return math.sin(float(to_number(value)))


@intrinsic("Cos", "Number", minimum=1)
def vba_cos(interpreter: Interpreter, value: object) -> object:
    return math.cos(float(to_number(value)))


@intrinsic("Tan", "Number", minimum=1)
def vba_tan(interpreter: Interpreter, value: object) -> object:
    return math.tan(float(to_number(value)))


@intrinsic("Atn", "Number", minimum=1)
def vba_atn(interpreter: Interpreter, value: object) -> object:
    return math.atan(float(to_number(value)))


@intrinsic("Rnd", "Number")
def vba_rnd(interpreter: Interpreter, value: object) -> object:
    """VBA's own generator, so a seeded run gives Excel's numbers."""
    state = getattr(interpreter, "_rnd_state", 0x50000)
    if value is not MISSING:
        number = float(to_number(value))
        if number < 0:
            state = int(abs(number) * (1 << 24)) & 0xFFFFFF
            state = (state * 1140671485 + 12820163) & 0xFFFFFF
            interpreter._rnd_state = state  # type: ignore[attr-defined]
            return VBASingle(state / (1 << 24))
        if number == 0:
            return VBASingle(state / (1 << 24))
    state = (state * 1140671485 + 12820163) & 0xFFFFFF
    interpreter._rnd_state = state  # type: ignore[attr-defined]
    return VBASingle(state / (1 << 24))


@intrinsic("Randomize", "Number")
def vba_randomize(interpreter: Interpreter, value: object) -> object:
    seed = int(time.time() * 1000) & 0xFFFFFF if value is MISSING else int(to_number(value)) & 0xFFFFFF
    interpreter._rnd_state = seed  # type: ignore[attr-defined]
    return EMPTY


# --- strings ----------------------------------------------------------------------------


@intrinsic("Len", "Expression", minimum=1)
def vba_len(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    if isinstance(value, VBAObject):
        value = value.vba_value()
    if value is EMPTY:
        return VBAInt(0, "Long")
    return VBAInt(len(to_text(value)), "Long")


@intrinsic("LenB", "Expression", minimum=1)
def vba_lenb(interpreter: Interpreter, value: object) -> object:
    if value is NULL:
        return NULL
    return VBAInt(len(to_text(value)) * 2, "Long")


@intrinsic("Left", "String", "Length", minimum=2)
def vba_left(interpreter: Interpreter, text: object, length: object) -> object:
    if text is NULL:
        return NULL
    count = int(to_integer(length, "Long"))
    if count < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return to_text(text)[:count]


@intrinsic("Right", "String", "Length", minimum=2)
def vba_right(interpreter: Interpreter, text: object, length: object) -> object:
    if text is NULL:
        return NULL
    count = int(to_integer(length, "Long"))
    if count < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    body = to_text(text)
    return body[max(0, len(body) - count) :] if count else ""


@intrinsic("Mid", "String", "Start", "Length", minimum=2)
def vba_mid(interpreter: Interpreter, text: object, start: object, length: object) -> object:
    if text is NULL:
        return NULL
    body = to_text(text)
    first = int(to_integer(start, "Long"))
    if first < 1:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    if length is MISSING:
        return body[first - 1 :]
    count = int(to_integer(length, "Long"))
    if count < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return body[first - 1 : first - 1 + count]


@intrinsic("LTrim", "String", minimum=1)
def vba_ltrim(interpreter: Interpreter, text: object) -> object:
    return NULL if text is NULL else to_text(text).lstrip(" ")


@intrinsic("RTrim", "String", minimum=1)
def vba_rtrim(interpreter: Interpreter, text: object) -> object:
    return NULL if text is NULL else to_text(text).rstrip(" ")


@intrinsic("Trim", "String", minimum=1)
def vba_trim(interpreter: Interpreter, text: object) -> object:
    return NULL if text is NULL else to_text(text).strip(" ")


@intrinsic("UCase", "String", minimum=1)
def vba_ucase(interpreter: Interpreter, text: object) -> object:
    return NULL if text is NULL else to_text(text).upper()


@intrinsic("LCase", "String", minimum=1)
def vba_lcase(interpreter: Interpreter, text: object) -> object:
    return NULL if text is NULL else to_text(text).lower()


@intrinsic("Space", "Number", minimum=1)
def vba_space(interpreter: Interpreter, count: object) -> object:
    times = int(to_integer(count, "Long"))
    if times < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return " " * times


@intrinsic("String", "Number", "Character", minimum=2)
def vba_string(interpreter: Interpreter, count: object, character: object) -> object:
    times = int(to_integer(count, "Long"))
    if times < 0:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    if isinstance(character, str):
        if not character:
            raise error(ERR_INVALID_PROCEDURE_CALL)
        return character[0] * times
    return chr(int(to_integer(character, "Long")) & 0xFF) * times


@intrinsic("StrReverse", "String", minimum=1)
def vba_strreverse(interpreter: Interpreter, text: object) -> object:
    return to_text(text)[::-1]


@intrinsic("Asc", "String", minimum=1)
def vba_asc(interpreter: Interpreter, text: object) -> object:
    body = to_text(text)
    if not body:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    code = ord(body[0])
    return VBAInt(code if code <= 0x7FFF else code - 0x10000, "Integer")


@intrinsic("AscW", "String", minimum=1)
def vba_ascw(interpreter: Interpreter, text: object) -> object:
    body = to_text(text)
    if not body:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return VBAInt(ord(body[0]), "Integer")


@intrinsic("Chr", "CharCode", minimum=1)
def vba_chr(interpreter: Interpreter, code: object) -> object:
    number = int(to_integer(code, "Long"))
    if not -32768 <= number <= 65535:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    return chr(number & 0xFFFF if number < 0 else number)


@intrinsic("ChrW", "CharCode", minimum=1)
def vba_chrw(interpreter: Interpreter, code: object) -> object:
    number = int(to_integer(code, "Long"))
    return chr(number & 0xFFFF if number < 0 else number)


@intrinsic("InStr", "Start", "String1", "String2", "Compare")
def vba_instr(
    interpreter: Interpreter, first: object, second: object, third: object, compare: object
) -> object:
    """InStr's first argument is the start only when three are given."""
    if third is MISSING and compare is MISSING:
        start, haystack, needle, mode = 1, first, second, MISSING
    else:
        start, haystack, needle, mode = int(to_integer(first, "Long")), second, third, compare
    if haystack is NULL or needle is NULL:
        return NULL
    if start < 1:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    body = "" if haystack is EMPTY or haystack is MISSING else to_text(haystack)
    wanted = "" if needle is EMPTY or needle is MISSING else to_text(needle)
    if vba_text_mode(interpreter, mode):
        body, wanted = body.lower(), wanted.lower()
    if start > len(body):
        return VBAInt(0, "Long")
    return VBAInt(body.find(wanted, int(start) - 1) + 1, "Long")


@intrinsic("InStrRev", "StringCheck", "StringMatch", "Start", "Compare", minimum=2)
def vba_instrrev(
    interpreter: Interpreter, haystack: object, needle: object, start: object, compare: object
) -> object:
    if haystack is NULL or needle is NULL:
        return NULL
    body = to_text(haystack)
    wanted = to_text(needle)
    if vba_text_mode(interpreter, compare):
        body, wanted = body.lower(), wanted.lower()
    end = len(body) if start is MISSING else int(to_integer(start, "Long"))
    if end == -1:
        end = len(body)
    return VBAInt(body.rfind(wanted, 0, end) + 1, "Long")


def vba_text_mode(interpreter: Interpreter, compare: object) -> bool:
    if compare is MISSING:
        return interpreter.option_compare_text
    return int(to_integer(compare, "Long")) == 1


@intrinsic("StrComp", "String1", "String2", "Compare", minimum=2)
def vba_strcomp(interpreter: Interpreter, first: object, second: object, compare: object) -> object:
    if first is NULL or second is NULL:
        return NULL
    left = to_text(first)
    right = to_text(second)
    if vba_text_mode(interpreter, compare):
        left, right = left.upper(), right.upper()
    return VBAInt((left > right) - (left < right), "Integer")


@intrinsic("Replace", "Expression", "Find", "Replace", "Start", "Count", "Compare", minimum=3)
def vba_replace(
    interpreter: Interpreter,
    text: object,
    find: object,
    put: object,
    start: object,
    count: object,
    compare: object,
) -> object:
    if text is NULL:
        return NULL
    body = to_text(text)
    wanted = to_text(find)
    replacement = to_text(put)
    first = 1 if start is MISSING else int(to_integer(start, "Long"))
    limit = -1 if count is MISSING else int(to_integer(count, "Long"))
    if first < 1 or limit < -1:
        raise error(ERR_INVALID_PROCEDURE_CALL)
    body = body[first - 1 :]
    if not wanted:
        return body
    if vba_text_mode(interpreter, compare):
        out: list[str] = []
        low = body.lower()
        target = wanted.lower()
        at = 0
        made = 0
        while limit == -1 or made < limit:
            found = low.find(target, at)
            if found < 0:
                break
            out.append(body[at:found])
            out.append(replacement)
            at = found + len(wanted)
            made += 1
        out.append(body[at:])
        return "".join(out)
    return body.replace(wanted, replacement) if limit == -1 else body.replace(wanted, replacement, limit)


@intrinsic("Split", "Expression", "Delimiter", "Limit", "Compare", minimum=1)
def vba_split(
    interpreter: Interpreter, text: object, delimiter: object, limit: object, compare: object
) -> object:
    body = to_text(text)
    separator = " " if delimiter is MISSING else to_text(delimiter)
    count = -1 if limit is MISSING else int(to_integer(limit, "Long"))
    if not body:
        # Splitting nothing gives an array with no elements at all,
        # whose UBound is -1.
        pieces = []
    elif not separator:
        pieces = [body]
    elif count == -1:
        pieces = body.split(separator)
    else:
        pieces = body.split(separator, count - 1) if count > 0 else []
    return VBAArray([(0, len(pieces) - 1)] if pieces else [(0, -1)], items=list(pieces))


@intrinsic("Join", "SourceArray", "Delimiter", minimum=1)
def vba_join(interpreter: Interpreter, source: object, delimiter: object) -> object:
    if not isinstance(source, VBAArray):
        raise error(ERR_TYPE_MISMATCH, "Join needs an array")
    separator = " " if delimiter is MISSING else to_text(delimiter)
    return separator.join(to_text(item) for item in source.elements())


@intrinsic("Filter", "SourceArray", "Match", "Include", "Compare", minimum=2)
def vba_filter(
    interpreter: Interpreter, source: object, match: object, include: object, compare: object
) -> object:
    if not isinstance(source, VBAArray):
        raise error(ERR_TYPE_MISMATCH, "Filter needs an array")
    wanted = to_text(match)
    keep = True if include is MISSING else to_bool(include)
    text_mode = vba_text_mode(interpreter, compare)
    found: list[object] = []
    for item in source.elements():
        body = to_text(item)
        hit = (wanted.lower() in body.lower()) if text_mode else (wanted in body)
        if hit == keep:
            found.append(item)
    return VBAArray([(0, len(found) - 1)] if found else [(0, -1)], items=found)


@intrinsic("Format", "Expression", "Format", "FirstDayOfWeek", "FirstWeekOfYear", minimum=1)
def vba_format(
    interpreter: Interpreter, value: object, pattern: object, first_day: object, first_week: object
) -> object:
    """Format, shared with the Access engine: the same function in Office."""
    from pyopenvba.access._format import format_value

    if value is NULL:
        return NULL
    if pattern is MISSING:
        return to_text(value)
    return format_value(_for_format(value), to_text(pattern))


def _for_format(value: object) -> object:
    if isinstance(value, VBADate):
        return value.to_datetime()
    if isinstance(value, VBAInt):
        return int(value)
    if isinstance(value, VBASingle):
        return float(value)
    return value


@intrinsic("FormatNumber", "Expression", "NumDigitsAfterDecimal", minimum=1)
def vba_format_number(interpreter: Interpreter, value: object, digits: object) -> object:
    places = 2 if digits is MISSING else int(to_integer(digits, "Long"))
    pattern = "#,##0" + ("." + "0" * places if places else "")
    return vba_format(interpreter, value, pattern, MISSING, MISSING)


@intrinsic("FormatCurrency", "Expression", "NumDigitsAfterDecimal", minimum=1)
def vba_format_currency(interpreter: Interpreter, value: object, digits: object) -> object:
    places = 2 if digits is MISSING else int(to_integer(digits, "Long"))
    pattern = "$#,##0" + ("." + "0" * places if places else "")
    return vba_format(interpreter, value, pattern, MISSING, MISSING)


@intrinsic("FormatPercent", "Expression", "NumDigitsAfterDecimal", minimum=1)
def vba_format_percent(interpreter: Interpreter, value: object, digits: object) -> object:
    places = 2 if digits is MISSING else int(to_integer(digits, "Long"))
    pattern = "0" + ("." + "0" * places if places else "") + "%"
    return vba_format(interpreter, value, pattern, MISSING, MISSING)


@intrinsic("FormatDateTime", "Date", "NamedFormat", minimum=1)
def vba_format_datetime(interpreter: Interpreter, value: object, named: object) -> object:
    which = 0 if named is MISSING else int(to_integer(named, "Long"))
    patterns = {
        0: "",
        1: "dddd, mmmm dd, yyyy",
        2: "m/d/yyyy",
        3: "h:mm:ss AM/PM",
        4: "hh:mm",
    }
    if which not in patterns:
        raise error(5)
    if which == 0:
        return to_text(to_date(value))
    return vba_format(interpreter, to_date(value), patterns[which], MISSING, MISSING)


# --- dates --------------------------------------------------------------------------------


@intrinsic("Now")
def vba_now(interpreter: Interpreter) -> object:
    return VBADate.from_datetime(interpreter.clock())


@intrinsic("Date")
def vba_date(interpreter: Interpreter) -> object:
    return VBADate.from_datetime(interpreter.clock().date())


@intrinsic("Time")
def vba_time(interpreter: Interpreter) -> object:
    return VBADate.from_datetime(interpreter.clock().time())


@intrinsic("Timer")
def vba_timer(interpreter: Interpreter) -> object:
    moment = interpreter.clock()
    seconds = moment.hour * 3600 + moment.minute * 60 + moment.second + moment.microsecond / 1e6
    return VBASingle(seconds)


@intrinsic("DateSerial", "Year", "Month", "Day", minimum=3)
def vba_dateserial(interpreter: Interpreter, year: object, month: object, day: object) -> object:
    y = int(to_integer(year, "Long"))
    m = int(to_integer(month, "Long"))
    d = int(to_integer(day, "Long"))
    if y < 100:
        y += 1900 if y >= 30 else 2000
    # Out-of-range months and days roll over, as VBA's do.
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    try:
        base = _dt.datetime(y, m, 1)
    except ValueError:
        raise error(5) from None
    return VBADate.from_datetime(base + _dt.timedelta(days=d - 1))


@intrinsic("TimeSerial", "Hour", "Minute", "Second", minimum=3)
def vba_timeserial(interpreter: Interpreter, hour: object, minute: object, second: object) -> object:
    seconds = (
        int(to_integer(hour, "Long")) * 3600
        + int(to_integer(minute, "Long")) * 60
        + int(to_integer(second, "Long"))
    )
    return VBADate(seconds / 86400.0)


@intrinsic("DateValue", "Date", minimum=1)
def vba_datevalue(interpreter: Interpreter, value: object) -> object:
    when = to_date(value)
    return VBADate(float(math.floor(when.serial)))


@intrinsic("TimeValue", "Time", minimum=1)
def vba_timevalue(interpreter: Interpreter, value: object) -> object:
    when = to_date(value)
    return VBADate(when.serial - math.floor(when.serial))


@intrinsic("Year", "Date", minimum=1)
def vba_year(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().year, "Integer")


@intrinsic("Month", "Date", minimum=1)
def vba_month(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().month, "Integer")


@intrinsic("Day", "Date", minimum=1)
def vba_day(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().day, "Integer")


@intrinsic("Hour", "Time", minimum=1)
def vba_hour(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().hour, "Integer")


@intrinsic("Minute", "Time", minimum=1)
def vba_minute(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().minute, "Integer")


@intrinsic("Second", "Time", minimum=1)
def vba_second(interpreter: Interpreter, value: object) -> object:
    return NULL if value is NULL else VBAInt(to_date(value).to_datetime().second, "Integer")


@intrinsic("Weekday", "Date", "FirstDayOfWeek", minimum=1)
def vba_weekday(interpreter: Interpreter, value: object, first: object) -> object:
    if value is NULL:
        return NULL
    start = 1 if first is MISSING or int(to_integer(first, "Long")) == 0 else int(to_integer(first, "Long"))
    # Python counts Monday as 0; VBA counts Sunday as 1.
    sunday_based = (to_date(value).to_datetime().weekday() + 1) % 7 + 1
    return VBAInt((sunday_based - start) % 7 + 1, "Integer")


@intrinsic("WeekdayName", "Weekday", "Abbreviate", "FirstDayOfWeek", minimum=1)
def vba_weekdayname(interpreter: Interpreter, number: object, short: object, first: object) -> object:
    names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    start = 1 if first is MISSING or int(to_integer(first, "Long")) == 0 else int(to_integer(first, "Long"))
    index = int(to_integer(number, "Long"))
    if not 1 <= index <= 7:
        raise error(5)
    name = names[(index - 1 + start - 1) % 7]
    return name[:3] if short is not MISSING and to_bool(short) else name


@intrinsic("MonthName", "Month", "Abbreviate", minimum=1)
def vba_monthname(interpreter: Interpreter, number: object, short: object) -> object:
    index = int(to_integer(number, "Long"))
    if not 1 <= index <= 12:
        raise error(5)
    name = _dt.date(2000, index, 1).strftime("%B")
    return name[:3] if short is not MISSING and to_bool(short) else name


_INTERVALS: Final = {"yyyy", "q", "m", "y", "d", "w", "ww", "h", "n", "s"}


@intrinsic("DateAdd", "Interval", "Number", "Date", minimum=3)
def vba_dateadd(interpreter: Interpreter, interval: object, count: object, value: object) -> object:
    unit = to_text(interval).lower()
    if unit not in _INTERVALS:
        raise error(5)
    how_many = int(to_integer(count, "Long"))
    when = to_date(value).to_datetime()
    if unit in ("yyyy", "q", "m"):
        months = how_many * (12 if unit == "yyyy" else 3 if unit == "q" else 1)
        total = when.year * 12 + when.month - 1 + months
        year, month = divmod(total, 12)
        day = min(when.day, _days_in_month(year, month + 1))
        return VBADate.from_datetime(when.replace(year=year, month=month + 1, day=day))
    seconds = {"d": 86400, "y": 86400, "w": 86400, "ww": 604800, "h": 3600, "n": 60, "s": 1}[unit]
    return VBADate.from_datetime(when + _dt.timedelta(seconds=how_many * seconds))


def _days_in_month(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (_dt.date(year, month + 1, 1) - _dt.date(year, month, 1)).days


@intrinsic("DateDiff", "Interval", "Date1", "Date2", "FirstDayOfWeek", "FirstWeekOfYear", minimum=3)
def vba_datediff(
    interpreter: Interpreter,
    interval: object,
    first: object,
    second: object,
    week_start: object,
    year_start: object,
) -> object:
    unit = to_text(interval).lower()
    if unit not in _INTERVALS:
        raise error(5)
    start = to_date(first).to_datetime()
    stop = to_date(second).to_datetime()
    if unit == "yyyy":
        return VBAInt(stop.year - start.year, "Long")
    if unit == "q":
        return VBAInt((stop.year * 4 + (stop.month - 1) // 3) - (start.year * 4 + (start.month - 1) // 3), "Long")
    if unit == "m":
        return VBAInt((stop.year * 12 + stop.month) - (start.year * 12 + start.month), "Long")
    delta = stop - start
    if unit in ("d", "y"):
        return VBAInt((stop.date() - start.date()).days, "Long")
    if unit == "w":
        return VBAInt((stop.date() - start.date()).days // 7, "Long")
    if unit == "ww":
        begin = 1 if week_start is MISSING else int(to_integer(week_start, "Long")) or 1
        return VBAInt((_week_start(stop.date(), begin) - _week_start(start.date(), begin)).days // 7, "Long")
    seconds = delta.days * 86400 + delta.seconds
    return VBAInt(seconds // {"h": 3600, "n": 60, "s": 1}[unit], "Long")


def _week_start(day: _dt.date, first: int) -> _dt.date:
    sunday_based = (day.weekday() + 1) % 7 + 1
    return day - _dt.timedelta(days=(sunday_based - first) % 7)


@intrinsic("DatePart", "Interval", "Date", "FirstDayOfWeek", "FirstWeekOfYear", minimum=2)
def vba_datepart(
    interpreter: Interpreter, interval: object, value: object, week_start: object, year_start: object
) -> object:
    unit = to_text(interval).lower()
    when = to_date(value).to_datetime()
    if unit == "yyyy":
        return VBAInt(when.year, "Integer")
    if unit == "q":
        return VBAInt((when.month - 1) // 3 + 1, "Integer")
    if unit == "m":
        return VBAInt(when.month, "Integer")
    if unit == "y":
        return VBAInt(when.timetuple().tm_yday, "Integer")
    if unit == "d":
        return VBAInt(when.day, "Integer")
    if unit == "w":
        return vba_weekday(interpreter, VBADate.from_datetime(when), week_start)
    if unit == "ww":
        return VBAInt(int(when.strftime("%U")) + 1, "Integer")
    if unit == "h":
        return VBAInt(when.hour, "Integer")
    if unit == "n":
        return VBAInt(when.minute, "Integer")
    if unit == "s":
        return VBAInt(when.second, "Integer")
    raise error(5)


# --- tests and information ------------------------------------------------------------------


@intrinsic("IsNull", "Expression", minimum=1)
def vba_isnull(interpreter: Interpreter, value: object) -> object:
    return value is NULL


@intrinsic("IsEmpty", "Expression", minimum=1)
def vba_isempty(interpreter: Interpreter, value: object) -> object:
    return value is EMPTY


@intrinsic("IsMissing", "ArgName", minimum=1)
def vba_ismissing(interpreter: Interpreter, value: object) -> object:
    return value is MISSING


@intrinsic("IsNumeric", "Expression", minimum=1)
def vba_isnumeric(interpreter: Interpreter, value: object) -> object:
    """IsNumeric is looser than CDbl: it allows the digit grouping."""
    if isinstance(value, (bool, int, float, Decimal)):
        return True
    if value is EMPTY:
        return False
    if isinstance(value, str):
        try:
            to_number(value.replace(",", ""))
        except Exception:
            return False
        return True
    return False


@intrinsic("IsDate", "Expression", minimum=1)
def vba_isdate(interpreter: Interpreter, value: object) -> object:
    if isinstance(value, VBADate):
        return True
    if isinstance(value, str):
        return parse_date_text(value) is not None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return -657434.0 <= float(value) < 2958466.0
    return False


@intrinsic("IsArray", "VarName", minimum=1)
def vba_isarray(interpreter: Interpreter, value: object) -> object:
    return isinstance(value, VBAArray)


@intrinsic("IsObject", "Expression", minimum=1)
def vba_isobject(interpreter: Interpreter, value: object) -> object:
    return isinstance(value, VBAObject) or value is NOTHING


@intrinsic("IsError", "Expression", minimum=1)
def vba_iserror(interpreter: Interpreter, value: object) -> object:
    return isinstance(value, VBAErrorValue)


@intrinsic("TypeName", "VarName", minimum=1)
def vba_typename(interpreter: Interpreter, value: object) -> object:
    return type_name(value)


@intrinsic("VarType", "VarName", minimum=1)
def vba_vartype(interpreter: Interpreter, value: object) -> object:
    from pyopenvba.interpreter._values import VAR_TYPES

    if isinstance(value, VBAArray):
        return VBAInt(8192 + VAR_TYPES.get(value.element_type, 12), "Integer")
    if isinstance(value, VBAObject):
        return VBAInt(9, "Integer")
    return VBAInt(VAR_TYPES.get(type_name(value), 12), "Integer")


# --- arrays and choice -----------------------------------------------------------------------


def vba_array_dispatch(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
    """Array() takes any number of arguments, so it binds its own."""
    return VBAArray([(0, len(args) - 1)] if args else [(0, -1)], items=list(args))


INTRINSICS["array"] = vba_array_dispatch


@intrinsic("LBound", "ArrayName", "Dimension", minimum=1)
def vba_lbound(interpreter: Interpreter, array: object, dimension: object) -> object:
    if not isinstance(array, VBAArray):
        raise error(ERR_TYPE_MISMATCH, "LBound needs an array")
    which = 1 if dimension is MISSING else int(to_integer(dimension, "Long"))
    if not 1 <= which <= array.dimensions:
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
    return VBAInt(array.bounds[which - 1][0], "Long")


@intrinsic("UBound", "ArrayName", "Dimension", minimum=1)
def vba_ubound(interpreter: Interpreter, array: object, dimension: object) -> object:
    if not isinstance(array, VBAArray):
        raise error(ERR_TYPE_MISMATCH, "UBound needs an array")
    which = 1 if dimension is MISSING else int(to_integer(dimension, "Long"))
    if not 1 <= which <= array.dimensions:
        raise error(ERR_SUBSCRIPT_OUT_OF_RANGE)
    return VBAInt(array.bounds[which - 1][1], "Long")


def vba_iif(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
    """IIf evaluates both arms, as VBA's does, so both can still fail."""
    if len(args) != 3:
        raise error(450, "IIf takes three arguments")
    return args[1] if to_bool(args[0]) else args[2]


INTRINSICS["iif"] = vba_iif


def vba_choose(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
    if not args:
        raise error(449)
    index = int(to_integer(args[0], "Long"))
    if 1 <= index < len(args):
        return args[index]
    return NULL


INTRINSICS["choose"] = vba_choose


def vba_switch(interpreter: Interpreter, args: list[object], named: dict[str, object]) -> object:
    if len(args) % 2:
        raise error(450, "Switch takes pairs")
    for index in range(0, len(args), 2):
        if to_bool(args[index]):
            return args[index + 1]
    return NULL


INTRINSICS["switch"] = vba_switch


@intrinsic("RGB", "Red", "Green", "Blue", minimum=3)
def vba_rgb(interpreter: Interpreter, red: object, green: object, blue: object) -> object:
    def channel(value: object) -> int:
        number = int(to_integer(value, "Long"))
        return 255 if number > 255 else max(0, number)

    return VBAInt(channel(red) + channel(green) * 256 + channel(blue) * 65536, "Long")


# --- dialogs and objects -----------------------------------------------------------------------


@intrinsic("MsgBox", "Prompt", "Buttons", "Title", "HelpFile", "Context", minimum=1)
def vba_msgbox(
    interpreter: Interpreter,
    prompt: object,
    buttons: object,
    title: object,
    help_file: object,
    context: object,
) -> object:
    style = 0 if buttons is MISSING else int(to_integer(buttons, "Long"))
    caption = to_text(title) if title is not MISSING else "Microsoft Excel"
    return interpreter.show_message(to_text(prompt), style, caption)


@intrinsic("InputBox", "Prompt", "Title", "Default", "XPos", "YPos", "HelpFile", "Context", minimum=1)
def vba_inputbox(
    interpreter: Interpreter,
    prompt: object,
    title: object,
    default: object,
    x: object,
    y: object,
    help_file: object,
    context: object,
) -> object:
    return interpreter.ask_for_input(
        to_text(prompt),
        to_text(title) if title is not MISSING else "",
        to_text(default) if default is not MISSING else "",
    )


@intrinsic("CreateObject", "Class", "ServerName", minimum=1)
def vba_createobject(interpreter: Interpreter, class_name: object, server: object) -> object:
    return interpreter.create_com(to_text(class_name))


# --- finance ------------------------------------------------------------------------------------


def _annuity(rate: float, periods: float, present: float, future: float, due: int) -> float:
    if rate == 0:
        return -(present + future) / periods
    growth = (1 + rate) ** periods
    return -(future + present * growth) / ((1 + rate * due) * (growth - 1) / rate)


@intrinsic("Pmt", "Rate", "NPer", "PV", "FV", "Due", minimum=3)
def vba_pmt(
    interpreter: Interpreter, rate: object, periods: object, present: object, future: object, due: object
) -> object:
    return _annuity(
        float(to_number(rate)),
        float(to_number(periods)),
        float(to_number(present)),
        0.0 if future is MISSING else float(to_number(future)),
        0 if due is MISSING else (1 if to_bool(due) else 0),
    )


@intrinsic("FV", "Rate", "NPer", "Pmt", "PV", "Due", minimum=3)
def vba_fv(
    interpreter: Interpreter, rate: object, periods: object, payment: object, present: object, due: object
) -> object:
    r = float(to_number(rate))
    n = float(to_number(periods))
    p = float(to_number(payment))
    pv = 0.0 if present is MISSING else float(to_number(present))
    at_start = 0 if due is MISSING else (1 if to_bool(due) else 0)
    if r == 0:
        return -(pv + p * n)
    growth = (1 + r) ** n
    return -(pv * growth + p * (1 + r * at_start) * (growth - 1) / r)


@intrinsic("PV", "Rate", "NPer", "Pmt", "FV", "Due", minimum=3)
def vba_pv(
    interpreter: Interpreter, rate: object, periods: object, payment: object, future: object, due: object
) -> object:
    r = float(to_number(rate))
    n = float(to_number(periods))
    p = float(to_number(payment))
    fv = 0.0 if future is MISSING else float(to_number(future))
    at_start = 0 if due is MISSING else (1 if to_bool(due) else 0)
    if r == 0:
        return -(fv + p * n)
    growth = (1 + r) ** n
    return -(fv + p * (1 + r * at_start) * (growth - 1) / r) / growth


@intrinsic("NPer", "Rate", "Pmt", "PV", "FV", "Due", minimum=3)
def vba_nper(
    interpreter: Interpreter, rate: object, payment: object, present: object, future: object, due: object
) -> object:
    r = float(to_number(rate))
    p = float(to_number(payment))
    pv = float(to_number(present))
    fv = 0.0 if future is MISSING else float(to_number(future))
    at_start = 0 if due is MISSING else (1 if to_bool(due) else 0)
    if r == 0:
        if p == 0:
            raise error(5)
        return -(pv + fv) / p
    adjusted = p * (1 + r * at_start) / r
    return math.log((adjusted - fv) / (adjusted + pv)) / math.log(1 + r)


@intrinsic("SLN", "Cost", "Salvage", "Life", minimum=3)
def vba_sln(interpreter: Interpreter, cost: object, salvage: object, life: object) -> object:
    span = float(to_number(life))
    if span == 0:
        raise error(5)
    return (float(to_number(cost)) - float(to_number(salvage))) / span


@intrinsic("SYD", "Cost", "Salvage", "Life", "Period", minimum=4)
def vba_syd(interpreter: Interpreter, cost: object, salvage: object, life: object, period: object) -> object:
    span = float(to_number(life))
    at = float(to_number(period))
    if span <= 0 or at <= 0 or at > span:
        raise error(5)
    return (float(to_number(cost)) - float(to_number(salvage))) * (span - at + 1) * 2 / (span * (span + 1))


@intrinsic("DDB", "Cost", "Salvage", "Life", "Period", "Factor", minimum=4)
def vba_ddb(
    interpreter: Interpreter, cost: object, salvage: object, life: object, period: object, factor: object
) -> object:
    start = float(to_number(cost))
    floor = float(to_number(salvage))
    span = float(to_number(life))
    at = float(to_number(period))
    rate = 2.0 if factor is MISSING else float(to_number(factor))
    if span <= 0 or at <= 0:
        raise error(5)
    book = start
    written = 0.0
    for _ in range(int(at)):
        written = min(book * rate / span, book - floor)
        written = max(0.0, written)
        book -= written
    return written
