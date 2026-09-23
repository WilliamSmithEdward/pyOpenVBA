"""Text functions.

A number given where text is wanted becomes the fifteen-digit text ``&``
makes of it, so ``LEN(1234.5)`` is 6 and ``LEFT(1234,2)`` is ``"12"``.
TEXT, DOLLAR and FIXED format through the same number format code Excel
shows a cell with, :func:`~pyopenvba.formula._calc.host.format_value`.

CHAR and CODE count in Windows-1252, Excel's character set on Windows, so
``CHAR(128)`` is the euro sign; UNICHAR and UNICODE count in Unicode.
"""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP
from urllib.parse import quote

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import rounded
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.registry import A, R, V, function
from pyopenvba.formula._calc.values import (
    MAX_TEXT,
    NA,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Reference,
    Scalar,
    Value,
    scalar_text,
    text_to_number,
)
from pyopenvba.formula._calc.host import format_value
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.cells import UnsupportedFormulaError


def _length(context: Context, value: Scalar | None, default: int) -> int:
    """A count of characters: cut to a whole number, never negative."""
    if value is None:
        return default
    count = context.number(value)
    if count < 0:
        raise ExcelError(VALUE)
    return int(count)


def _checked(text: str) -> str:
    if len(text) > MAX_TEXT:
        raise ExcelError(VALUE)
    return text


@function("LEFT", V, V, minimum=1)
def LEFT(context: Context, text: Scalar, count: Scalar | None = None) -> Value:
    return context.text(text)[: _length(context, count, 1)]


@function("RIGHT", V, V, minimum=1)
def RIGHT(context: Context, text: Scalar, count: Scalar | None = None) -> Value:
    source = context.text(text)
    length = _length(context, count, 1)
    return source[len(source) - length :] if length else ""


@function("MID", V, V, V)
def MID(context: Context, text: Scalar, start: Scalar, count: Scalar) -> Value:
    source = context.text(text)
    first = context.number(start)
    if first < 1:
        return VALUE
    length = _length(context, count, 0)
    begin = int(first) - 1
    return source[begin : begin + length]


@function("LEN", V)
def LEN(context: Context, text: Scalar) -> Value:
    return float(len(context.text(text)))


@function("UPPER", V)
def UPPER(context: Context, text: Scalar) -> Value:
    return "".join(char if char == "\u00df" else char.upper() for char in context.text(text))


@function("LOWER", V)
def LOWER(context: Context, text: Scalar) -> Value:
    return "".join(_lower(char) for char in context.text(text))


def _lower(char: str) -> str:
    """One character in lower case, as one character: Windows maps the
    dotted capital I to a plain i where Unicode adds a combining dot."""
    lowered = char.lower()
    return lowered if len(lowered) == 1 else lowered[0]


@function("PROPER", V)
def PROPER(context: Context, text: Scalar) -> Value:
    out: list[str] = []
    previous_letter = False
    for char in context.text(text):
        out.append(char.lower() if previous_letter else char.upper() if char != "\u00df" else char)
        previous_letter = char.isalpha()
    return "".join(out)


@function("TRIM", V)
def TRIM(context: Context, text: Scalar) -> Value:
    return " ".join(part for part in context.text(text).split(" ") if part)


@function("CLEAN", V)
def CLEAN(context: Context, text: Scalar) -> Value:
    return "".join(char for char in context.text(text) if ord(char) >= 32)


def _start(context: Context, value: Scalar | None, within: str) -> int:
    begin = 1 if value is None else context.number(value)
    if begin < 1 or begin > len(within) + 1:
        raise ExcelError(VALUE)
    return int(begin) - 1


@function("FIND", V, V, V, minimum=2)
def FIND(context: Context, find: Scalar, within: Scalar, start: Scalar | None = None) -> Value:
    needle = context.text(find)
    haystack = context.text(within)
    begin = _start(context, start, haystack)
    found = haystack.find(needle, begin)
    return VALUE if found < 0 else float(found + 1)


def _search_pattern(needle: str) -> re.Pattern[str]:
    parts: list[str] = []
    index = 0
    while index < len(needle):
        char = needle[index]
        if char == "~" and index + 1 < len(needle) and needle[index + 1] in "*?~":
            parts.append(re.escape(needle[index + 1]))
            index += 2
            continue
        parts.append(".*?" if char == "*" else "." if char == "?" else re.escape(char))
        index += 1
    return re.compile("".join(parts), re.IGNORECASE | re.DOTALL)


@function("SEARCH", V, V, V, minimum=2)
def SEARCH(context: Context, find: Scalar, within: Scalar, start: Scalar | None = None) -> Value:
    needle = context.text(find)
    haystack = context.text(within)
    begin = _start(context, start, haystack)
    found = _search_pattern(needle).search(haystack, begin)
    return VALUE if found is None else float(found.start() + 1)


@function("SUBSTITUTE", V, V, V, V, minimum=3)
def SUBSTITUTE(context: Context, text: Scalar, old: Scalar, new: Scalar, instance: Scalar | None = None) -> Value:
    source = context.text(text)
    before = context.text(old)
    after = context.text(new)
    if not before:
        return source
    if instance is None:
        return _checked(source.replace(before, after))
    which = context.number(instance)
    if which < 1:
        return VALUE
    position = -1
    for _ in range(int(which)):
        position = source.find(before, position + 1)
        if position < 0:
            return source
    return _checked(source[:position] + after + source[position + len(before) :])


@function("REPLACE", V, V, V, V)
def REPLACE(context: Context, text: Scalar, start: Scalar, count: Scalar, new: Scalar) -> Value:
    source = context.text(text)
    first = context.number(start)
    if first < 1:
        return VALUE
    length = _length(context, count, 0)
    begin = int(first) - 1
    return _checked(source[:begin] + context.text(new) + source[begin + length :])


@function("CONCATENATE", V, maximum=255)
def CONCATENATE(context: Context, *parts: Scalar) -> Value:
    return _checked("".join(context.text(part) for part in parts))


def _pieces(context: Context, value: Value) -> list[Scalar]:
    if isinstance(value, (Reference, Array)):
        return list(matrix(context, value).items())
    return [value]


@function("CONCAT", R, maximum=254)
def CONCAT(context: Context, *parts: Value) -> Value:
    return _checked("".join(context.text(piece) for part in parts for piece in _pieces(context, part)))


@function("TEXTJOIN", R, V, R, maximum=252)
def TEXTJOIN(context: Context, delimiter: Value, ignore_empty: Scalar, *parts: Value) -> Value:
    separators = [context.text(piece) for piece in _pieces(context, delimiter)] or [""]
    skip = context.logical(ignore_empty)
    texts: list[str] = []
    for part in parts:
        for piece in _pieces(context, part):
            text = context.text(piece)
            if skip and text == "":
                continue
            texts.append(text)
    out: list[str] = []
    for index, text in enumerate(texts):
        if index:
            out.append(separators[(index - 1) % len(separators)])
        out.append(text)
    return _checked("".join(out))


@function("REPT", V, V)
def REPT(context: Context, text: Scalar, times: Scalar) -> Value:
    source = context.text(text)
    count = context.number(times)
    if count < 0:
        return VALUE
    if len(source) * int(count) > MAX_TEXT:
        return VALUE
    return source * int(count)


@function("EXACT", V, V)
def EXACT(context: Context, first: Scalar, second: Scalar) -> Value:
    return context.text(first) == context.text(second)


@function("TEXT", V, V)
def TEXT(context: Context, value: Scalar, code: Scalar) -> Value:
    if isinstance(value, CellError):
        return value
    fmt = context.text(code)
    subject: object = value
    if isinstance(value, str):
        number = text_to_number(value, context.today, epoch_1904=context.epoch_1904)
        subject = value if number is None else number
    elif isinstance(value, Empty):
        subject = 0.0
    return format_value(subject, fmt, epoch_1904=context.epoch_1904)


@function("VALUE", V)
def VALUE_(context: Context, text: Scalar) -> Value:
    if isinstance(text, bool):
        return VALUE
    return context.number(text)


@function("NUMBERVALUE", V, V, V, minimum=1)
def NUMBERVALUE(context: Context, text: Scalar, decimal: Scalar | None = None, group: Scalar | None = None) -> Value:
    if isinstance(text, (float, bool)):
        return context.number(text)
    source = "".join(context.text(text).split())
    point = "." if decimal is None else context.text(decimal)[:1]
    # Left out, the group separator is the one the decimal mark is not.
    separator = ("." if point == "," else ",") if group is None else context.text(group)[:1]
    if not source:
        return 0.0
    if not point or point == separator:
        return VALUE
    percent = len(source) - len(source.rstrip("%"))
    source = source.rstrip("%")
    whole, _, fraction = source.partition(point)
    if point in fraction:
        return VALUE
    if separator:
        if separator in fraction:
            return VALUE
        whole = whole.replace(separator, "")
    try:
        number = float(f"{whole}.{fraction}" if fraction else whole)
    except ValueError:
        return VALUE
    return number / 100**percent


_CP1252 = "cp1252"
#: The codes Windows-1252 leaves undefined.
_UNMAPPED = frozenset({0x81, 0x8D, 0x8F, 0x90, 0x9D})


@function("CHAR", V)
def CHAR(context: Context, number: Scalar) -> Value:
    code = int(context.number(number))
    if not 1 <= code <= 255:
        return VALUE
    try:
        return bytes([code]).decode(_CP1252)
    except UnicodeDecodeError:
        # The five codes Windows-1252 leaves undefined map straight across.
        return chr(code)


@function("CODE", V)
def CODE(context: Context, text: Scalar) -> Value:
    source = context.text(text)
    if not source:
        return VALUE
    try:
        encoded = source[0].encode(_CP1252)
    except UnicodeEncodeError:
        # The five codes Windows-1252 leaves out come back as themselves,
        # as CHAR makes them; anything else it cannot hold is "?".
        return float(ord(source[0])) if ord(source[0]) in _UNMAPPED else 63.0
    return float(encoded[0])


@function("UNICHAR", V)
def UNICHAR(context: Context, number: Scalar) -> Value:
    code = int(context.number(number))
    if not 1 <= code <= 0x10FFFF or 0xD800 <= code <= 0xDFFF:
        return VALUE
    return chr(code)


@function("UNICODE", V)
def UNICODE(context: Context, text: Scalar) -> Value:
    source = context.text(text)
    if not source:
        return VALUE
    return float(ord(source[0]))


@function("T", V)
def T(context: Context, value: Scalar) -> Value:
    if isinstance(value, CellError):
        return value
    return value if isinstance(value, str) else ""


@function("N", V)
def N(context: Context, value: Scalar) -> Value:
    if isinstance(value, CellError):
        return value
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    return value if isinstance(value, float) else 0.0


def _fixed_code(decimals: int, grouping: bool) -> str:
    whole = "#,##0" if grouping else "0"
    return whole + ("." + "0" * decimals if decimals > 0 else "")


@function("DOLLAR", V, V, minimum=1)
def DOLLAR(context: Context, number: Scalar, decimals: Scalar | None = None) -> Value:
    value = context.number(number)
    places = 2 if decimals is None else context.integer(decimals)
    if places > 127:
        return VALUE
    value = rounded(value, places, ROUND_HALF_UP)
    body = _fixed_code(max(places, 0), True)
    return format_value(value, f"${body}_);(${body})", epoch_1904=context.epoch_1904).rstrip()


@function("FIXED", V, V, V, minimum=1)
def FIXED(context: Context, number: Scalar, decimals: Scalar | None = None, no_commas: Scalar | None = None) -> Value:
    value = context.number(number)
    places = 2 if decimals is None else context.integer(decimals)
    if places > 127:
        return VALUE
    grouping = not (no_commas is not None and context.logical(no_commas))
    value = rounded(value, places, ROUND_HALF_UP)
    return format_value(value, _fixed_code(max(places, 0), grouping), epoch_1904=context.epoch_1904)


@function("ASC", V)
def ASC(context: Context, text: Scalar) -> Value:
    return context.text(text)


@function("DBCS", V)
def DBCS(context: Context, text: Scalar) -> Value:
    return context.text(text)


@function("VALUETOTEXT", V, V, minimum=1)
def VALUETOTEXT(context: Context, value: Scalar, strict: Scalar | None = None) -> Value:
    quoted = strict is not None and context.number(strict) == 1
    if isinstance(value, CellError):
        return value.code
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"' if quoted else value
    return scalar_text(value)


# ----------------------------------------------------------------------
# TEXTBEFORE, TEXTAFTER and TEXTSPLIT
# ----------------------------------------------------------------------


def _delimiters(context: Context, value: Value) -> list[str]:
    """One delimiter or an array of them, any of which counts."""
    found = [context.text(item) for item in matrix(context, value).items()]
    if not found:
        raise ExcelError(VALUE)
    return found


def _folded(text: str, insensitive: bool) -> str:
    """Text to search in, one character for one, so that positions found
    in it are positions in the text."""
    if not insensitive:
        return text
    return "".join(lower if len(lower := char.lower()) == 1 else char for char in text)


def _occurrences(text: str, delimiters: list[str], insensitive: bool) -> list[tuple[int, int]]:
    """Where the delimiters occur, left to right and not overlapping, as
    (start, length); at one place the first delimiter listed wins."""
    haystack = _folded(text, insensitive)
    needles = [_folded(delimiter, insensitive) for delimiter in delimiters if delimiter]
    found: list[tuple[int, int]] = []
    index = 0
    while index < len(haystack):
        for needle in needles:
            if haystack.startswith(needle, index):
                found.append((index, len(needle)))
                index += len(needle)
                break
        else:
            index += 1
    return found


def _split_at(
    context: Context,
    text: Scalar,
    delimiter: Value,
    instance: Scalar | None,
    mode: Scalar | None,
    end: Scalar | None,
) -> tuple[str, tuple[int, int] | None]:
    """The text, and the delimiter occurrence TEXTBEFORE and TEXTAFTER cut
    at, or ``None`` when there is none."""
    source = context.text(text)
    delimiters = _delimiters(context, delimiter)
    which = 1 if instance is None or isinstance(instance, Empty) else context.integer(instance)
    insensitive = mode is not None and not isinstance(mode, Empty) and context.integer(mode) == 1
    to_end = end is not None and not isinstance(end, Empty) and context.logical(end)
    if which == 0 or (source and abs(which) > len(source)):
        raise ExcelError(VALUE)
    if all(delimiter == "" for delimiter in delimiters):
        # The empty delimiter is found at once, at the start or the end.
        return source, (0, 0) if which > 0 else (len(source), 0)
    places = _occurrences(source, delimiters, insensitive)
    if to_end:
        places = [*places, (len(source), 0)] if which > 0 else [(0, 0), *places]
    if abs(which) > len(places):
        return source, None
    return source, places[which - 1] if which > 0 else places[which]


def _missing(value: Value | None) -> Value:
    return NA if value is None or isinstance(value, Empty) else value


@function("TEXTBEFORE", V, A, V, V, V, A, minimum=2)
def TEXTBEFORE(
    context: Context,
    text: Scalar,
    delimiter: Value,
    instance: Scalar | None = None,
    mode: Scalar | None = None,
    end: Scalar | None = None,
    missing: Value | None = None,
) -> Value:
    source, place = _split_at(context, text, delimiter, instance, mode, end)
    return _missing(missing) if place is None else source[: place[0]]


@function("TEXTAFTER", V, A, V, V, V, A, minimum=2)
def TEXTAFTER(
    context: Context,
    text: Scalar,
    delimiter: Value,
    instance: Scalar | None = None,
    mode: Scalar | None = None,
    end: Scalar | None = None,
    missing: Value | None = None,
) -> Value:
    source, place = _split_at(context, text, delimiter, instance, mode, end)
    return _missing(missing) if place is None else source[place[0] + place[1] :]


def _split_text(text: str, delimiters: list[str], insensitive: bool, ignore_empty: bool) -> list[str]:
    pieces: list[str] = []
    start = 0
    for index, length in _occurrences(text, delimiters, insensitive):
        pieces.append(text[start:index])
        start = index + length
    pieces.append(text[start:])
    return [piece for piece in pieces if piece] if ignore_empty else pieces


@function("TEXTSPLIT", V, A, A, V, V, A, minimum=2)
def TEXTSPLIT(
    context: Context,
    text: Scalar,
    columns: Value,
    rows: Value | None = None,
    ignore_empty: Scalar | None = None,
    mode: Scalar | None = None,
    pad: Value | None = None,
) -> Value:
    source = context.text(text)
    across = [] if isinstance(columns, Empty) else _delimiters(context, columns)
    down = [] if rows is None or isinstance(rows, Empty) else _delimiters(context, rows)
    if not across and not down:
        return VALUE
    insensitive = mode is not None and not isinstance(mode, Empty) and context.integer(mode) == 1
    skip = ignore_empty is not None and not isinstance(ignore_empty, Empty) and context.logical(ignore_empty)
    lines = _split_text(source, down, insensitive, skip) if down else [source]
    grid: list[list[Scalar]] = [
        list(_split_text(line, across, insensitive, skip)) if across else [line] for line in lines
    ]
    grid = [line for line in grid if line] or [[""]]
    width = max(len(line) for line in grid)
    filler: Scalar = NA if pad is None or isinstance(pad, Empty) else context.first(pad)
    return Array([line + [filler] * (width - len(line)) for line in grid])


# ----------------------------------------------------------------------
# The byte-counting twins, and other names for the same function
# ----------------------------------------------------------------------

# LEFTB and the rest count bytes only where a double-byte language is the
# default; in the en-US Excel this engine follows they count characters,
# exactly as LEFT and the rest do. USDOLLAR is DOLLAR. JIS is not a function
# there at all: measured, ``=JIS("abc")`` is #NAME?.
function("LEFTB", V, V, minimum=1)(LEFT)
function("RIGHTB", V, V, minimum=1)(RIGHT)
function("MIDB", V, V, V)(MID)
function("LENB", V)(LEN)
function("FINDB", V, V, V, minimum=2)(FIND)
function("SEARCHB", V, V, V, minimum=2)(SEARCH)
function("REPLACEB", V, V, V, V)(REPLACE)
function("USDOLLAR", V, V, minimum=1)(DOLLAR)


@function("ENCODEURL", V)
def ENCODEURL(context: Context, text: Scalar) -> Value:
    """Text made safe for a URL: every byte of its UTF-8 but the letters,
    digits and ``-_.`` percent-encoded; measured, ``~`` is ``%7E``, which
    Python's quote always leaves alone."""
    return _checked(quote(context.text(text), safe="-_.").replace("~", "%7E"))


# ----------------------------------------------------------------------
# Regular expressions
# ----------------------------------------------------------------------


#: Syntax PCRE2 reads and Python's re does not: a pattern Python rejects
#: that uses any of it is refused rather than guessed at.
_PCRE_ONLY = re.compile(r"\\[pPKGhHvVRXNQEg]|\(\*|\(\?(?:R|[0-9+-]|\||&|P>|>)|[*+?}]\+")


def _pattern(context: Context, pattern: Scalar, case: Scalar | None) -> re.Pattern[str]:
    """A regular expression as Python reads one. Excel's are PCRE2's; the
    difference that comes up, a named group written ``(?<name>...)``, is
    translated. A pattern neither reads is ``#VALUE!``, as ``[`` is in
    Excel; one only PCRE2 reads is refused."""
    source = context.text(pattern)
    flags = re.IGNORECASE if case is not None and not isinstance(case, Empty) and context.integer(case) == 1 else 0
    source = re.sub(r"\(\?<(?=[A-Za-z_])", "(?P<", source)
    try:
        return re.compile(source, flags)
    except re.error:
        if _PCRE_ONLY.search(source):
            raise UnsupportedFormulaError("a regular expression Python reads differently") from None
        raise ExcelError(VALUE) from None


@function("REGEXTEST", V, V, V, minimum=2)
def REGEXTEST(context: Context, text: Scalar, pattern: Scalar, case: Scalar | None = None) -> Value:
    return _pattern(context, pattern, case).search(context.text(text)) is not None


@function("REGEXEXTRACT", V, V, V, V, minimum=2)
def REGEXEXTRACT(
    context: Context, text: Scalar, pattern: Scalar, mode: Scalar | None = None, case: Scalar | None = None
) -> Value:
    """The first match, every match as a column, or the first match's
    groups as a row, for return modes 0, 1 and 2."""
    source = context.text(text)
    expression = _pattern(context, pattern, case)
    which = 0 if mode is None or isinstance(mode, Empty) else context.integer(mode)
    if which == 1:
        found: list[list[Scalar]] = [[match.group(0)] for match in expression.finditer(source)]
        return Array(found) if found else NA
    match = expression.search(source)
    if match is None:
        return NA
    if which == 0:
        return match.group(0)
    if which == 2:
        groups: list[Scalar] = [group if group is not None else "" for group in match.groups()]
        return Array([groups]) if groups else match.group(0)
    return VALUE


@function("REGEXREPLACE", V, V, V, V, V, minimum=3)
def REGEXREPLACE(
    context: Context,
    text: Scalar,
    pattern: Scalar,
    replacement: Scalar,
    occurrence: Scalar | None = None,
    case: Scalar | None = None,
) -> Value:
    """Every match replaced, or only the n-th, from the end when n is
    negative; ``$1`` in the replacement is the first group."""
    source = context.text(text)
    expression = _pattern(context, pattern, case)
    template = re.sub(r"\$(\d+)|\$\{(\w+)\}", lambda found: "\\g<" + (found.group(1) or found.group(2)) + ">",
                      context.text(replacement).replace("\\", "\\\\"))  # fmt: skip
    which = 0 if occurrence is None or isinstance(occurrence, Empty) else context.integer(occurrence)
    if which == 0:
        return _checked(expression.sub(template, source))
    matches = list(expression.finditer(source))
    if abs(which) > len(matches):
        return source
    chosen = matches[which - 1] if which > 0 else matches[which]
    return _checked(source[: chosen.start()] + chosen.expand(template) + source[chosen.end() :])


__all__: list[str] = []
