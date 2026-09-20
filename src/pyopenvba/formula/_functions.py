"""The worksheet functions.

A function here either behaves as Excel's does or is absent.  An absent
name is answered from the type library rather than guessed: a real
Excel function this does not implement raises
:class:`~pyopenvba.exceptions.VBAUnsupportedError` naming itself, and a
name Excel has never had is ``#NAME?``, which is what Excel shows.

Two rules run through the aggregations and are easy to get wrong:

* Text and logicals inside a *range* are ignored, while the same values
  written as arguments are converted.  ``SUM(A1,"2")`` is three when A1
  is one; ``SUM(A1:A2)`` is one when A2 holds the text "2".
* ROUND rounds half away from zero, and MOD takes the sign of its
  divisor.  Neither matches the VBA function of the same name.
"""

from __future__ import annotations

import datetime as _dt
import math
import random
import re
import statistics
from collections.abc import Callable, Iterator, Sequence
from typing import Any, Final

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._engine import Context, evaluate
from pyopenvba.formula._values import (
    BLANK,
    DIV0,
    ERROR_NUMBERS,
    NA,
    NAME,
    NUM,
    REF,
    VALUE,
    ExcelError,
    Matrix,
    as_bool,
    as_number,
    as_text,
    compare,
    single,
    text_as_number,
)

#: Day zero of the serial numbers a date is stored as.
EPOCH: Final = _dt.datetime(1899, 12, 30)

Implementation = Callable[[Context, list[Any]], object]

FUNCTIONS: dict[str, tuple[Implementation, bool]] = {}


def function(*names: str, lazy: bool = False) -> Callable[[Implementation], Implementation]:
    """Register a worksheet function.

    A lazy one is handed the unevaluated arguments, which is what IF
    needs so that the branch not taken never runs.
    """

    def register(implementation: Implementation) -> Implementation:
        for name in names:
            FUNCTIONS[name.upper()] = (implementation, lazy)
        return implementation

    return register


def call(name: str, args: list[P.Node], context: Context) -> object:
    """Run a worksheet function, or say honestly why not."""
    upper = name.upper()
    found = FUNCTIONS.get(upper)
    if found is None:
        raise _missing(upper)
    implementation, lazy = found
    if lazy:
        return implementation(context, list(args))
    return implementation(context, [evaluate(node, context) for node in args])


def _missing(name: str) -> Exception:
    from pyopenvba.formula._inventory import excel_has_function

    if excel_has_function(name):
        return VBAUnsupportedError(
            f"{name} is a real Excel function that pyOpenVBA does not implement"
        )
    return NAME


# --- reading the arguments ---------------------------------------------------------------


def _numbers(args: Sequence[object]) -> Iterator[float]:
    """Every number an aggregation should see.

    A block contributes only its numbers; a scalar contributes whatever
    it converts to, which is why SUM(1,"2") is three and SUM of a range
    holding "2" is not.
    """
    for value in args:
        if isinstance(value, Matrix):
            for item in value.flat():
                if isinstance(item, ExcelError):
                    raise item
                if isinstance(item, bool) or item is BLANK or isinstance(item, str):
                    continue
                if isinstance(item, (int, float)):
                    yield float(item)
        elif isinstance(value, ExcelError):
            raise value
        elif value is BLANK:
            yield 0.0
        else:
            yield as_number(value)


def _values(args: Sequence[object]) -> Iterator[object]:
    """Every value, blocks flattened, blanks included."""
    for value in args:
        if isinstance(value, Matrix):
            yield from value.flat()
        else:
            yield value


def _one(args: Sequence[object], at: int, default: object = None) -> object:
    if at >= len(args):
        return default
    value = single(args[at])
    if isinstance(value, ExcelError):
        raise value
    return default if value is BLANK and default is not None else value


def _int(args: Sequence[object], at: int, default: int | None = None) -> int:
    if at >= len(args) or (isinstance(args[at], type(BLANK))):
        if default is None:
            raise VALUE
        return default
    value = single(args[at])
    if value is BLANK and default is not None:
        return default
    number = as_number(value)
    return int(number) if number >= 0 else -int(-number)


def _matrix(value: object) -> Matrix:
    return value if isinstance(value, Matrix) else Matrix([[value]])


# --- maths ----------------------------------------------------------------------------------


@function("SUM")
def fn_sum(context: Context, args: list[Any]) -> object:
    return math.fsum(_numbers(args))


@function("PRODUCT")
def fn_product(context: Context, args: list[Any]) -> object:
    total = 1.0
    seen = False
    for number in _numbers(args):
        total *= number
        seen = True
    return total if seen else 0.0


@function("ABS")
def fn_abs(context: Context, args: list[Any]) -> object:
    return abs(as_number(_one(args, 0)))


@function("SIGN")
def fn_sign(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    return float((number > 0) - (number < 0))


@function("INT")
def fn_int_of(context: Context, args: list[Any]) -> object:
    return float(math.floor(as_number(_one(args, 0))))


@function("TRUNC")
def fn_trunc(context: Context, args: list[Any]) -> object:
    places = _int(args, 1, 0)
    number = as_number(_one(args, 0))
    scale = 10.0**places
    return float(math.trunc(number * scale) / scale)


def _round_half_up(number: float, places: int) -> float:
    """Excel rounds a half away from zero, where VBA rounds to even."""
    scale = 10.0**places
    scaled = number * scale
    rounded = math.floor(abs(scaled) + 0.5)
    # The floating point representation of a value one digit past the
    # cut can sit just below the half; Excel rounds the decimal it
    # shows, so compare at fifteen digits as it does.
    if abs(abs(scaled) - (math.floor(abs(scaled)) + 0.5)) < 1e-9:
        rounded = math.floor(abs(scaled)) + 1
    return math.copysign(rounded, scaled) / scale


@function("ROUND")
def fn_round(context: Context, args: list[Any]) -> object:
    return _round_half_up(as_number(_one(args, 0)), _int(args, 1, 0))


@function("ROUNDUP")
def fn_roundup(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    scale = 10.0 ** _int(args, 1, 0)
    return math.copysign(math.ceil(abs(number) * scale - 1e-9), number) / scale


@function("ROUNDDOWN")
def fn_rounddown(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    scale = 10.0 ** _int(args, 1, 0)
    return math.copysign(math.floor(abs(number) * scale + 1e-9), number) / scale


@function("MROUND")
def fn_mround(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    step = as_number(_one(args, 1))
    if step == 0:
        return 0.0
    if (number < 0) != (step < 0):
        raise NUM
    return _round_half_up(number / step, 0) * step


@function("CEILING", "CEILING.MATH")
def fn_ceiling(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    step = as_number(_one(args, 1)) if len(args) > 1 else 1.0
    if step == 0:
        return 0.0
    if (number < 0) != (step < 0) and len(args) > 1:
        raise NUM
    return math.ceil(number / step) * step


@function("FLOOR", "FLOOR.MATH")
def fn_floor(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    step = as_number(_one(args, 1)) if len(args) > 1 else 1.0
    if step == 0:
        raise DIV0
    if (number < 0) != (step < 0) and len(args) > 1:
        raise NUM
    return math.floor(number / step) * step


@function("MOD")
def fn_mod(context: Context, args: list[Any]) -> object:
    """MOD takes the sign of its divisor, where VBA's Mod takes the dividend's."""
    number = as_number(_one(args, 0))
    divisor = as_number(_one(args, 1))
    if divisor == 0:
        raise DIV0
    return number - divisor * math.floor(number / divisor)


@function("POWER")
def fn_power(context: Context, args: list[Any]) -> object:
    try:
        result = as_number(_one(args, 0)) ** as_number(_one(args, 1))
    except (OverflowError, ValueError, ZeroDivisionError):
        raise NUM from None
    if isinstance(result, complex):
        raise NUM
    return float(result)


@function("SQRT")
def fn_sqrt(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    if number < 0:
        raise NUM
    return math.sqrt(number)


@function("EXP")
def fn_exp(context: Context, args: list[Any]) -> object:
    try:
        return math.exp(as_number(_one(args, 0)))
    except OverflowError:
        raise NUM from None


@function("LN")
def fn_ln(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    if number <= 0:
        raise NUM
    return math.log(number)


@function("LOG")
def fn_log(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    base = as_number(_one(args, 1)) if len(args) > 1 else 10.0
    if number <= 0 or base <= 0 or base == 1:
        raise NUM
    return math.log(number, base)


@function("LOG10")
def fn_log10(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    if number <= 0:
        raise NUM
    return math.log10(number)


@function("PI")
def fn_pi(context: Context, args: list[Any]) -> object:
    return math.pi


@function("RAND")
def fn_rand(context: Context, args: list[Any]) -> object:
    return random.random()


@function("RANDBETWEEN")
def fn_randbetween(context: Context, args: list[Any]) -> object:
    low = _int(args, 0)
    high = _int(args, 1)
    if low > high:
        raise NUM
    return float(random.randint(low, high))


@function("SUMPRODUCT")
def fn_sumproduct(context: Context, args: list[Any]) -> object:
    if not args:
        raise VALUE
    blocks = [_matrix(one) for one in args]
    height = max(block.height for block in blocks)
    width = max(block.width for block in blocks)
    total = 0.0
    for row in range(height):
        for column in range(width):
            product = 1.0
            for block in blocks:
                item = block.at(row, column)
                if isinstance(item, ExcelError):
                    raise item
                if isinstance(item, bool) or isinstance(item, str) or item is BLANK:
                    product = 0.0
                    continue
                product *= float(item)  # type: ignore[arg-type]
            total += product
    return total


# --- statistics --------------------------------------------------------------------------------


@function("AVERAGE")
def fn_average(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    if not numbers:
        raise DIV0
    return math.fsum(numbers) / len(numbers)


@function("MEDIAN")
def fn_median(context: Context, args: list[Any]) -> object:
    numbers = sorted(_numbers(args))
    if not numbers:
        raise NUM
    return float(statistics.median(numbers))


@function("MAX")
def fn_max(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    return max(numbers) if numbers else 0.0


@function("MIN")
def fn_min(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    return min(numbers) if numbers else 0.0


@function("LARGE")
def fn_large(context: Context, args: list[Any]) -> object:
    numbers = sorted(_numbers([args[0]]), reverse=True)
    which = _int(args, 1)
    if which < 1 or which > len(numbers):
        raise NUM
    return numbers[which - 1]


@function("SMALL")
def fn_small(context: Context, args: list[Any]) -> object:
    numbers = sorted(_numbers([args[0]]))
    which = _int(args, 1)
    if which < 1 or which > len(numbers):
        raise NUM
    return numbers[which - 1]


@function("COUNT")
def fn_count(context: Context, args: list[Any]) -> object:
    """COUNT counts numbers and walks past an error, where SUM stops at one."""
    total = 0
    for value in _values(args):
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            total += 1
    return float(total)


@function("COUNTA")
def fn_counta(context: Context, args: list[Any]) -> object:
    return float(sum(1 for value in _values(args) if value is not BLANK))


@function("COUNTBLANK")
def fn_countblank(context: Context, args: list[Any]) -> object:
    return float(sum(1 for value in _values(args) if value is BLANK or value == ""))


@function("STDEV", "STDEV.S")
def fn_stdev(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    if len(numbers) < 2:
        raise DIV0
    return float(statistics.stdev(numbers))


@function("STDEVP", "STDEV.P")
def fn_stdevp(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    if not numbers:
        raise DIV0
    return float(statistics.pstdev(numbers))


@function("VAR", "VAR.S")
def fn_var(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    if len(numbers) < 2:
        raise DIV0
    return float(statistics.variance(numbers))


@function("VARP", "VAR.P")
def fn_varp(context: Context, args: list[Any]) -> object:
    numbers = list(_numbers(args))
    if not numbers:
        raise DIV0
    return float(statistics.pvariance(numbers))


# --- criteria ------------------------------------------------------------------------------------

_COMPARISON: Final = re.compile(r"^(<>|>=|<=|=|<|>)(.*)$", re.DOTALL)


def _criterion(spec: object) -> Callable[[object], bool]:
    """The test a COUNTIF-style criterion stands for."""
    if isinstance(spec, Matrix):
        spec = spec.first()
    if isinstance(spec, ExcelError):
        raise spec
    if not isinstance(spec, str):
        wanted = spec
        return lambda value: _equal(value, wanted)
    match = _COMPARISON.match(spec)
    op, body = (match.group(1), match.group(2)) if match else ("=", spec)
    number = text_as_number(body)
    target: object = number if number is not None else body
    if op in ("=", "<>"):
        if isinstance(target, str) and ("*" in target or "?" in target):
            pattern = _wildcard(target)
            return lambda value: (
                pattern.fullmatch(as_text(value)) is not None if op == "=" else pattern.fullmatch(as_text(value)) is None
            )
        if body == "":
            return (lambda value: value is BLANK or value == "") if op == "=" else (
                lambda value: not (value is BLANK or value == "")
            )
        return (lambda value: _equal(value, target)) if op == "=" else (
            lambda value: not _equal(value, target)
        )

    def ordered(value: object) -> bool:
        if value is BLANK:
            return False
        try:
            return compare(op, value, target)
        except ExcelError:
            return False

    return ordered


def _equal(value: object, target: object) -> bool:
    if value is BLANK:
        return target is BLANK
    if isinstance(target, str) and isinstance(value, str):
        return value.upper() == target.upper()
    if isinstance(target, str) != isinstance(value, str):
        return False
    try:
        return bool(compare("=", value, target))
    except ExcelError:
        return False


def _wildcard(pattern: str) -> re.Pattern[str]:
    out: list[str] = []
    index = 0
    while index < len(pattern):
        char = pattern[index]
        if char == "~" and index + 1 < len(pattern) and pattern[index + 1] in "*?~":
            out.append(re.escape(pattern[index + 1]))
            index += 2
            continue
        if char == "*":
            out.append(".*")
        elif char == "?":
            out.append(".")
        else:
            out.append(re.escape(char))
        index += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


def _pairs(args: Sequence[object], start: int) -> list[tuple[Matrix, Callable[[object], bool]]]:
    out: list[tuple[Matrix, Callable[[object], bool]]] = []
    index = start
    while index + 1 < len(args):
        out.append((_matrix(args[index]), _criterion(args[index + 1])))
        index += 2
    return out


def _matching(tests: list[tuple[Matrix, Callable[[object], bool]]]) -> Iterator[tuple[int, int]]:
    if not tests:
        return
    height = max(block.height for block, _ in tests)
    width = max(block.width for block, _ in tests)
    for row in range(height):
        for column in range(width):
            if all(test(block.at(row, column)) for block, test in tests):
                yield row, column


@function("COUNTIF")
def fn_countif(context: Context, args: list[Any]) -> object:
    block = _matrix(args[0])
    test = _criterion(args[1] if len(args) > 1 else BLANK)
    return float(sum(1 for value in block.flat() if test(value)))


@function("COUNTIFS")
def fn_countifs(context: Context, args: list[Any]) -> object:
    return float(sum(1 for _ in _matching(_pairs(args, 0))))


@function("SUMIF")
def fn_sumif(context: Context, args: list[Any]) -> object:
    block = _matrix(args[0])
    test = _criterion(args[1] if len(args) > 1 else BLANK)
    target = _matrix(args[2]) if len(args) > 2 else block
    total = 0.0
    for row in range(block.height):
        for column in range(block.width):
            if not test(block.at(row, column)):
                continue
            value = target.at(row, column) if target is not block else block.at(row, column)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                total += float(value)
    return total


@function("SUMIFS")
def fn_sumifs(context: Context, args: list[Any]) -> object:
    target = _matrix(args[0])
    total = 0.0
    for row, column in _matching(_pairs(args, 1)):
        value = target.at(row, column)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            total += float(value)
    return total


@function("AVERAGEIF")
def fn_averageif(context: Context, args: list[Any]) -> object:
    block = _matrix(args[0])
    test = _criterion(args[1] if len(args) > 1 else BLANK)
    target = _matrix(args[2]) if len(args) > 2 else block
    kept: list[float] = []
    for row in range(block.height):
        for column in range(block.width):
            if not test(block.at(row, column)):
                continue
            value = target.at(row, column) if target is not block else block.at(row, column)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                kept.append(float(value))
    if not kept:
        raise DIV0
    return math.fsum(kept) / len(kept)


# --- logic -------------------------------------------------------------------------------------------


@function("IF", lazy=True)
def fn_if(context: Context, nodes: list[Any]) -> object:
    """IF runs only the branch it takes, which is why it is lazy."""
    if not nodes:
        raise VALUE
    condition = as_bool(single(evaluate(nodes[0], context)))
    if condition:
        return single(evaluate(nodes[1], context)) if len(nodes) > 1 else True
    if len(nodes) > 2:
        return single(evaluate(nodes[2], context))
    return False


@function("IFS", lazy=True)
def fn_ifs(context: Context, nodes: list[Any]) -> object:
    for index in range(0, len(nodes) - 1, 2):
        if as_bool(single(evaluate(nodes[index], context))):
            return single(evaluate(nodes[index + 1], context))
    raise NA


@function("IFERROR", lazy=True)
def fn_iferror(context: Context, nodes: list[Any]) -> object:
    try:
        value = evaluate(nodes[0], context)
    except ExcelError:
        return single(evaluate(nodes[1], context)) if len(nodes) > 1 else BLANK
    value = single(value)
    if isinstance(value, ExcelError):
        return single(evaluate(nodes[1], context)) if len(nodes) > 1 else BLANK
    return value


@function("IFNA", lazy=True)
def fn_ifna(context: Context, nodes: list[Any]) -> object:
    try:
        value = single(evaluate(nodes[0], context))
    except ExcelError as failure:
        if failure != NA:
            raise
        return single(evaluate(nodes[1], context)) if len(nodes) > 1 else BLANK
    if value == NA:
        return single(evaluate(nodes[1], context)) if len(nodes) > 1 else BLANK
    return value


@function("AND")
def fn_and(context: Context, args: list[Any]) -> object:
    seen = False
    for value in _values(args):
        if value is BLANK:
            continue
        if isinstance(value, ExcelError):
            raise value
        if isinstance(value, str):
            continue
        seen = True
        if not as_bool(value):
            return False
    if not seen:
        raise VALUE
    return True


@function("OR")
def fn_or(context: Context, args: list[Any]) -> object:
    seen = False
    answer = False
    for value in _values(args):
        if value is BLANK:
            continue
        if isinstance(value, ExcelError):
            raise value
        if isinstance(value, str):
            continue
        seen = True
        answer = answer or as_bool(value)
    if not seen:
        raise VALUE
    return answer


@function("XOR")
def fn_xor(context: Context, args: list[Any]) -> object:
    count = sum(1 for value in _values(args) if value is not BLANK and as_bool(value))
    return count % 2 == 1


@function("NOT")
def fn_not(context: Context, args: list[Any]) -> object:
    return not as_bool(_one(args, 0))


@function("TRUE")
def fn_true(context: Context, args: list[Any]) -> object:
    return True


@function("FALSE")
def fn_false(context: Context, args: list[Any]) -> object:
    return False


@function("SWITCH", lazy=True)
def fn_switch(context: Context, nodes: list[Any]) -> object:
    subject = single(evaluate(nodes[0], context))
    index = 1
    while index + 1 < len(nodes):
        if _equal(subject, single(evaluate(nodes[index], context))):
            return single(evaluate(nodes[index + 1], context))
        index += 2
    if index < len(nodes):
        return single(evaluate(nodes[index], context))
    raise NA


@function("CHOOSE", lazy=True)
def fn_choose(context: Context, nodes: list[Any]) -> object:
    which = int(as_number(single(evaluate(nodes[0], context))))
    if which < 1 or which >= len(nodes):
        raise VALUE
    return single(evaluate(nodes[which], context))


# --- lookup ----------------------------------------------------------------------------------------


@function("VLOOKUP")
def fn_vlookup(context: Context, args: list[Any]) -> object:
    wanted = _one(args, 0)
    table = _matrix(args[1])
    column = _int(args, 2)
    approximate = as_bool(_one(args, 3)) if len(args) > 3 and args[3] is not BLANK else True
    if column < 1 or column > table.width:
        raise VALUE
    row = _find_row(table, wanted, 0, approximate)
    if row is None:
        raise NA
    return table.at(row, column - 1)


@function("HLOOKUP")
def fn_hlookup(context: Context, args: list[Any]) -> object:
    wanted = _one(args, 0)
    table = _matrix(args[1])
    row = _int(args, 2)
    approximate = as_bool(_one(args, 3)) if len(args) > 3 and args[3] is not BLANK else True
    if row < 1 or row > table.height:
        raise VALUE
    column = _find_column(table, wanted, 0, approximate)
    if column is None:
        raise NA
    return table.at(row - 1, column)


def _find_row(table: Matrix, wanted: object, column: int, approximate: bool) -> int | None:
    if not approximate:
        test = _criterion(wanted) if isinstance(wanted, str) else None
        for row in range(table.height):
            value = table.at(row, column)
            if test(value) if test is not None else _equal(value, wanted):
                return row
        return None
    best: int | None = None
    for row in range(table.height):
        value = table.at(row, column)
        if value is BLANK:
            continue
        try:
            if compare("<=", value, wanted):
                best = row
            else:
                break
        except ExcelError:
            continue
    return best


def _find_column(table: Matrix, wanted: object, row: int, approximate: bool) -> int | None:
    if not approximate:
        for column in range(table.width):
            if _equal(table.at(row, column), wanted):
                return column
        return None
    best: int | None = None
    for column in range(table.width):
        value = table.at(row, column)
        if value is BLANK:
            continue
        try:
            if compare("<=", value, wanted):
                best = column
            else:
                break
        except ExcelError:
            continue
    return best


@function("MATCH")
def fn_match(context: Context, args: list[Any]) -> object:
    wanted = _one(args, 0)
    block = _matrix(args[1])
    kind = _int(args, 2, 1)
    items = list(block.flat())
    if kind == 0:
        test = _criterion(wanted) if isinstance(wanted, str) else None
        for index, value in enumerate(items):
            if test(value) if test is not None else _equal(value, wanted):
                return float(index + 1)
        raise NA
    best: int | None = None
    for index, value in enumerate(items):
        if value is BLANK:
            continue
        try:
            fits = compare("<=", value, wanted) if kind > 0 else compare(">=", value, wanted)
        except ExcelError:
            continue
        if fits:
            best = index
        else:
            break
    if best is None:
        raise NA
    return float(best + 1)


@function("INDEX")
def fn_index(context: Context, args: list[Any]) -> object:
    block = _matrix(args[0])
    row = _int(args, 1, 0)
    column = _int(args, 2, 0)
    if block.single:
        return block.first()
    if row == 0 and column == 0:
        return block
    if row == 0:
        if column < 1 or column > block.width:
            raise REF
        return Matrix([[line[column - 1]] for line in block.rows])
    if column == 0:
        if block.width == 1:
            if row < 1 or row > block.height:
                raise REF
            return block.rows[row - 1][0]
        if row < 1 or row > block.height:
            raise REF
        return Matrix([list(block.rows[row - 1])])
    if row < 1 or row > block.height or column < 1 or column > block.width:
        raise REF
    return block.rows[row - 1][column - 1]


@function("XLOOKUP")
def fn_xlookup(context: Context, args: list[Any]) -> object:
    wanted = _one(args, 0)
    where = _matrix(args[1])
    give = _matrix(args[2])
    items = list(where.flat())
    test = _criterion(wanted) if isinstance(wanted, str) else None
    for index, value in enumerate(items):
        if test(value) if test is not None else _equal(value, wanted):
            return give.at(index, 0) if give.width == 1 else give.at(0, index)
    if len(args) > 3 and args[3] is not BLANK:
        return single(args[3])
    raise NA


@function("ROW", lazy=True)
def fn_row(context: Context, nodes: list[Any]) -> object:
    if not nodes:
        return float(context.row)
    area = context.area_of(nodes[0])
    if area is None:
        raise REF
    if area.rows == 1:
        return float(area.top)
    return Matrix([[float(row)] for row in range(area.top, min(area.bottom, area.top + 9999) + 1)])


@function("COLUMN", lazy=True)
def fn_column(context: Context, nodes: list[Any]) -> object:
    if not nodes:
        return float(context.column)
    area = context.area_of(nodes[0])
    if area is None:
        raise REF
    if area.columns == 1:
        return float(area.left)
    return Matrix([[float(column) for column in range(area.left, min(area.right, area.left + 9999) + 1)]])


@function("ROWS", lazy=True)
def fn_rows(context: Context, nodes: list[Any]) -> object:
    area = context.area_of(nodes[0]) if nodes else None
    if area is not None:
        return float(area.rows)
    return float(_matrix(evaluate(nodes[0], context)).height)


@function("COLUMNS", lazy=True)
def fn_columns(context: Context, nodes: list[Any]) -> object:
    area = context.area_of(nodes[0]) if nodes else None
    if area is not None:
        return float(area.columns)
    return float(_matrix(evaluate(nodes[0], context)).width)


@function("OFFSET", lazy=True)
def fn_offset(context: Context, nodes: list[Any]) -> object:
    area = context.area_of(nodes[0]) if nodes else None
    if area is None:
        raise REF
    down = int(as_number(single(evaluate(nodes[1], context)))) if len(nodes) > 1 else 0
    across = int(as_number(single(evaluate(nodes[2], context)))) if len(nodes) > 2 else 0
    height = area.rows
    width = area.columns
    if len(nodes) > 3 and not isinstance(nodes[3], P.Literal):
        height = int(as_number(single(evaluate(nodes[3], context))))
    elif len(nodes) > 3 and isinstance(nodes[3], P.Literal) and nodes[3].value is not None:
        height = int(as_number(nodes[3].value))
    if len(nodes) > 4 and not isinstance(nodes[4], P.Literal):
        width = int(as_number(single(evaluate(nodes[4], context))))
    elif len(nodes) > 4 and isinstance(nodes[4], P.Literal) and nodes[4].value is not None:
        width = int(as_number(nodes[4].value))
    top = area.top + down
    left = area.left + across
    if top < 1 or left < 1 or height < 1 or width < 1:
        raise REF
    if top + height - 1 > MAX_ROWS or left + width - 1 > MAX_COLUMNS:
        raise REF
    moved = Area(top, left, top + height - 1, left + width - 1, area.sheet)
    return context.grid.block(context.sheet, moved)


@function("INDIRECT", lazy=True)
def fn_indirect(context: Context, nodes: list[Any]) -> object:
    text = as_text(single(evaluate(nodes[0], context)))
    sheet, body = P.split_sheet(text)
    reference = P.Reference(text=body.replace("$", ""), sheet=sheet)
    try:
        area = context.resolve(reference)
    except ExcelError:
        raise REF from None
    return context.grid.block(context.sheet, area)


@function("TRANSPOSE")
def fn_transpose(context: Context, args: list[Any]) -> object:
    block = _matrix(args[0])
    return Matrix([[block.rows[row][column] for row in range(block.height)] for column in range(block.width)])


# --- text ---------------------------------------------------------------------------------------------


@function("LEN")
def fn_len(context: Context, args: list[Any]) -> object:
    return float(len(as_text(_one(args, 0, BLANK))))


@function("LEFT")
def fn_left(context: Context, args: list[Any]) -> object:
    count = _int(args, 1, 1)
    if count < 0:
        raise VALUE
    return as_text(_one(args, 0, BLANK))[:count]


@function("RIGHT")
def fn_right(context: Context, args: list[Any]) -> object:
    count = _int(args, 1, 1)
    if count < 0:
        raise VALUE
    body = as_text(_one(args, 0, BLANK))
    return body[max(0, len(body) - count) :] if count else ""


@function("MID")
def fn_mid(context: Context, args: list[Any]) -> object:
    body = as_text(_one(args, 0, BLANK))
    start = _int(args, 1)
    count = _int(args, 2)
    if start < 1 or count < 0:
        raise VALUE
    return body[start - 1 : start - 1 + count]


@function("TRIM")
def fn_trim(context: Context, args: list[Any]) -> object:
    return " ".join(as_text(_one(args, 0, BLANK)).split())


@function("UPPER")
def fn_upper(context: Context, args: list[Any]) -> object:
    return as_text(_one(args, 0, BLANK)).upper()


@function("LOWER")
def fn_lower(context: Context, args: list[Any]) -> object:
    return as_text(_one(args, 0, BLANK)).lower()


@function("PROPER")
def fn_proper(context: Context, args: list[Any]) -> object:
    return re.sub(
        r"[A-Za-z]+",
        lambda match: match.group(0)[0].upper() + match.group(0)[1:].lower(),
        as_text(_one(args, 0, BLANK)),
    )


@function("CONCATENATE", "CONCAT")
def fn_concatenate(context: Context, args: list[Any]) -> object:
    return "".join(as_text(value) for value in _values(args) if value is not BLANK)


@function("TEXTJOIN")
def fn_textjoin(context: Context, args: list[Any]) -> object:
    separator = as_text(_one(args, 0, BLANK))
    skip = as_bool(_one(args, 1)) if len(args) > 1 else True
    pieces = [
        as_text(value)
        for value in _values(args[2:])
        if not (skip and (value is BLANK or value == ""))
    ]
    return separator.join(pieces)


@function("SUBSTITUTE")
def fn_substitute(context: Context, args: list[Any]) -> object:
    body = as_text(_one(args, 0, BLANK))
    old = as_text(_one(args, 1, BLANK))
    new = as_text(_one(args, 2, BLANK))
    if not old:
        return body
    if len(args) > 3 and args[3] is not BLANK:
        which = _int(args, 3)
        if which < 1:
            raise VALUE
        pieces = body.split(old)
        if len(pieces) <= which:
            return body
        return old.join(pieces[:which]) + new + old.join(pieces[which:])
    return body.replace(old, new)


@function("REPLACE")
def fn_replace(context: Context, args: list[Any]) -> object:
    body = as_text(_one(args, 0, BLANK))
    start = _int(args, 1)
    count = _int(args, 2)
    new = as_text(_one(args, 3, BLANK))
    if start < 1 or count < 0:
        raise VALUE
    return body[: start - 1] + new + body[start - 1 + count :]


@function("FIND")
def fn_find(context: Context, args: list[Any]) -> object:
    needle = as_text(_one(args, 0, BLANK))
    body = as_text(_one(args, 1, BLANK))
    start = _int(args, 2, 1)
    if start < 1 or start > len(body) + 1:
        raise VALUE
    at = body.find(needle, start - 1)
    if at < 0:
        raise VALUE
    return float(at + 1)


@function("SEARCH")
def fn_search(context: Context, args: list[Any]) -> object:
    needle = as_text(_one(args, 0, BLANK))
    body = as_text(_one(args, 1, BLANK))
    start = _int(args, 2, 1)
    if start < 1 or start > len(body) + 1:
        raise VALUE
    pattern = _wildcard(needle)
    match = pattern.search(body, start - 1)
    if match is None:
        raise VALUE
    return float(match.start() + 1)


@function("REPT")
def fn_rept(context: Context, args: list[Any]) -> object:
    times = _int(args, 1)
    if times < 0:
        raise VALUE
    return as_text(_one(args, 0, BLANK)) * times


@function("EXACT")
def fn_exact(context: Context, args: list[Any]) -> object:
    return as_text(_one(args, 0, BLANK)) == as_text(_one(args, 1, BLANK))


@function("CHAR")
def fn_char(context: Context, args: list[Any]) -> object:
    code = _int(args, 0)
    if not 1 <= code <= 255:
        raise VALUE
    return chr(code)


@function("CODE")
def fn_code(context: Context, args: list[Any]) -> object:
    body = as_text(_one(args, 0, BLANK))
    if not body:
        raise VALUE
    return float(ord(body[0]))


@function("VALUE")
def fn_value(context: Context, args: list[Any]) -> object:
    value = _one(args, 0, BLANK)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    number = text_as_number(as_text(value))
    if number is None:
        raise VALUE
    return number


@function("TEXT")
def fn_text_function(context: Context, args: list[Any]) -> object:
    from pyopenvba.access._format import format_value

    value = _one(args, 0, BLANK)
    pattern = as_text(_one(args, 1, BLANK))
    if value is BLANK:
        value = 0.0
    if _is_date_pattern(pattern):
        return format_value(_as_datetime(as_number(value)), pattern)
    return format_value(as_number(value) if not isinstance(value, str) else value, pattern)


def _is_date_pattern(pattern: str) -> bool:
    body = re.sub(r'"[^"]*"', "", pattern)
    return any(char in body for char in "ymdhs")


@function("T")
def fn_t(context: Context, args: list[Any]) -> object:
    value = _one(args, 0, BLANK)
    return value if isinstance(value, str) else ""


@function("N")
def fn_n(context: Context, args: list[Any]) -> object:
    value = _one(args, 0, BLANK)
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    return 0.0


# --- dates ----------------------------------------------------------------------------------------------


def _as_datetime(serial: float) -> _dt.datetime:
    days = math.floor(serial)
    seconds = round((serial - days) * 86400)
    return EPOCH + _dt.timedelta(days=days, seconds=seconds)


def _as_serial(when: _dt.datetime) -> float:
    delta = when - EPOCH
    return delta.days + delta.seconds / 86400.0


@function("TODAY")
def fn_today(context: Context, args: list[Any]) -> object:
    now = context.grid.now()
    return float((_dt.datetime(now.year, now.month, now.day) - EPOCH).days)


@function("NOW")
def fn_now(context: Context, args: list[Any]) -> object:
    return _as_serial(context.grid.now())


@function("DATE")
def fn_date(context: Context, args: list[Any]) -> object:
    year = _int(args, 0)
    month = _int(args, 1)
    day = _int(args, 2)
    if year < 1900:
        year += 1900
    year += (month - 1) // 12
    month = (month - 1) % 12 + 1
    try:
        base = _dt.datetime(year, month, 1)
    except ValueError:
        raise NUM from None
    return float((base - EPOCH).days + day - 1)


@function("TIME")
def fn_time(context: Context, args: list[Any]) -> object:
    seconds = _int(args, 0) * 3600 + _int(args, 1) * 60 + _int(args, 2)
    return (seconds % 86400) / 86400.0


@function("YEAR")
def fn_year(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).year)


@function("MONTH")
def fn_month(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).month)


@function("DAY")
def fn_day(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).day)


@function("HOUR")
def fn_hour(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).hour)


@function("MINUTE")
def fn_minute(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).minute)


@function("SECOND")
def fn_second(context: Context, args: list[Any]) -> object:
    return float(_as_datetime(as_number(_one(args, 0, BLANK))).second)


@function("WEEKDAY")
def fn_weekday(context: Context, args: list[Any]) -> object:
    when = _as_datetime(as_number(_one(args, 0, BLANK)))
    kind = _int(args, 1, 1)
    sunday_based = (when.weekday() + 1) % 7 + 1
    if kind == 1:
        return float(sunday_based)
    if kind == 2:
        return float((sunday_based - 2) % 7 + 1)
    if kind == 3:
        return float((sunday_based - 2) % 7)
    raise NUM


@function("EDATE")
def fn_edate(context: Context, args: list[Any]) -> object:
    when = _as_datetime(as_number(_one(args, 0, BLANK)))
    months = _int(args, 1)
    total = when.year * 12 + when.month - 1 + months
    year, month = divmod(total, 12)
    day = min(when.day, _days_in_month(year, month + 1))
    return float((_dt.datetime(year, month + 1, day) - EPOCH).days)


@function("EOMONTH")
def fn_eomonth(context: Context, args: list[Any]) -> object:
    when = _as_datetime(as_number(_one(args, 0, BLANK)))
    months = _int(args, 1)
    total = when.year * 12 + when.month - 1 + months
    year, month = divmod(total, 12)
    last = _days_in_month(year, month + 1)
    return float((_dt.datetime(year, month + 1, last) - EPOCH).days)


def _days_in_month(year: int, month: int) -> int:
    if month == 12:
        return 31
    return (_dt.date(year, month + 1, 1) - _dt.date(year, month, 1)).days


@function("DAYS")
def fn_days(context: Context, args: list[Any]) -> object:
    return float(int(as_number(_one(args, 0, BLANK))) - int(as_number(_one(args, 1, BLANK))))


@function("DATEVALUE")
def fn_datevalue(context: Context, args: list[Any]) -> object:
    from pyopenvba.interpreter._values import parse_date_text

    parsed = parse_date_text(as_text(_one(args, 0, BLANK)))
    if parsed is None:
        raise VALUE
    return float(math.floor(parsed.serial))


# --- information -------------------------------------------------------------------------------------------


@function("ISBLANK")
def fn_isblank(context: Context, args: list[Any]) -> object:
    return single(args[0]) is BLANK if args else False


@function("ISNUMBER")
def fn_isnumber(context: Context, args: list[Any]) -> object:
    value = single(args[0]) if args else BLANK
    return isinstance(value, (int, float)) and not isinstance(value, bool)


@function("ISTEXT")
def fn_istext(context: Context, args: list[Any]) -> object:
    return isinstance(single(args[0]) if args else BLANK, str)


@function("ISNONTEXT")
def fn_isnontext(context: Context, args: list[Any]) -> object:
    return not isinstance(single(args[0]) if args else BLANK, str)


@function("ISLOGICAL")
def fn_islogical(context: Context, args: list[Any]) -> object:
    return isinstance(single(args[0]) if args else BLANK, bool)


@function("ISERROR", lazy=True)
def fn_iserror(context: Context, nodes: list[Any]) -> object:
    try:
        return isinstance(single(evaluate(nodes[0], context)), ExcelError)
    except ExcelError:
        return True


@function("ISERR", lazy=True)
def fn_iserr(context: Context, nodes: list[Any]) -> object:
    try:
        value = single(evaluate(nodes[0], context))
    except ExcelError as failure:
        return failure != NA
    return isinstance(value, ExcelError) and value != NA


@function("ISNA", lazy=True)
def fn_isna(context: Context, nodes: list[Any]) -> object:
    try:
        return single(evaluate(nodes[0], context)) == NA
    except ExcelError as failure:
        return failure == NA


@function("NA")
def fn_na(context: Context, args: list[Any]) -> object:
    raise NA


@function("ERROR.TYPE", lazy=True)
def fn_error_type(context: Context, nodes: list[Any]) -> object:
    try:
        value = single(evaluate(nodes[0], context))
    except ExcelError as failure:
        return float(ERROR_NUMBERS.get(failure.name, 8))
    if isinstance(value, ExcelError):
        return float(ERROR_NUMBERS.get(value.name, 8))
    raise NA


@function("TYPE")
def fn_type(context: Context, args: list[Any]) -> object:
    value = args[0] if args else BLANK
    if isinstance(value, Matrix) and not value.single:
        return 64.0
    value = single(value)
    if isinstance(value, bool):
        return 4.0
    if isinstance(value, str):
        return 2.0
    if isinstance(value, ExcelError):
        return 16.0
    return 1.0


@function("ADDRESS")
def fn_address(context: Context, args: list[Any]) -> object:
    row = _int(args, 0)
    column = _int(args, 1)
    kind = _int(args, 2, 1)
    if row < 1 or column < 1:
        raise VALUE
    row_mark = "$" if kind in (1, 3) else ""
    column_mark = "$" if kind in (1, 2) else ""
    return f"{column_mark}{column_letter(column)}{row_mark}{row}"


@function("FORMULATEXT", lazy=True)
def fn_formulatext(context: Context, nodes: list[Any]) -> object:
    area = context.area_of(nodes[0]) if nodes else None
    if area is None:
        raise NA
    text = context.grid.formula_at(area.sheet, area.top, area.left)
    if not text:
        raise NA
    return text


#: Functions whose answer is a date, so a cell holding one shows a date
#: rather than the serial number behind it.
#: Measured: DATE and TODAY bring a date format with them, while
#: EDATE, EOMONTH, TIME and DATEVALUE leave the cell showing the serial
#: number.
DATE_RESULTS: Final = frozenset({"DATE", "TODAY"})
TIME_RESULTS: Final[frozenset[str]] = frozenset()
MOMENT_RESULTS: Final = frozenset({"NOW"})


def result_format(node: object) -> str:
    """The number format Excel gives a cell for this formula, if any.

    A formula that answers with a date arrives formatted as one: the
    cell shows 3/4/2021 rather than 44259, and VBA reading it back gets
    a Date.  It is the function that decides, not the value, which is
    why =44259 on its own stays a number.
    """
    if isinstance(node, P.Call):
        name = node.name.upper()
        if name in DATE_RESULTS:
            return "m/d/yyyy"
        if name in TIME_RESULTS:
            return "h:mm:ss AM/PM"
        if name in MOMENT_RESULTS:
            return "m/d/yyyy h:mm"
    return ""


def known_names() -> frozenset[str]:
    """Every function this implements, for the inventory to check against."""
    return frozenset(FUNCTIONS)
