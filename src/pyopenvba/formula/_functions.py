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
from collections.abc import Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass, replace
from typing import Any, Final

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, column_letter
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._engine import Context, clip, evaluate, intersected
from pyopenvba.formula._values import (
    BLANK,
    DIV0,
    ERROR_NUMBERS,
    NA,
    NAME,
    NUM,
    REF,
    VALUE,
    Areas,
    ExcelError,
    Matrix,
    as_bool,
    as_number,
    as_text,
    compare,
    single,
    snapped,
    text_as_number,
)

#: Day zero of the serial numbers a date is stored as.
EPOCH: Final = _dt.datetime(1899, 12, 30)

Implementation = Callable[[Context, list[Any]], object]
Referring = Callable[[Context, list[P.Node]], "list[Area] | None"]

FUNCTIONS: dict[str, tuple[Implementation, bool]] = {}
#: Functions whose answer can be cells, and how each works out which: what the range operator, an
#: intersection or ROWS reads from INDEX(A:A,5) or OFFSET(A1,2,0) (tests/fixtures/reference_forms.json).
REFERENCES: dict[str, Referring] = {}
#: Functions measured reading every block of a union, (A1:A2,A4), as one argument.
_READS_AREAS: Final = frozenset({"SUM", "COUNT", "COUNTA", "AVERAGE", "MAX", "MIN", "LARGE", "INDEX"})
#: Functions measured refusing a union with #VALUE!.
_REFUSES_AREAS: Final = frozenset({"COUNTIF"})


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


def refers(*names: str) -> Callable[[Referring], Referring]:
    """Register how a function that can answer with cells works out which cells."""

    def register(implementation: Referring) -> Referring:
        for name in names:
            REFERENCES[name.upper()] = implementation
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
    count = len(args)
    lifted = _places(_LIFTED, upper, count)
    one = lifted | _places(_FIRST_ITEM, upper, count)
    whole = _places(_ARRAYS, upper, count)
    for index in _places(CELLS, upper, count):
        if isinstance(args[index], (P.NameNode, P.Call)) and context.areas_of(args[index]) is None:
            # A name or a function standing where cells are read, and coming to a value: COUNTIF(dbl,4).
            raise VALUE
    as_array = replace(context, array_argument=True) if whole else context
    values = [_argument(node, as_array if index in whole else context, one=index in one)
              for index, node in enumerate(args)]
    if any(isinstance(value, Areas) for value in values):
        values = _with_areas(upper, values)
    arrays = {index: value for index in lifted if isinstance(value := values[index], Matrix) and not value.single}
    if arrays:
        return _item_by_item(implementation, context, values, arrays)
    return implementation(context, values)


def _argument(node: P.Node, context: Context, *, one: bool) -> object:
    """What an argument comes to, an error included: a function is handed #DIV/0! as a value, and COUNT passes it."""
    try:
        return intersected(node, context) if one else evaluate(node, context)
    except ExcelError as failure:
        if failure.name == "#CIRCULAR!":
            raise
        return failure


def _item_by_item(implementation: Implementation, context: Context, values: list[object],
                  arrays: dict[int, Matrix]) -> object:
    """A function run once for each item of the arrays given where it wants one value, its answers an array.

    The arrays line up as :meth:`Matrix.at` lines them up: one row or
    one column repeats across the others, and past the end of a shorter
    array an item is #N/A.
    """
    height = max(array.height for array in arrays.values())
    width = max(array.width for array in arrays.values())
    rows: list[list[object]] = []
    for row in range(height):
        line: list[object] = []
        for column in range(width):
            given = list(values)
            for index, array in arrays.items():
                given[index] = array.at(row, column)
            try:
                answer = implementation(context, given)
            except ExcelError as failure:
                answer = failure
            line.append(single(answer) if isinstance(answer, Matrix) else answer)
        rows.append(line)
    return Matrix(rows)


@dataclass(frozen=True)
class Places:
    """Which of a function's arguments a rule covers: some places, and every ``step``-th from ``start`` on."""

    fixed: tuple[int, ...] = ()
    start: int | None = None
    step: int = 1

    def within(self, count: int) -> frozenset[int]:
        run = range(self.start, count, self.step) if self.start is not None else range(0)
        return frozenset(index for index in (*self.fixed, *run) if index < count)


_EVERY: Final = Places(start=0)

# Measured in tests/fixtures/formula/probes.txt, one probe or more per function.

#: Where a function wants one value and runs item by item through an array given there, its answer an array:
#: SUM(LEN({"a","bb"})) is 3. In a cell, cells given there are first cut to the formula's own row or column.
_LIFTED: Final[dict[str, Places]] = {
    **dict.fromkeys((
        "ABS", "ADDRESS", "CEILING", "CEILING.MATH", "CHAR", "CODE", "CONCATENATE", "DATE", "DATEVALUE", "DAY",
        "DAYS", "EDATE", "EOMONTH", "ERROR.TYPE", "EXACT", "EXP", "FIND", "FLOOR", "FLOOR.MATH", "HOUR", "INT",
        "ISBLANK", "ISERR", "ISERROR", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISTEXT", "LEFT", "LEN", "LN",
        "LOG", "LOG10", "LOWER", "MID", "MINUTE", "MOD", "MONTH", "MROUND", "NOT", "POWER", "PROPER", "RANDBETWEEN",
        "REPLACE", "REPT", "RIGHT", "ROUND", "ROUNDDOWN", "ROUNDUP", "SEARCH", "SECOND", "SIGN", "SQRT", "SUBSTITUTE",
        "TEXT", "TIME", "TRIM", "TRUNC", "UPPER", "VALUE", "WEEKDAY", "YEAR"), _EVERY),
    "COUNTIF": Places((1,)), "SUMIF": Places((1,)), "AVERAGEIF": Places((1,)),
    "COUNTIFS": Places(start=1, step=2), "SUMIFS": Places(start=2, step=2),
    "MATCH": Places((0, 2)), "LARGE": Places((1,)), "SMALL": Places((1,)),
}
#: Where a function wants one value but takes only the first item of an array: SUM(INDEX({1,2;3,4},{1,2},{1,2}))
#: is 1, even inside SUMPRODUCT. Cells are cut as for the functions above.
_FIRST_ITEM: Final[dict[str, Places]] = {
    "INDEX": Places((1, 2, 3)), "VLOOKUP": Places((0, 2, 3)), "HLOOKUP": Places((0, 2, 3)), "XLOOKUP": Places((0,)),
}
#: Arguments worked out as arrays even in a cell: SUMPRODUCT(LEN(A1:A3)) adds every length, and INDEX(A1:A3*2,2)
#: is A2*2 in any row.
_ARRAYS: Final[dict[str, Places]] = {"SUMPRODUCT": _EVERY, "INDEX": Places((0,))}
#: Arguments that have to be cells. Excel will not take a formula with a value there (error 1004): a number, text,
#: an array, an operator's answer, or a function that answers with a value, SUMIF(LEN(A1:A3),1). A name or a
#: function that can answer with cells is taken, and is #VALUE! when it comes to a value instead.
CELLS: Final[dict[str, Places]] = {
    "SUBTOTAL": Places(start=1), "COUNTIF": Places((0,)), "SUMIF": Places((0, 2)), "AVERAGEIF": Places((0, 2)),
    "COUNTIFS": Places(start=0, step=2), "SUMIFS": Places((0,), start=1, step=2),
    "AVERAGEIFS": Places((0,), start=1, step=2), "COUNTBLANK": Places((0,)), "OFFSET": Places((0,)),
    "ROW": Places((0,)), "COLUMN": Places((0,)), "AREAS": Places((0,)),
}


def _places(table: dict[str, Places], name: str, count: int) -> frozenset[int]:
    places = table.get(name)
    return frozenset() if places is None else places.within(count)


#: Where the functions that work out their own arguments read one value through intersected.
_OWN_PLACES: Final[dict[str, Places]] = {
    "IF": Places((0,)), "IFS": Places(start=0, step=2), "IFERROR": _EVERY, "IFNA": _EVERY, "CHOOSE": Places((0,)),
    "OFFSET": Places(start=1),
}
#: The functions that work their arguments out as a cell does even inside an argument worked out as an array.
AS_CELL: Final = frozenset({"IF", "IFS", "IFERROR", "IFNA", "SWITCH", "CHOOSE"})


def one_value_places(name: str, count: int) -> frozenset[int]:
    """The arguments of a call that want one value, cells there cut to the formula's own row or column."""
    upper = name.upper()
    found = FUNCTIONS.get(upper)
    if found is None:
        return frozenset()
    if not found[1]:
        return _places(_LIFTED, upper, count) | _places(_FIRST_ITEM, upper, count)
    if upper == "SWITCH":
        # The subject and each case; a value, and the default after the last case, can be cells.
        return frozenset({0, *range(1, count - 1, 2)})
    return _places(_OWN_PLACES, upper, count)


def array_places(name: str, count: int) -> frozenset[int]:
    """The arguments of a call worked out as arrays even in a cell."""
    return _places(_ARRAYS, name.upper(), count)


def _with_areas(name: str, values: list[object]) -> list[object]:
    """Arguments holding a union, as the function reads it: all its cells, one area, #VALUE!, or not known."""
    if name in _REFUSES_AREAS:
        raise VALUE
    if name not in _READS_AREAS:
        raise VBAUnsupportedError(f"{name} given several areas at once, (A1:A2,A4), is not implemented")
    if name == "INDEX":
        return values
    return [value.joined() if isinstance(value, Areas) else value for value in values]


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
    return _summed(_numbers(args))


def _summed(numbers: Iterable[float]) -> float:
    """Numbers added as SUM and AVERAGE add them: in order, the last addition set to zero when it cancels.

    Measured bit for bit (scripts/measure_zero_snap.py): SUM(a,-b) is 0
    for every a and b that =a-b sets to 0, while SUM(a,-b,1E-20), whose
    last addition does not cancel, is the plain sum. SUMPRODUCT and
    SUMIF add without it.
    """
    total = before = 0.0
    for number in numbers:
        before, total = total, total + number
    return snapped(before, total)


def _added(numbers: Iterable[float]) -> float:
    """Numbers added as Excel adds them: one after another, each sum rounded to a double.

    Measured bit for bit (scripts/measure_variance.py, 87 sets): SUM,
    AVERAGE, SUMIF, AVERAGEIF, SUMSQ and SUMPRODUCT all add this way, SUM
    and AVERAGE then setting a last addition that cancels to zero
    (_summed); fsum's exact total differs in the last bit for a quarter of
    the sets, and Python's own sum compensates from 3.12 on.
    """
    total = 0.0
    for number in numbers:
        total += number
    return total


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


@function("CEILING")
def fn_ceiling(context: Context, args: list[Any]) -> object:
    number = as_number(_one(args, 0))
    step = as_number(_one(args, 1)) if len(args) > 1 else 1.0
    if step == 0:
        return 0.0
    if (number < 0) != (step < 0) and len(args) > 1:
        raise NUM
    return math.ceil(number / step) * step


@function("CEILING.MATH")
def fn_ceiling_math(context: Context, args: list[Any]) -> object:
    """Up to a multiple of the significance, whatever its sign; a negative number with a mode goes away from zero."""
    number = as_number(_one(args, 0))
    step = abs(as_number(_one(args, 1, 1.0))) if len(args) > 1 else 1.0
    mode = as_number(_one(args, 2, 0.0)) if len(args) > 2 else 0.0
    if step == 0:
        return 0.0
    if number < 0 and mode != 0:
        return -math.ceil(-number / step) * step
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
    base, exponent = as_number(_one(args, 0, BLANK)), as_number(_one(args, 1, BLANK))
    if base == 0 and exponent == 0:
        # Excel has no answer for 0 to the 0, where Python says 1.
        raise NUM
    try:
        result = base**exponent
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
    return _summed(numbers) / len(numbers)


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
    """COUNT counts numbers and walks past an error, where SUM stops at one.

    Inside a range or array only numbers count; an argument given on its
    own counts when it is a number, TRUE or FALSE, text that reads as a
    number, or left out, as in COUNT(,).
    """
    total = 0
    for value in args:
        if isinstance(value, Matrix):
            total += sum(1 for item in value.flat() if isinstance(item, (int, float)) and not isinstance(item, bool))
        elif value is BLANK or isinstance(value, (bool, int, float)) or (
                isinstance(value, str) and text_as_number(value) is not None):
            total += 1
    return float(total)


@function("COUNTA")
def fn_counta(context: Context, args: list[Any]) -> object:
    """Everything that is not blank; an argument left out counts, as in COUNTA(,1)."""
    total = 0
    for value in args:
        if isinstance(value, Matrix):
            total += sum(1 for item in value.flat() if item is not BLANK)
        else:
            total += 1
    return float(total)


@function("COUNTBLANK")
def fn_countblank(context: Context, args: list[Any]) -> object:
    return float(sum(1 for value in _values(args) if value is BLANK or value == ""))


def _variance(args: Sequence[object], *, sample: bool) -> float:
    """The variance as Excel works it out, to the bit.

    Measured bit for bit (scripts/measure_variance.py, 87 sets): Excel
    takes the one-pass sum of squares -- (S2 - S1^2/n)/(n-1) for VAR,
    (n*S2 - S1^2)/(n*n) for VARP, added as SUM adds -- unless it cancels,
    and then the two-pass one, the squares about the mean. It cancels,
    by the rule that fits, when the numerator is under a hundredth of S2
    or the variance under 1e-6; that misses only data smaller than a
    thousandth, 8 answers of 348. STDEV and STDEVP are the square roots.
    """
    numbers = list(_numbers(args))
    count = len(numbers)
    if count < (2 if sample else 1):
        raise DIV0
    divisor = count - 1 if sample else count
    first = _added(numbers)
    second = _added(number * number for number in numbers)
    one_pass = (second - first * first / count) / divisor if sample else \
        (count * second - first * first) / (count * divisor)
    if second and (second - first * first / count) / second >= 0.01 and one_pass >= 1e-6:
        return one_pass
    mean = first / count
    return _added((number - mean) * (number - mean) for number in numbers) / divisor


@function("STDEV", "STDEV.S")
def fn_stdev(context: Context, args: list[Any]) -> object:
    return math.sqrt(_variance(args, sample=True))


@function("STDEVP", "STDEV.P")
def fn_stdevp(context: Context, args: list[Any]) -> object:
    return math.sqrt(_variance(args, sample=False))


@function("VAR", "VAR.S")
def fn_var(context: Context, args: list[Any]) -> object:
    return _variance(args, sample=True)


@function("VARP", "VAR.P")
def fn_varp(context: Context, args: list[Any]) -> object:
    return _variance(args, sample=False)


@function("SUMSQ")
def fn_sumsq(context: Context, args: list[Any]) -> object:
    return _added(number * number for number in _numbers(args))


@function("DEVSQ")
def fn_devsq(context: Context, args: list[Any]) -> object:
    """The squares about the mean, in two passes, as Excel adds them: measured bit for bit."""
    numbers = list(_numbers(args))
    if not numbers:
        raise NUM
    mean = _added(numbers) / len(numbers)
    return _added((number - mean) * (number - mean) for number in numbers)


#: What each of SUBTOTAL's function numbers works out, 101 to 111 as 1 to 11.
_SUBTOTALS: Final = {1: fn_average, 2: fn_count, 3: fn_counta, 4: fn_max, 5: fn_min, 6: fn_product, 7: fn_stdev,
                     8: fn_stdevp, 9: fn_sum, 10: fn_var, 11: fn_varp}


@function("SUBTOTAL", lazy=True)
def fn_subtotal(context: Context, nodes: list[Any]) -> object:
    """SUBTOTAL over the cells of its references, as Excel works it out.

    Measured (scripts/measure_subtotal.py): the function number is read
    as a number and cut to a whole one, TRUE being 1 and "9" 9; outside 1
    to 11 and 101 to 111 it is #VALUE!. 1 to 11 pass over the rows a
    filter hid, 101 to 111 over every hidden row, as the grid says which;
    hidden columns count. A cell whose formula holds SUBTOTAL or
    AGGREGATE is passed over, so subtotals are not counted twice. The
    cells then go to the function as a range would: SUM over an error is
    that error, COUNT walks past it and COUNTA counts it.
    """
    if len(nodes) < 2:
        raise VALUE
    number = single(context.value(nodes[0]))
    if isinstance(number, ExcelError):
        raise number
    value = as_number(number)
    kind = int(value) if value >= 0 else -1
    if not (1 <= kind <= 11 or 101 <= kind <= 111):
        raise VALUE
    values: list[list[object]] = []
    for node in nodes[1:]:
        areas = context.areas_of(node)
        if areas is None:
            # Written as a value, which Excel refuses when the formula is written: only a formula made some other
            # way, a name standing for one, reaches here.
            raise VALUE
        for area in areas:
            sheet = area.sheet or context.sheet
            bounded = clip(area, context.grid.used(sheet))
            passed = context.grid.hidden_rows(sheet, kind > 100)
            for row in range(bounded.top, bounded.bottom + 1):
                if row in passed:
                    continue
                for column in range(bounded.left, bounded.right + 1):
                    if not _subtotalled(context.grid.formula_at(sheet, row, column)):
                        values.append([context.grid.cell_value(sheet, row, column)])
    return _SUBTOTALS[kind % 100](context, [Matrix(values if values else [[BLANK]])])


def _subtotalled(formula: str) -> bool:
    """Whether a cell's formula calls SUBTOTAL or AGGREGATE, which SUBTOTAL passes over."""
    if not formula:
        return False
    tokens = P.tokenize(formula[1:] if formula.startswith("=") else formula)
    return any(token.kind == "name" and token.text.upper() in ("SUBTOTAL", "AGGREGATE") and following.kind == "open"
               for token, following in zip(tokens, tokens[1:]))


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


def _equal(value: object, target: object, *, exact: bool = False) -> bool:
    """Whether a value matches: numbers to fifteen digits, or with ``exact`` bit for bit as the lookups match."""
    if value is BLANK:
        return target is BLANK
    if isinstance(target, str) and isinstance(value, str):
        return value.upper() == target.upper()
    if isinstance(target, str) != isinstance(value, str):
        return False
    try:
        return bool(compare("=", value, target, exact=exact))
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
    return _added(kept) / len(kept)


# --- logic -------------------------------------------------------------------------------------------


def _as_cell(context: Context) -> Context:
    """Where IF, CHOOSE, IFERROR, IFNA, IFS and SWITCH work out their arguments.

    An argument worked out as an array stops at them: inside SUMPRODUCT,
    IF(A1:A3>1,1,0) still reads the one row the formula is on, which is
    why SUMPRODUCT(IF(...)) wants Ctrl+Shift+Enter. A whole array
    formula, as Evaluate works one, reaches through.
    """
    return replace(context, array_argument=False) if context.array_argument else context


@function("IF", lazy=True)
def fn_if(context: Context, nodes: list[Any]) -> object:
    """IF runs only the branch it takes, which is why it is lazy; a block it lands on stays whole.

    The condition wants one value, so in a cell A1:A3>2 reads the row
    the formula is on. An array of conditions takes a branch for each
    item, and the answer is an array: IF({TRUE,FALSE},{1,2},0) is {1,0}.
    """
    if not nodes:
        raise VALUE
    context = _as_cell(context)
    condition = intersected(nodes[0], context)
    if isinstance(condition, Matrix) and not condition.single:
        return _if_each(context, nodes, condition)
    if as_bool(single(condition)):
        return evaluate(nodes[1], context) if len(nodes) > 1 else True
    if len(nodes) > 2:
        return evaluate(nodes[2], context)
    return False


def _if_each(context: Context, nodes: list[Any], conditions: Matrix) -> Matrix:
    """IF over an array of conditions: each item takes its branch, and a branch that is an array gives its item there.

    Both branches are worked out, and the answer is as big as the biggest
    of the three, #N/A past the end of a shorter one: IF({TRUE,FALSE},
    {1,2,3},0) is {1,0,#N/A}, and IF({TRUE,TRUE},1,{1,2,3}) {1,1,#N/A}.
    """
    branches = [_branch(context, nodes, 1), _branch(context, nodes, 2)]
    shapes = [conditions, *(branch for branch in branches if isinstance(branch, Matrix))]
    rows: list[list[object]] = []
    for row in range(max(shape.height for shape in shapes)):
        line: list[object] = []
        for column in range(max(shape.width for shape in shapes)):
            try:
                taken = branches[0] if as_bool(conditions.at(row, column)) else branches[1]
            except ExcelError as failure:
                taken = failure
            line.append(taken.at(row, column) if isinstance(taken, Matrix) else taken)
        rows.append(line)
    return Matrix(rows)


def _branch(context: Context, nodes: list[Any], index: int) -> object:
    """One of IF's branches, worked out whole; one left out is TRUE or FALSE, as IF answers without it."""
    if index >= len(nodes):
        return index == 1
    try:
        return evaluate(nodes[index], context)
    except ExcelError as failure:
        return failure


@refers("IF")
def ref_if(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    if not nodes:
        raise VALUE
    context = _as_cell(context)
    condition = intersected(nodes[0], context)
    if isinstance(condition, Matrix) and not condition.single:
        # A branch for each item is an array, never cells.
        return None
    branch = 1 if as_bool(single(condition)) else 2
    return context.areas_of(nodes[branch]) if branch < len(nodes) else None


@function("IFS", lazy=True)
def fn_ifs(context: Context, nodes: list[Any]) -> object:
    """The value after the first condition that holds, a block kept whole as IF keeps one."""
    context = _as_cell(context)
    return evaluate(nodes[_ifs_taken(context, nodes)], context)


@refers("IFS")
def ref_ifs(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    context = _as_cell(context)
    return context.areas_of(nodes[_ifs_taken(context, nodes)])


def _ifs_taken(context: Context, nodes: list[Any]) -> int:
    for index in range(0, len(nodes) - 1, 2):
        if as_bool(single(intersected(nodes[index], context))):
            return index + 1
    raise NA


@function("IFERROR", lazy=True)
def fn_iferror(context: Context, nodes: list[Any]) -> object:
    """The value, or the fallback when it is an error; of an array given for either, the first item."""
    context = _as_cell(context)
    try:
        value = single(intersected(nodes[0], context))
    except ExcelError:
        value = VALUE
    if isinstance(value, ExcelError):
        return single(intersected(nodes[1], context)) if len(nodes) > 1 else BLANK
    return value


@function("IFNA", lazy=True)
def fn_ifna(context: Context, nodes: list[Any]) -> object:
    context = _as_cell(context)
    try:
        value = single(intersected(nodes[0], context))
    except ExcelError as failure:
        if failure != NA:
            raise
        value = NA
    if value == NA:
        return single(intersected(nodes[1], context)) if len(nodes) > 1 else BLANK
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
    """The value after the first case equal to the subject, or the default; a block kept whole as IF keeps one."""
    context = _as_cell(context)
    return evaluate(nodes[_switch_taken(context, nodes)], context)


@refers("SWITCH")
def ref_switch(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    context = _as_cell(context)
    return context.areas_of(nodes[_switch_taken(context, nodes)])


def _switch_taken(context: Context, nodes: list[Any]) -> int:
    subject = single(intersected(nodes[0], context))
    if isinstance(subject, ExcelError):
        raise subject
    index = 1
    while index + 1 < len(nodes):
        if _equal(subject, single(intersected(nodes[index], context))):
            return index + 1
        index += 2
    if index < len(nodes):
        return index
    raise NA


@function("CHOOSE", lazy=True)
def fn_choose(context: Context, nodes: list[Any]) -> object:
    """The value chosen, a block kept whole: SUM(CHOOSE(2,A1:A2,B1:B3)) adds B1:B3.

    An array of indexes chooses for each item, as IF does for an array of
    conditions: CHOOSE({1,2},A1:A3,B1:B3) is A1:A3 beside B1:B3.
    """
    context = _as_cell(context)
    which = intersected(nodes[0], context) if nodes else VALUE
    if isinstance(which, Matrix) and not which.single:
        return _choose_each(context, nodes, which)
    return evaluate(nodes[_choice(single(which), nodes)], context)


def _choose_each(context: Context, nodes: list[Any], indexes: Matrix) -> Matrix:
    """CHOOSE over an array of indexes: every choice worked out, the answer as big as the biggest of them and the
    indexes, #N/A past the end of a shorter one."""
    choices = [_argument(node, context, one=False) for node in nodes[1:]]
    shapes = [indexes, *(choice for choice in choices if isinstance(choice, Matrix))]
    rows: list[list[object]] = []
    for row in range(max(shape.height for shape in shapes)):
        line: list[object] = []
        for column in range(max(shape.width for shape in shapes)):
            try:
                picked = choices[_choice(indexes.at(row, column), nodes) - 1]
            except ExcelError as failure:
                picked = failure
            line.append(picked.at(row, column) if isinstance(picked, Matrix) else picked)
        rows.append(line)
    return Matrix(rows)


@refers("CHOOSE")
def ref_choose(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    context = _as_cell(context)
    which = intersected(nodes[0], context) if nodes else VALUE
    if isinstance(which, Matrix) and not which.single:
        # A choice for each item is an array, never cells.
        return None
    return context.areas_of(nodes[_choice(single(which), nodes)])


def _choice(which: object, nodes: list[Any]) -> int:
    """Which argument an index picks, 1 the first after it; #VALUE! past either end."""
    number = int(as_number(which))
    if number < 1 or number >= len(nodes):
        raise VALUE
    return number


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
    """The row a lookup lands on. Lookups compare numbers bit for bit, not to fifteen digits (zero_snap.json)."""
    if not approximate:
        test = _criterion(wanted) if isinstance(wanted, str) else None
        for row in range(table.height):
            value = table.at(row, column)
            if test(value) if test is not None else _equal(value, wanted, exact=True):
                return row
        return None
    best: int | None = None
    for row in range(table.height):
        value = table.at(row, column)
        if value is BLANK:
            continue
        try:
            if compare("<=", value, wanted, exact=True):
                best = row
            else:
                break
        except ExcelError:
            continue
    return best


def _find_column(table: Matrix, wanted: object, row: int, approximate: bool) -> int | None:
    if not approximate:
        for column in range(table.width):
            if _equal(table.at(row, column), wanted, exact=True):
                return column
        return None
    best: int | None = None
    for column in range(table.width):
        value = table.at(row, column)
        if value is BLANK:
            continue
        try:
            if compare("<=", value, wanted, exact=True):
                best = column
            else:
                break
        except ExcelError:
            continue
    return best


@function("MATCH")
def fn_match(context: Context, args: list[Any]) -> object:
    wanted = _one(args, 0)
    if isinstance(args[1], ExcelError):
        raise args[1]
    if not isinstance(args[1], Matrix):
        # A value is not something to look in, even one equal to the value sought: Match(3, 3, 0) is #N/A.
        raise NA
    block = args[1]
    kind = _int(args, 2, 1)
    items = list(block.flat())
    if kind == 0:
        test = _criterion(wanted) if isinstance(wanted, str) else None
        for index, value in enumerate(items):
            if test(value) if test is not None else _equal(value, wanted, exact=True):
                return float(index + 1)
        raise NA
    best: int | None = None
    for index, value in enumerate(items):
        if value is BLANK:
            continue
        try:
            fits = compare("<=" if kind > 0 else ">=", value, wanted, exact=True)
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
    first = args[0]
    if isinstance(first, Areas):
        # INDEX((A1:A2,B1:B2),1,1,2): the fourth argument picks the area.
        number = _int(args, 3, 1)
        if number < 1 or number > len(first.blocks):
            raise REF
        first = first.blocks[number - 1]
    block = _matrix(first)
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
        if block.height == 1 and len(args) < 3:
            # Given one index, a single row is counted along.
            if row < 1 or row > block.width:
                raise REF
            return block.rows[0][row - 1]
        if row < 1 or row > block.height:
            raise REF
        return Matrix([list(block.rows[row - 1])])
    if row < 1 or row > block.height or column < 1 or column > block.width:
        raise REF
    return block.rows[row - 1][column - 1]


@refers("INDEX")
def ref_index(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    """INDEX of cells is cells: A1:INDEX(A:A,5) runs to A5, and INDEX(A1:B10,0,2) is B1:B10."""
    areas = context.areas_of(nodes[0]) if len(nodes) >= 2 else None
    if areas is None:
        return None
    number = _node_int(context, nodes, 3, 1)
    if number < 1 or number > len(areas):
        raise REF
    area = areas[number - 1]
    row = _node_int(context, nodes, 1, 0)
    column = _node_int(context, nodes, 2, 0)
    if len(nodes) < 3 and area.rows == 1 and area.columns > 1:
        # Given one index, a single row is counted along.
        row, column = 1, row
    if row < 0 or column < 0 or row > area.rows or column > area.columns:
        raise REF
    top, bottom = (area.top + row - 1,) * 2 if row else (area.top, area.bottom)
    left, right = (area.left + column - 1,) * 2 if column else (area.left, area.right)
    return [Area(top, left, bottom, right, area.sheet)]


def _node_int(context: Context, nodes: list[P.Node], at: int, default: int) -> int:
    """A whole-number argument of a function handed its nodes, cut toward zero; left out, the default.

    It wants one value, so in a cell INDEX(A1:A3,A1:A3) in row 2 reads A2.
    """
    node = nodes[at] if at < len(nodes) else None
    if node is None or (isinstance(node, P.Literal) and node.value is None):
        return default
    value = single(intersected(node, context))
    if isinstance(value, ExcelError):
        raise value
    if value is BLANK:
        return default
    number = as_number(value)
    return int(number) if number >= 0 else -int(-number)


@function("AREAS", lazy=True)
def fn_areas(context: Context, nodes: list[Any]) -> object:
    areas = context.areas_of(nodes[0]) if nodes else None
    if areas is None:
        raise VALUE
    return float(len(areas))


@function("XLOOKUP")
def fn_xlookup(context: Context, args: list[Any]) -> object:
    give = _matrix(args[2])
    index = _found_at(_one(args, 0), _matrix(args[1]))
    if index is not None:
        return give.at(index, 0) if give.width == 1 else give.at(0, index)
    if len(args) > 3 and args[3] is not BLANK:
        return single(args[3])
    raise NA


@refers("XLOOKUP")
def ref_xlookup(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    """XLOOKUP into cells answers with the cell it finds, so COUNTIF(XLOOKUP(2,A1:A3,A1:A3),2) counts one cell."""
    area = context.area_of(nodes[2]) if len(nodes) > 2 else None
    if area is None:
        return None
    index = _found_at(single(intersected(nodes[0], context)), _matrix(evaluate(nodes[1], context)))
    if index is None:
        return context.areas_of(nodes[3]) if len(nodes) > 3 else None
    if area.columns == 1:
        row = area.top + index
        return [Area(row, area.left, row, area.left, area.sheet)] if row <= area.bottom else None
    column = area.left + index
    return [Area(area.top, column, area.top, column, area.sheet)] if column <= area.right else None


def _found_at(wanted: object, where: Matrix) -> int | None:
    """Where XLOOKUP finds a value, counting along the cells it looks in."""
    if isinstance(wanted, ExcelError):
        raise wanted
    test = _criterion(wanted) if isinstance(wanted, str) else None
    for index, value in enumerate(where.flat()):
        if test(value) if test is not None else _equal(value, wanted, exact=True):
            return index
    return None


@function("ROW", lazy=True)
def fn_row(context: Context, nodes: list[Any]) -> object:
    """The row number, or in an array formula or argument every row's: SUM(ROW(A1:A3)) in a cell is 1.

    With no argument it is the formula's own row, which an array formula
    answers as an array of one: Evaluate("ROW()") is {1}.
    """
    if not nodes:
        return Matrix([[float(context.row)]]) if context.array else float(context.row)
    area = context.area_of(nodes[0])
    if area is None:
        raise REF
    if area.rows == 1 or not (context.array or context.array_argument):
        return float(area.top)
    return Matrix([[float(row)] for row in range(area.top, min(area.bottom, area.top + 9999) + 1)])


@function("COLUMN", lazy=True)
def fn_column(context: Context, nodes: list[Any]) -> object:
    """The column number, or in an array formula or argument every column's; with no argument, as ROW()."""
    if not nodes:
        return Matrix([[float(context.column)]]) if context.array else float(context.column)
    area = context.area_of(nodes[0])
    if area is None:
        raise REF
    if area.columns == 1 or not (context.array or context.array_argument):
        return float(area.left)
    return Matrix([[float(column) for column in range(area.left, min(area.right, area.left + 9999) + 1)]])


@function("ROWS", lazy=True)
def fn_rows(context: Context, nodes: list[Any]) -> object:
    """How many rows; anything other than cells is worked out as an array, so ROWS(A1:A3*2) is 3 in any row.
    An error on its own is that error: ROWS(#REF!) is #REF! (tests/fixtures/tables/edits/)."""
    area = context.area_of(nodes[0]) if nodes else None
    if area is not None:
        return float(area.rows)
    return float(_matrix(_not_an_error(evaluate(nodes[0], replace(context, array_argument=True)))).height)


@function("COLUMNS", lazy=True)
def fn_columns(context: Context, nodes: list[Any]) -> object:
    area = context.area_of(nodes[0]) if nodes else None
    if area is not None:
        return float(area.columns)
    return float(_matrix(_not_an_error(evaluate(nodes[0], replace(context, array_argument=True)))).width)


def _not_an_error(value: object) -> object:
    """A value, raised where it is an error on its own; an array holding errors is still an array."""
    if isinstance(value, ExcelError):
        raise value
    return value


@function("OFFSET", lazy=True)
def fn_offset(context: Context, nodes: list[Any]) -> object:
    moved = ref_offset(context, nodes)
    assert moved is not None
    return context.grid.block(context.sheet, moved[0])


@refers("OFFSET")
def ref_offset(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    area = context.area_of(nodes[0]) if nodes else None
    if area is None:
        raise REF
    top = area.top + _node_int(context, nodes, 1, 0)
    left = area.left + _node_int(context, nodes, 2, 0)
    height = _node_int(context, nodes, 3, area.rows)
    width = _node_int(context, nodes, 4, area.columns)
    if top < 1 or left < 1 or height < 1 or width < 1:
        raise REF
    if top + height - 1 > MAX_ROWS or left + width - 1 > MAX_COLUMNS:
        raise REF
    return [Area(top, left, top + height - 1, left + width - 1, area.sheet)]


@function("INDIRECT", lazy=True)
def fn_indirect(context: Context, nodes: list[Any]) -> object:
    found = ref_indirect(context, nodes)
    assert found is not None
    return context.grid.block(context.sheet, found[0])


#: An R1C1 reference counted from the top left, which INDIRECT(text,FALSE) reads: R1C1 or R1C1:R3C1.
_R1C1: Final = re.compile(r"R([0-9]+)C([0-9]+)(?::R([0-9]+)C([0-9]+))?", re.IGNORECASE)


@refers("INDIRECT")
def ref_indirect(context: Context, nodes: list[P.Node]) -> list[Area] | None:
    """The cells a text names: A1 style, a defined name, or R1C1 style when the second argument is FALSE."""
    text = as_text(single(evaluate(nodes[0], context)))
    sheet, body = P.split_sheet(text)
    if sheet and not context.grid.sheet_exists(sheet):
        raise REF
    if len(nodes) > 1 and not as_bool(single(evaluate(nodes[1], context))):
        match = _R1C1.fullmatch(body)
        if match is None:
            if "[" in body or re.fullmatch(r"R\[?-?[0-9]*\]?C\[?-?[0-9]*\]?(?::.*)?", body, re.IGNORECASE):
                raise VBAUnsupportedError(f"INDIRECT of the relative R1C1 reference {text!r} is not implemented")
            raise REF
        top, bottom = sorted((int(match[1]), int(match[3] or match[1])))
        left, right = sorted((int(match[2]), int(match[4] or match[2])))
        if top < 1 or left < 1 or bottom > MAX_ROWS or right > MAX_COLUMNS:
            raise REF
        return [Area(top, left, bottom, right, sheet or context.sheet)]
    try:
        return [context.resolve(P.Reference(text=body.replace("$", ""), sheet=sheet))]
    except ExcelError:
        pass
    named = context.areas_of(P.NameNode(name=body, sheet=sheet)) if re.fullmatch(r"[A-Za-z_\\][\w.\\]*", body) else None
    if named is None:
        raise REF
    return named


@function("TRANSPOSE")
def fn_transpose(context: Context, args: list[Any]) -> object:
    if not isinstance(args[0], Matrix):
        # A single value is its own transpose.
        return args[0]
    block = args[0]
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
    """A value through a number format, as Excel's formats show it; one the format cannot show is #VALUE!."""
    from pyopenvba.formula._display import UndisplayableError, format_value

    value = _one(args, 0, BLANK)
    pattern = as_text(_one(args, 1, BLANK))
    if value is BLANK:
        value = 0.0
    if isinstance(value, str):
        # Text that reads as a number is formatted as the number.
        number = text_as_number(value)
        value = value if number is None else number
    elif not isinstance(value, bool):
        value = as_number(value)
    try:
        return format_value(value, pattern)
    except UndisplayableError:
        raise VALUE from None


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
    return float(_calendar(as_number(_one(args, 0, BLANK)))[0])


@function("MONTH")
def fn_month(context: Context, args: list[Any]) -> object:
    return float(_calendar(as_number(_one(args, 0, BLANK)))[1])


@function("DAY")
def fn_day(context: Context, args: list[Any]) -> object:
    return float(_calendar(as_number(_one(args, 0, BLANK)))[2])


#: The last day Excel has a date for, 31 December 9999.
_LAST_SERIAL: Final = 2958465


def _calendar(serial: float) -> tuple[int, int, int]:
    """The year, month and day Excel shows for a serial number.

    Excel's calendar has a 29 February 1900, serial 60, so every serial
    before it is a day later than counting from 30 December 1899 gives:
    1 is 1 January 1900 (DAY({1,2}) adds to 3), and 0 is day 0 of it.
    """
    if serial < 0 or serial >= _LAST_SERIAL + 1:
        raise NUM
    when = _as_datetime(serial)
    days = (when.date() - EPOCH.date()).days
    if days >= 61:
        return when.year, when.month, when.day
    if days == 60:
        return 1900, 2, 29
    if days == 0:
        return 1900, 1, 0
    when += _dt.timedelta(days=1)
    return when.year, when.month, when.day


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
    """The days between two dates, either of which may be text that reads as one."""
    return float(int(_day(_one(args, 0, BLANK))) - int(_day(_one(args, 1, BLANK))))


def _day(value: object) -> float:
    if isinstance(value, str) and text_as_number(value) is None:
        return _date_value(value)
    return as_number(value)


@function("DATEVALUE")
def fn_datevalue(context: Context, args: list[Any]) -> object:
    return _date_value(as_text(_one(args, 0, BLANK)))


def _date_value(text: str) -> float:
    """The day a text names, or #VALUE! where it names none."""
    from pyopenvba.interpreter._values import parse_date_text

    parsed = parse_date_text(text)
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


@function("ISERROR")
def fn_iserror(context: Context, args: list[Any]) -> object:
    return isinstance(_given(args), ExcelError)


@function("ISERR")
def fn_iserr(context: Context, args: list[Any]) -> object:
    value = _given(args)
    return isinstance(value, ExcelError) and value != NA


@function("ISNA")
def fn_isna(context: Context, args: list[Any]) -> object:
    return _given(args) == NA


def _given(args: list[Any]) -> object:
    """The one value an IS function looks at, an error as much as any other."""
    return single(args[0]) if args else BLANK


@function("ISREF", lazy=True)
def fn_isref(context: Context, nodes: list[Any]) -> object:
    """Whether the argument is cells, in any row: ISREF(A1:A3) is TRUE, ISREF(1) FALSE."""
    try:
        return bool(nodes) and context.areas_of(nodes[0]) is not None
    except ExcelError:
        return False


@function("NA")
def fn_na(context: Context, args: list[Any]) -> object:
    raise NA


@function("ERROR.TYPE")
def fn_error_type(context: Context, args: list[Any]) -> object:
    value = _given(args)
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


def known_names() -> frozenset[str]:
    """Every function this implements, for the inventory to check against."""
    return frozenset(FUNCTIONS)
