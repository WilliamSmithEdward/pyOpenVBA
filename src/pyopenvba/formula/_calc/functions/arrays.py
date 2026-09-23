"""The dynamic-array functions that make and reshape arrays: SEQUENCE,
RANDARRAY, SORT, SORTBY, UNIQUE, FILTER, TAKE, DROP, CHOOSEROWS,
CHOOSECOLS, VSTACK, HSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, EXPAND and
ARRAYTOTEXT.

Their array parameters take a range or an expression whole, even in a
formula written before dynamic arrays: ``FILTER(B2:B11,A2:A11="Apple")``
compares all ten cells, where the same comparison inside SUM would be cut
to the formula's own row. A result with no items at all is ``#CALC!``,
Excel's error for an empty array.
"""

from __future__ import annotations

import random

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.numbers import number_text
from pyopenvba.formula._calc.registry import A, V, function
from pyopenvba.formula._calc.values import (
    CALC,
    MAX_TEXT,
    NA,
    NUM,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Scalar,
    Value,
)
from pyopenvba.formula._calc.cells import CellError

Line = list[Scalar]
Key = tuple[int, object]

# ----------------------------------------------------------------------
# Shared pieces
# ----------------------------------------------------------------------


def _given(value: Value | None) -> bool:
    return value is not None and not isinstance(value, Empty)


def _count(context: Context, value: Scalar | None, default: int) -> int:
    return default if value is None or isinstance(value, Empty) else context.integer(value)


def _flag(context: Context, value: Scalar | None) -> bool:
    return False if value is None or isinstance(value, Empty) else context.logical(value)


def _columns(array: Array) -> list[Line]:
    return [[row[column] for row in array.rows] for column in range(array.width)]


def _lines(array: Array, by_column: bool) -> list[Line]:
    return _columns(array) if by_column else [list(row) for row in array.rows]


def _from_lines(lines: list[Line], by_column: bool) -> Array:
    """An array from its rows, or from its columns when ``by_column``."""
    if not lines:
        raise ExcelError(CALC)
    if not by_column:
        return Array(lines)
    return Array([[line[row] for line in lines] for row in range(len(lines[0]))])


def _sort_key(value: Scalar) -> Key:
    """Where a value sorts: numbers, then text, logicals, errors, blanks."""
    if isinstance(value, bool):
        return 2, value
    if isinstance(value, float):
        return 0, value
    if isinstance(value, str):
        return 1, _collate.sort_key(value)
    if isinstance(value, CellError):
        return 3, value.code
    return 4, 0


def _same_key(value: Scalar) -> Key:
    """What makes two values one for UNIQUE: case never matters, and a
    blank is the empty text."""
    if isinstance(value, Empty):
        return _sort_key("")
    return _sort_key(value)


# ----------------------------------------------------------------------
# Making arrays
# ----------------------------------------------------------------------


@function("SEQUENCE", V, V, V, V, minimum=1)
def SEQUENCE(
    context: Context,
    rows: Scalar,
    columns: Scalar | None = None,
    start: Scalar | None = None,
    step: Scalar | None = None,
) -> Value:
    height = context.integer(rows)
    width = _count(context, columns, 1)
    first = context.number(start) if start is not None and not isinstance(start, Empty) else 1.0
    by = context.number(step) if step is not None and not isinstance(step, Empty) else 1.0
    if height < 0 or width < 0:
        return VALUE
    if height == 0 or width == 0:
        return CALC
    return Array([[checked(first + by * (row * width + column)) for column in range(width)] for row in range(height)])


@function("RANDARRAY", V, V, V, V, V, minimum=0, volatile=True)
def RANDARRAY(
    context: Context,
    rows: Scalar | None = None,
    columns: Scalar | None = None,
    low: Scalar | None = None,
    high: Scalar | None = None,
    whole: Scalar | None = None,
) -> Value:
    height = _count(context, rows, 1)
    width = _count(context, columns, 1)
    bottom = context.number(low) if low is not None and not isinstance(low, Empty) else 0.0
    top = context.number(high) if high is not None and not isinstance(high, Empty) else 1.0
    if height < 0 or width < 0 or bottom > top:
        return VALUE
    if height == 0 or width == 0:
        return CALC
    if _flag(context, whole):
        if not (bottom.is_integer() and top.is_integer()):
            return VALUE
        return Array([[float(random.randint(int(bottom), int(top))) for _ in range(width)] for _ in range(height)])
    return Array([[bottom + random.random() * (top - bottom) for _ in range(width)] for _ in range(height)])


# ----------------------------------------------------------------------
# Sorting, uniqueness, filtering
# ----------------------------------------------------------------------


def _orders(context: Context, value: Value | None, count: int) -> list[int]:
    """The sort orders given, 1 or -1, one for each of ``count`` keys."""
    if not _given(value):
        return [1] * count
    assert value is not None
    found = [context.integer(item) for item in matrix(context, value).items()]
    if any(order not in (1, -1) for order in found):
        raise ExcelError(VALUE)
    if len(found) == 1:
        return found * count
    if len(found) != count:
        raise ExcelError(VALUE)
    return found


def _sorted(lines: list[Line], keys: list[Line], orders: list[int]) -> list[Line]:
    """Lines put in order by their keys, each ascending or descending, the
    first key first; lines with equal keys keep their order."""
    positions = list(range(len(lines)))
    for index in reversed(range(len(orders))):
        positions.sort(key=lambda position: _sort_key(keys[position][index]), reverse=orders[index] < 0)
    return [lines[position] for position in positions]


@function("SORT", A, A, A, V, minimum=1)
def SORT(
    context: Context,
    array: Value,
    index: Value | None = None,
    order: Value | None = None,
    by_column: Scalar | None = None,
) -> Value:
    grid = matrix(context, array)
    across = _flag(context, by_column)
    lines = _lines(grid, across)
    size = grid.height if across else grid.width
    indexes = [1] if index is None or not _given(index) else [context.integer(i) for i in matrix(context, index).items()]
    if any(not 1 <= which <= size for which in indexes):
        return VALUE
    orders = _orders(context, order, len(indexes))
    keys = [[line[which - 1] for which in indexes] for line in lines]
    return _from_lines(_sorted(lines, keys, orders), across)


@function("SORTBY", A, A, A, repeat=2, minimum=2, maximum=253)
def SORTBY(context: Context, array: Value, *pairs: Value) -> Value:
    grid = matrix(context, array)
    columns: list[Line] = []
    orders: list[int] = []
    across: bool | None = None
    for position in range(0, len(pairs), 2):
        key = matrix(context, pairs[position])
        order = pairs[position + 1] if position + 1 < len(pairs) else None
        if key.width == 1 and key.height == grid.height:
            direction = False
        elif key.height == 1 and key.width == grid.width:
            direction = True
        else:
            return VALUE
        if across is not None and direction != across:
            return VALUE
        across = direction
        columns.append(list(key.items()))
        orders.extend(_orders(context, order, 1))
    lines = _lines(grid, bool(across))
    keys = [[column[line] for column in columns] for line in range(len(lines))]
    return _from_lines(_sorted(lines, keys, orders), bool(across))


@function("UNIQUE", A, V, V, minimum=1)
def UNIQUE(context: Context, array: Value, by_column: Scalar | None = None, once: Scalar | None = None) -> Value:
    grid = matrix(context, array)
    across = _flag(context, by_column)
    exactly = _flag(context, once)
    lines = _lines(grid, across)
    keys = [tuple(_same_key(value) for value in line) for line in lines]
    counts: dict[tuple[Key, ...], int] = {}
    for key in keys:
        counts[key] = counts.get(key, 0) + 1
    kept: list[Line] = []
    seen: set[tuple[Key, ...]] = set()
    for line, key in zip(lines, keys, strict=True):
        if key in seen or (exactly and counts[key] > 1):
            continue
        seen.add(key)
        kept.append(line)
    return _from_lines(kept, across)


def _keeps(context: Context, flag: Scalar) -> bool:
    """Whether FILTER keeps a line: TRUE or a number not 0 does; text is
    ``#VALUE!`` and an error is itself."""
    if isinstance(flag, CellError):
        raise ExcelError(flag)
    if isinstance(flag, str):
        raise ExcelError(VALUE)
    return context.logical(flag)


@function("FILTER", A, A, A, minimum=2)
def FILTER(context: Context, array: Value, include: Value, empty: Value | None = None) -> Value:
    grid = matrix(context, array)
    keep = matrix(context, include)
    if keep.width == 1 and keep.height == grid.height:
        across = False
        flags = [row[0] for row in keep.rows]
    elif keep.height == 1 and keep.width == grid.width:
        across = True
        flags = list(keep.rows[0])
    else:
        return VALUE
    kept = [line for line, flag in zip(_lines(grid, across), flags, strict=True) if _keeps(context, flag)]
    if not kept:
        return empty if empty is not None and _given(empty) else CALC
    return _from_lines(kept, across)


# ----------------------------------------------------------------------
# Taking and dropping
# ----------------------------------------------------------------------


def _span(size: int, count: int, *, take: bool) -> range:
    """The positions TAKE keeps, or DROP leaves, of ``size``: a negative
    count counts from the end."""
    if take:
        return range(max(size + count, 0), size) if count < 0 else range(min(count, size))
    return range(max(size + count, 0)) if count < 0 else range(min(count, size), size)


def _cut(context: Context, array: Value, rows: Scalar | None, columns: Scalar | None, *, take: bool) -> Value:
    grid = matrix(context, array)
    down = range(grid.height)
    across = range(grid.width)
    if rows is not None and not isinstance(rows, Empty):
        down = _span(grid.height, context.integer(rows), take=take)
    if columns is not None and not isinstance(columns, Empty):
        across = _span(grid.width, context.integer(columns), take=take)
    if not down or not across:
        return CALC
    return Array([[grid.rows[row][column] for column in across] for row in down])


@function("TAKE", A, V, V, minimum=2)
def TAKE(context: Context, array: Value, rows: Scalar | None = None, columns: Scalar | None = None) -> Value:
    return _cut(context, array, rows, columns, take=True)


@function("DROP", A, V, V, minimum=2)
def DROP(context: Context, array: Value, rows: Scalar | None = None, columns: Scalar | None = None) -> Value:
    return _cut(context, array, rows, columns, take=False)


def _choose(context: Context, array: Value, picks: tuple[Value, ...], *, columns: bool) -> Value:
    lines = _lines(matrix(context, array), columns)
    chosen: list[Line] = []
    for pick in picks:
        for item in matrix(context, pick).items():
            index = context.integer(item)
            if index == 0 or abs(index) > len(lines):
                return VALUE
            chosen.append(lines[index - 1] if index > 0 else lines[len(lines) + index])
    return _from_lines(chosen, columns)


@function("CHOOSEROWS", A, A, maximum=254)
def CHOOSEROWS(context: Context, array: Value, *picks: Value) -> Value:
    return _choose(context, array, picks, columns=False)


@function("CHOOSECOLS", A, A, maximum=254)
def CHOOSECOLS(context: Context, array: Value, *picks: Value) -> Value:
    return _choose(context, array, picks, columns=True)


# ----------------------------------------------------------------------
# Stacking and reshaping
# ----------------------------------------------------------------------


@function("VSTACK", A, maximum=254)
def VSTACK(context: Context, *arrays: Value) -> Value:
    grids = [matrix(context, array) for array in arrays]
    width = max(grid.width for grid in grids)
    return Array([list(row) + [NA] * (width - len(row)) for grid in grids for row in grid.rows])


@function("HSTACK", A, maximum=254)
def HSTACK(context: Context, *arrays: Value) -> Value:
    grids = [matrix(context, array) for array in arrays]
    height = max(grid.height for grid in grids)
    rows: list[Line] = [[] for _ in range(height)]
    for grid in grids:
        for index in range(height):
            rows[index].extend(grid.rows[index] if index < grid.height else [NA] * grid.width)
    return Array(rows)


def _flattened(context: Context, array: Value, ignore: Scalar | None, by_column: Scalar | None) -> Line:
    grid = matrix(context, array)
    skip = _count(context, ignore, 0)
    if not 0 <= skip <= 3:
        raise ExcelError(VALUE)
    lines = _columns(grid) if _flag(context, by_column) else grid.rows
    found = [
        value
        for line in lines
        for value in line
        if not ((skip & 1 and isinstance(value, Empty)) or (skip & 2 and isinstance(value, CellError)))
    ]
    if not found:
        raise ExcelError(CALC)
    return found


@function("TOCOL", A, V, V, minimum=1)
def TOCOL(context: Context, array: Value, ignore: Scalar | None = None, by_column: Scalar | None = None) -> Value:
    return Array([[value] for value in _flattened(context, array, ignore, by_column)])


@function("TOROW", A, V, V, minimum=1)
def TOROW(context: Context, array: Value, ignore: Scalar | None = None, by_column: Scalar | None = None) -> Value:
    return Array([_flattened(context, array, ignore, by_column)])


def _filler(context: Context, pad: Value | None) -> Scalar:
    return NA if pad is None or not _given(pad) else context.first(pad)


def _wrap(context: Context, vector: Value, count: Scalar, pad: Value | None, *, columns: bool) -> Value:
    grid = matrix(context, vector)
    if grid.height != 1 and grid.width != 1:
        return VALUE
    size = context.integer(count)
    if size < 1:
        return NUM
    items = list(grid.items())
    lines = [items[start : start + size] for start in range(0, len(items), size)]
    lines[-1] = lines[-1] + [_filler(context, pad)] * (size - len(lines[-1]))
    return _from_lines(lines, columns)


@function("WRAPROWS", A, V, A, minimum=2)
def WRAPROWS(context: Context, vector: Value, count: Scalar, pad: Value | None = None) -> Value:
    return _wrap(context, vector, count, pad, columns=False)


@function("WRAPCOLS", A, V, A, minimum=2)
def WRAPCOLS(context: Context, vector: Value, count: Scalar, pad: Value | None = None) -> Value:
    return _wrap(context, vector, count, pad, columns=True)


@function("EXPAND", A, V, V, A, minimum=2)
def EXPAND(
    context: Context,
    array: Value,
    rows: Scalar | None = None,
    columns: Scalar | None = None,
    pad: Value | None = None,
) -> Value:
    grid = matrix(context, array)
    height = _count(context, rows, grid.height)
    width = _count(context, columns, grid.width)
    if height < grid.height or width < grid.width:
        return VALUE
    filler = _filler(context, pad)
    out = [list(row) + [filler] * (width - grid.width) for row in grid.rows]
    out.extend([filler] * width for _ in range(height - grid.height))
    return Array(out)


# ----------------------------------------------------------------------
# ARRAYTOTEXT
# ----------------------------------------------------------------------


def _item_text(value: Scalar, *, strict: bool) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return number_text(value)
    if isinstance(value, CellError):
        return value.code
    if isinstance(value, str):
        return '"' + value.replace('"', '""') + '"' if strict else value
    return ""


@function("ARRAYTOTEXT", A, V, minimum=1)
def ARRAYTOTEXT(context: Context, array: Value, form: Scalar | None = None) -> Value:
    """All of an array as text: concise, ``1, 2, 3, 4``, or strict,
    ``{1,2;3,4}`` with text in quotes."""
    grid = matrix(context, array)
    strict = _count(context, form, 0)
    if strict not in (0, 1):
        return VALUE
    if strict:
        text = "{" + ";".join(",".join(_item_text(item, strict=True) for item in row) for row in grid.rows) + "}"
    else:
        text = ", ".join(_item_text(item, strict=False) for item in grid.items())
    return VALUE if len(text) > MAX_TEXT else text


__all__: list[str] = []
