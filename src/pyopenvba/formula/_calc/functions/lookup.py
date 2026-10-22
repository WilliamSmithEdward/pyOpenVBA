"""Lookup and reference functions.

An exact match compares as ``=`` does, case aside, and with text VLOOKUP,
HLOOKUP and MATCH read ``*``, ``?`` and ``~`` as wildcards: MATCH("i*",...)
finds "ink". A number never matches text that looks like it.

An approximate match is a binary search for the last value not above the
one sought, among values of its own type, so a text heading above a column
of numbers does not throw it off. On data that is not sorted Excel's own
search can land elsewhere; on sorted data the two agree.

INDEX, OFFSET, INDIRECT and CHOOSE give a reference back, so
``SUM(INDEX(A1:C9,0,2))`` adds up a column.
"""

from __future__ import annotations

import re

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import AreaReference, AxisReference, CellReference, Missing, NameReference, Node
from pyopenvba.formula._calc.parser import parse
from pyopenvba.formula._calc.registry import LAZY, REF, A, R, V, function
from pyopenvba.formula._calc.values import (
    NA,
    VALUE,
    Area,
    Array,
    Empty,
    ExcelError,
    Reference,
    Scalar,
    Value,
    compare,
)
from pyopenvba.formula._calc.values import (
    REF as REF_ERROR,
)
from pyopenvba.formula._calc.host import quote_sheet_name
from pyopenvba.formula._calc.reference import MAX_COLUMN, MAX_ROW, column_letter
from pyopenvba.formula._calc.cells import CellError


def _same_kind(first: Scalar, second: Scalar) -> bool:
    if isinstance(first, bool) or isinstance(second, bool):
        return isinstance(first, bool) and isinstance(second, bool)
    return isinstance(first, str) == isinstance(second, str) and not isinstance(second, (Empty, CellError))


def _order(value: Scalar, target: Scalar) -> int:
    """How a lookup orders a value against what it looks for: numbers to the bit, where = reads fifteen digits --
    MATCH, VLOOKUP, HLOOKUP and XLOOKUP pass over a double one unit in the last place from the one they look for,
    exact or approximate (pyOpenVBA's tests/fixtures/zero_snap.json) -- and text and logicals as compare has them."""
    if isinstance(value, float) and isinstance(target, float):
        return (value > target) - (value < target)
    return compare(value, target)


def _exact(values: list[Scalar], target: Scalar, *, wildcards: bool, last: bool = False) -> int | None:
    """The position of the first value equal to ``target``, or with
    ``last`` the last one."""
    pattern = wildcards and isinstance(target, str) and _collate.has_wildcards(target)
    order = range(len(values) - 1, -1, -1) if last else range(len(values))
    for index in order:
        value = values[index]
        if pattern:
            if isinstance(value, str) and isinstance(target, str) and _collate.wildcard_match(target, value):
                return index
            continue
        if _same_kind(target, value) and _order(value, target) == 0:
            return index
    return None


def _first_cell(context: Context, line: Area, target: Scalar) -> int | None:
    """Where the first cell of one row or column of cells equal to ``target`` is along it, the cells read in order
    and none past it, as Excel reads them for an exact lookup: MATCH(3,A1:A10,0) in A5 stops at A3, so reads no A5
    and closes no loop (pyOpenVBA's tests/fixtures/self_reference.json). A blank cell matches nothing, as in
    :func:`_exact`."""
    pattern = isinstance(target, str) and _collate.has_wildcards(target)
    for row, column, value in context.book.cells(line):
        if pattern:
            found = isinstance(value, str) and isinstance(target, str) and _collate.wildcard_match(target, value)
        else:
            found = _same_kind(target, value) and _order(value, target) == 0
        if found:
            return row - line.top if line.width == 1 else column - line.left
    return None


def _approximate(values: list[Scalar], target: Scalar, *, descending: bool = False) -> int | None:
    """Binary search, among values of ``target``'s type, for the last not
    above it, or with ``descending`` the last not below it."""
    positions = [index for index, value in enumerate(values) if _same_kind(target, value)]
    low, high = 0, len(positions) - 1
    found: int | None = None
    while low <= high:
        middle = (low + high) // 2
        order = _order(values[positions[middle]], target)
        if (order <= 0) if not descending else (order >= 0):
            found = positions[middle]
            low = middle + 1
        else:
            high = middle - 1
    return found


def _lookup_value(value: Scalar) -> Scalar:
    if isinstance(value, CellError):
        raise ExcelError(value)
    return value


@function("VLOOKUP", V, R, V, V, minimum=3, legacy_first=(0, 2, 3))
def VLOOKUP(context: Context, value: Scalar, table: Value, column: Scalar, approximate: Scalar | None = None) -> Value:
    return _table_lookup(context, value, table, column, approximate, vertical=True)


@function("HLOOKUP", V, R, V, V, minimum=3, legacy_first=(0, 2, 3))
def HLOOKUP(context: Context, value: Scalar, table: Value, row: Scalar, approximate: Scalar | None = None) -> Value:
    return _table_lookup(context, value, table, row, approximate, vertical=False)


def _table_lookup(
    context: Context, value: Scalar, table: Value, index: Scalar, approximate: Scalar | None, *, vertical: bool
) -> Value:
    target = _lookup_value(value)
    which = int(context.number(index))
    exact = approximate is not None and not context.logical(approximate)
    area = table.area if isinstance(table, Reference) else None
    if exact and area is not None:
        if which < 1:
            return VALUE
        if which > (area.width if vertical else area.height):
            return REF_ERROR
        keys = Area(area.sheet, area.top, area.left, area.bottom if vertical else area.top,
                    area.left if vertical else area.right)
        position = _first_cell(context, keys, target)
        if position is None:
            return NA
        if vertical:
            return context.book.cell(area.sheet, area.top + position, area.left + which - 1)
        return context.book.cell(area.sheet, area.top + which - 1, area.left + position)
    grid = matrix(context, table)
    lines = grid.rows if vertical else [list(column) for column in zip(*grid.rows, strict=True)]
    if which < 1:
        return VALUE
    if which > (grid.width if vertical else grid.height):
        return REF_ERROR
    keys = [line[0] for line in lines]
    position = _exact(keys, target, wildcards=True) if exact else _approximate(keys, target)
    if position is None:
        return NA
    return lines[position][which - 1]


def _vector(context: Context, value: Value) -> list[Scalar]:
    grid = matrix(context, value)
    if grid.height != 1 and grid.width != 1:
        raise ExcelError(NA)
    return list(grid.items())


@function("MATCH", V, R, V, minimum=2)
def MATCH(context: Context, value: Scalar, array: Value, kind: Scalar | None = None) -> Value:
    if isinstance(array, CellError):
        return array
    if not isinstance(array, (Reference, Array)):
        # A value where the cells or the array go is nothing to look through: MATCH(3,3,0) is #N/A
        # (tests/fixtures/formula/).
        return NA
    target = _lookup_value(value)
    mode = 1.0 if kind is None else context.number(kind)
    area = array.area if isinstance(array, Reference) else None
    if mode == 0 and area is not None and (area.height == 1 or area.width == 1):
        position = _first_cell(context, area, target)
        return NA if position is None else float(position + 1)
    values = _vector(context, array)
    if mode == 0:
        position = _exact(values, target, wildcards=True)
    else:
        position = _approximate(values, target, descending=mode < 0)
    return NA if position is None else float(position + 1)


@function("XMATCH", V, R, V, V, minimum=2)
def XMATCH(
    context: Context, value: Scalar, array: Value, match_mode: Scalar | None = None, search_mode: Scalar | None = None
) -> Value:
    values = _vector(context, array)
    position = _x_search(context, values, _lookup_value(value), match_mode, search_mode)
    return NA if position is None else float(position + 1)


def _x_search(
    context: Context, values: list[Scalar], target: Scalar, match_mode: Scalar | None, search_mode: Scalar | None
) -> int | None:
    mode = 0 if match_mode is None or isinstance(match_mode, Empty) else int(context.number(match_mode))
    search = 1 if search_mode is None or isinstance(search_mode, Empty) else int(context.number(search_mode))
    if mode not in (-1, 0, 1, 2) or search not in (-2, -1, 1, 2):
        raise ExcelError(VALUE)
    if search in (2, -2):
        if mode == 0:
            found = _approximate(values, target, descending=search == -2)
            return found if found is not None and _order(values[found], target) == 0 else None
        if mode == -1:
            return _approximate(values, target, descending=search == -2)
    exact = _exact(values, target, wildcards=mode == 2, last=search == -1)
    if exact is not None or mode in (0, 2):
        return exact
    best: int | None = None
    order = range(len(values) - 1, -1, -1) if search == -1 else range(len(values))
    for index in order:
        candidate = values[index]
        if not _same_kind(target, candidate):
            continue
        direction = _order(candidate, target)
        if (mode == -1 and direction < 0) or (mode == 1 and direction > 0):
            if best is None:
                best = index
                continue
            closer = _order(candidate, values[best])
            if (mode == -1 and closer > 0) or (mode == 1 and closer < 0):
                best = index
    return best


@function("XLOOKUP", V, R, R, LAZY, V, V, minimum=3, legacy_first=(0,))
def XLOOKUP(
    context: Context,
    value: Scalar,
    lookup: Value,
    result: Value,
    if_not_found: Node | None = None,
    match_mode: Scalar | None = None,
    search_mode: Scalar | None = None,
) -> Value:
    keys = matrix(context, lookup)
    if keys.height != 1 and keys.width != 1:
        return VALUE
    position = _x_search(context, list(keys.items()), _lookup_value(value), match_mode, search_mode)
    if position is None:
        # Left empty, as in XLOOKUP(x,a,b,,1), it is not given at all.
        if if_not_found is None or isinstance(if_not_found, Missing):
            return NA
        return context.evaluate(if_not_found)
    return _slice(context, result, position, vertical=keys.width == 1)


def _slice(context: Context, source: Value, position: int, *, vertical: bool) -> Value:
    """The row, or column, of ``source`` at ``position``."""
    if isinstance(source, Reference) and source.area is not None:
        area = source.area
        if vertical:
            if position >= area.height:
                return REF_ERROR
            return Reference.of(Area(area.sheet, area.top + position, area.left, area.top + position, area.right))
        if position >= area.width:
            return REF_ERROR
        return Reference.of(Area(area.sheet, area.top, area.left + position, area.bottom, area.left + position))
    grid = matrix(context, source)
    if vertical:
        if position >= grid.height:
            return REF_ERROR
        row = grid.rows[position]
        return row[0] if len(row) == 1 else Array([list(row)])
    if position >= grid.width:
        return REF_ERROR
    column = [[line[position]] for line in grid.rows]
    return column[0][0] if len(column) == 1 else Array(column)


@function("LOOKUP", V, R, R, minimum=2)
def LOOKUP(context: Context, value: Scalar, lookup: Value, result: Value | None = None) -> Value:
    target = _lookup_value(value)
    grid = matrix(context, lookup)
    if result is None:
        if grid.width > grid.height:
            keys, answers = grid.rows[0], grid.rows[-1]
        else:
            keys, answers = [row[0] for row in grid.rows], [row[-1] for row in grid.rows]
    else:
        if grid.height != 1 and grid.width != 1:
            return NA
        keys = list(grid.items())
        answers = list(matrix(context, result).items())
    position = _approximate(keys, target)
    if position is None:
        return NA
    return answers[position] if position < len(answers) else NA


@function("INDEX", A, V, V, V, minimum=1, legacy_first=(1, 2, 3), reference_result=True)
def INDEX(
    context: Context, array: Value, row: Scalar | None = None, column: Scalar | None = None, area: Scalar | None = None
) -> Value:
    down = 0 if row is None or isinstance(row, Empty) else int(context.number(row))
    across = 0 if column is None or isinstance(column, Empty) else int(context.number(column))
    if down < 0 or across < 0:
        return VALUE
    if isinstance(array, Reference):
        which = 1 if area is None else int(context.number(area))
        if not 1 <= which <= len(array.areas):
            return REF_ERROR
        block = array.areas[which - 1]
        if column is None and block.height == 1:
            # One index into a row counts along it.
            down, across = 1, down
        elif column is None and block.width > 1:
            # One index into cells in rows and columns is #REF!, even 0, where a column written and left empty is
            # the whole row: INDEX(A1:B2,2,) (tests/fixtures/formula/).
            return REF_ERROR
        if down > block.height or across > block.width:
            return REF_ERROR
        top, bottom = (block.top, block.bottom) if down == 0 else (block.top + down - 1,) * 2
        left, right = (block.left, block.right) if across == 0 else (block.left + across - 1,) * 2
        return Reference((Area(block.sheet, top, left, bottom, right),), indirect=array.indirect)
    grid = matrix(context, array)
    if column is None and grid.height == 1:
        down, across = 1, down
    if down > grid.height or across > grid.width:
        return REF_ERROR
    if down and across:
        return grid.rows[down - 1][across - 1]
    if down:
        return Array([list(grid.rows[down - 1])])
    if across:
        return Array([[line[across - 1]] for line in grid.rows])
    return grid


@function("OFFSET", REF, V, V, V, V, minimum=3, reference_result=True)
def OFFSET(
    context: Context,
    reference: Reference,
    rows: Scalar,
    columns: Scalar,
    height: Scalar | None = None,
    width: Scalar | None = None,
) -> Value:
    area = reference.area
    # Measured: a closed workbook's range, read from its link, is #VALUE!.
    if area is None or area.external:
        return VALUE
    down = int(context.number(rows))
    across = int(context.number(columns))
    tall = area.height if height is None or isinstance(height, Empty) else int(context.number(height))
    wide = area.width if width is None or isinstance(width, Empty) else int(context.number(width))
    if tall == 0 or wide == 0:
        return REF_ERROR
    top = area.top + down
    left = area.left + across
    bottom = top + tall - 1 if tall > 0 else top
    right = left + wide - 1 if wide > 0 else left
    if tall < 0:
        top = top + tall + 1
    if wide < 0:
        left = left + wide + 1
    if top < 1 or left < 1 or bottom > MAX_ROW or right > MAX_COLUMN:
        return REF_ERROR
    return Reference((Area(area.sheet, top, left, bottom, right),), indirect=reference.indirect)


_R1C1 = re.compile(
    r"(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?R(?P<row>\[-?\d+\]|\d+)?C(?P<column>\[-?\d+\]|\d+)?"
    r"(?::R(?P<row2>\[-?\d+\]|\d+)?C(?P<column2>\[-?\d+\]|\d+)?)?",
    re.IGNORECASE,
)


def _r1c1(text: str, context: Context) -> Node | None:
    """An R1C1 reference, a cell or the block between two, as the node its A1 spelling would parse to:
    INDIRECT("R1C1:R3C1",FALSE) is A1:A3 (tests/fixtures/reference_forms.json)."""
    match = _R1C1.fullmatch(text)
    if match is None:
        return None

    def resolve(part: str | None, base: int) -> int:
        if part is None:
            return base
        if part.startswith("["):
            return base + int(part[1:-1])
        return int(part)

    corners: list[str] = []
    for row_part, column_part in (("row", "column"), ("row2", "column2")):
        if row_part == "row2" and match.group(row_part) is None and match.group(column_part) is None \
                and ":" not in text:
            break
        row = resolve(match.group(row_part), context.row)
        column = resolve(match.group(column_part), context.column)
        if not (1 <= row <= MAX_ROW and 1 <= column <= MAX_COLUMN):
            return None
        corners.append(f"{column_letter(column)}{row}")
    sheet = match.group("sheet")
    prefix = sheet + "!" if sheet else ""
    return parse(prefix + ":".join(corners))


@function("INDIRECT", V, V, minimum=1, volatile=True, reference_result=True)
def INDIRECT(context: Context, text: Scalar, a1: Scalar | None = None) -> Value:
    source = context.text(text).strip()
    style_a1 = a1 is None or isinstance(a1, Empty) or context.logical(a1)
    try:
        node = parse(source) if style_a1 else _r1c1(source, context)
    except FormulaSyntaxError:
        return REF_ERROR
    if not isinstance(node, (CellReference, AreaReference, AxisReference, NameReference)):
        return REF_ERROR
    # Measured: text naming another workbook's cells is #REF! while that
    # workbook is closed, as it is to everything here, and so is a 3D
    # reference, even ThreeA:ThreeA!A1 (pyOpenVBA's
    # tests/fixtures/three_d_references.json).
    if node.prefix is not None and (node.prefix.book is not None or node.prefix.last_sheet is not None):
        return REF_ERROR
    found = context.evaluate(node)
    return Reference(found.areas, indirect=True) if isinstance(found, Reference) else REF_ERROR


@function("ADDRESS", V, V, V, V, V, minimum=2)
def ADDRESS(
    context: Context,
    row: Scalar,
    column: Scalar,
    absolute: Scalar | None = None,
    a1: Scalar | None = None,
    sheet: Scalar | None = None,
) -> Value:
    down = int(context.number(row))
    across = int(context.number(column))
    kind = 1 if absolute is None or isinstance(absolute, Empty) else int(context.number(absolute))
    if not (1 <= down <= MAX_ROW and 1 <= across <= MAX_COLUMN) or kind not in (1, 2, 3, 4):
        return VALUE
    style_a1 = a1 is None or isinstance(a1, Empty) or context.logical(a1)
    fixed_row = kind in (1, 2)
    fixed_column = kind in (1, 3)
    if style_a1:
        text = f"{'$' if fixed_column else ''}{column_letter(across)}{'$' if fixed_row else ''}{down}"
    else:
        text = f"R{down if fixed_row else f'[{down}]'}C{across if fixed_column else f'[{across}]'}"
    if sheet is not None and not isinstance(sheet, Empty):
        text = quote_sheet_name(context.text(sheet)) + "!" + text
    return text


# pyOpenVBA's legacy_cell marker works TRANSPOSE's range out as a cell's value in a legacy formula, an expression
# over a range included: SUM(TRANSPOSE(A1:A3*2)) in row 5 is #VALUE! (tests/fixtures/formula/).
@function("TRANSPOSE", A, legacy_cell=(0,))
def TRANSPOSE(context: Context, array: Value) -> Value:
    """Measured: in a formula written before dynamic arrays, a range is cut
    to the formula's own row or column first, as a single value would be,
    so ``=SUM(TRANSPOSE(A1:A3))`` in row 5 is ``#VALUE!``."""
    if isinstance(array, Reference) and not context.array:
        array = context.implicit(array)
    grid = matrix(context, array)
    return Array([list(column) for column in zip(*grid.rows, strict=True)])


@function("SINGLE", R)
def SINGLE(context: Context, value: Value) -> Value:
    return context.first(value)


@function("TRIMRANGE", R, V, V, minimum=1)
def TRIMRANGE(context: Context, value: Value, rows: Scalar | None = None, columns: Scalar | None = None) -> Value:
    """A range less its blank rows and columns at either end, or at the
    start (1) or the end (2) only; 0 keeps them. Measured: an array loses
    its blank items the same way, and one all blank is ``#VALUE!`` where a
    range all blank is ``#REF!``."""
    if isinstance(value, Reference) and value.area is None:
        return VALUE
    trim_rows = 3 if rows is None or isinstance(rows, Empty) else context.integer(rows)
    trim_columns = 3 if columns is None or isinstance(columns, Empty) else context.integer(columns)
    if trim_rows not in (0, 1, 2, 3) or trim_columns not in (0, 1, 2, 3):
        return VALUE
    if isinstance(value, Reference) and value.area is not None:
        area = value.area
        filled = [(row, column) for row, column, _ in context.book.cells(area)]
        if not filled:
            return REF_ERROR
        top, left, bottom, right = _trimmed(filled, (area.top, area.left, area.bottom, area.right), trim_rows, trim_columns)
        return Reference.of(Area(area.sheet, top, left, bottom, right))
    grid = matrix(context, value)
    filled = [(row, column) for row, line in enumerate(grid.rows) for column, item in enumerate(line) if not isinstance(item, Empty)]
    if not filled:
        return VALUE
    top, left, bottom, right = _trimmed(filled, (0, 0, grid.height - 1, grid.width - 1), trim_rows, trim_columns)
    return Array([line[left : right + 1] for line in grid.rows[top : bottom + 1]])


def _trimmed(filled: list[tuple[int, int]], edges: tuple[int, int, int, int], rows: int, columns: int) -> tuple[int, int, int, int]:
    """The edges, top, left, bottom and right, moved in to the filled
    positions on the sides TRIMRANGE's modes name: 1 the start, 2 the end,
    3 both."""
    top, left, bottom, right = edges
    if rows & 1:
        top = min(row for row, _ in filled)
    if rows & 2:
        bottom = max(row for row, _ in filled)
    if columns & 1:
        left = min(column for _, column in filled)
    if columns & 2:
        right = max(column for _, column in filled)
    return top, left, bottom, right


@function("ANCHORARRAY", R)
def ANCHORARRAY(context: Context, reference: Value) -> Value:
    """The block a formula spilled into, which a file writes for ``A1#``.
    Measured: a reference that is not one cell holding a spilled formula,
    such as ``ANCHORARRAY(A1:A20)``, is ``#REF!``."""
    return context.spilled(reference)


__all__: list[str] = []
