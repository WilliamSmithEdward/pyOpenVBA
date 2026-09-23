"""IF and the other logical functions.

IF, IFS, SWITCH and CHOOSE evaluate only the branch they pick, and give it
back as it is, so ``SUM(IF(A1,B1:B5,C1:C5))`` adds up a range. Given an
array to test, IF picks item by item instead.

AND and OR skip text and blanks inside a range but not a text argument
that is not TRUE or FALSE, and with nothing logical to look at at all they
are ``#VALUE!``. They do not stop early: ``AND(FALSE,#N/A)`` is ``#N/A``.
"""

from __future__ import annotations

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.registry import LAZY, R, V, function
from pyopenvba.formula._calc.values import NA, VALUE, Array, Empty, ExcelError, Reference, Scalar, Value, compare
from pyopenvba.formula._calc.cells import CellError


def _truth(context: Context, value: Scalar) -> bool | CellError:
    if isinstance(value, CellError):
        return value
    try:
        return context.logical(value)
    except ExcelError as error:
        return error.error


@function("IF", LAZY, LAZY, LAZY, minimum=2, legacy_as_cell=True)
def IF(context: Context, test: Node, then: Node, otherwise: Node | None = None) -> Value:
    condition = context.operand(context.evaluate(test))
    if isinstance(condition, Array):
        chosen = context.operand(context.evaluate(then))
        other = context.operand(context.evaluate(otherwise)) if otherwise is not None else False
        arrays = [item for item in (condition, chosen, other) if isinstance(item, Array)]
        rows: list[list[Scalar]] = []
        for row in range(max(array.height for array in arrays)):
            items: list[Scalar] = []
            for column in range(max(array.width for array in arrays)):
                truth = _truth(context, condition.at(row, column))
                if isinstance(truth, CellError):
                    items.append(truth)
                    continue
                source = chosen if truth else other
                items.append(source.at(row, column) if isinstance(source, Array) else source)
            rows.append(items)
        return Array(rows)
    truth = _truth(context, condition)
    if isinstance(truth, CellError):
        return truth
    if truth:
        return context.evaluate(then)
    if otherwise is None:
        return False
    return context.evaluate(otherwise)


def _logicals(context: Context, args: tuple[Value, ...]) -> list[bool]:
    found: list[bool] = []
    for arg in args:
        if isinstance(arg, (Reference, Array)):
            for value, _ in context.scalars(arg):
                if isinstance(value, CellError):
                    raise ExcelError(value)
                if isinstance(value, (bool, float)):
                    found.append(bool(value))
        elif isinstance(arg, CellError):
            raise ExcelError(arg)
        elif isinstance(arg, str) and arg.upper() not in ("TRUE", "FALSE"):
            # Text that says neither is passed over, as text in a range is: AND(TRUE,"x") and AND(TRUE,"1") are TRUE,
            # and AND("x") has nothing logical (pyOpenVBA's tests/fixtures/formula/).
            continue
        elif not isinstance(arg, Empty):
            found.append(context.logical(arg))
    if not found:
        raise ExcelError(VALUE)
    return found


@function("AND", R, maximum=255)
def AND(context: Context, *args: Value) -> Value:
    return all(_logicals(context, args))


@function("OR", R, maximum=255)
def OR(context: Context, *args: Value) -> Value:
    return any(_logicals(context, args))


@function("XOR", R, maximum=255)
def XOR(context: Context, *args: Value) -> Value:
    return sum(_logicals(context, args)) % 2 == 1


@function("NOT", V)
def NOT(context: Context, value: Scalar) -> Value:
    return not context.logical(value)


@function("TRUE")
def TRUE(context: Context) -> Value:
    return True


@function("FALSE")
def FALSE(context: Context) -> Value:
    return False


@function("IFERROR", V, V, legacy_first=(0, 1), legacy_as_cell=True)
def IFERROR(context: Context, value: Scalar, fallback: Scalar) -> Value:
    return fallback if isinstance(value, CellError) else value


@function("IFNA", V, V, legacy_first=(0, 1), legacy_as_cell=True)
def IFNA(context: Context, value: Scalar, fallback: Scalar) -> Value:
    return fallback if isinstance(value, CellError) and value.code == NA.code else value


@function("IFS", LAZY, LAZY, maximum=254, repeat=2, legacy_as_cell=True)
def IFS(context: Context, *args: Node) -> Value:
    if len(args) % 2:
        return NA
    for index in range(0, len(args), 2):
        truth = _truth(context, context.first(context.evaluate(args[index])))
        if isinstance(truth, CellError):
            return truth
        if truth:
            return context.evaluate(args[index + 1])
    return NA


@function("SWITCH", V, LAZY, LAZY, maximum=254, repeat=2, legacy_as_cell=True)
def SWITCH(context: Context, expression: Scalar, *args: Node) -> Value:
    if isinstance(expression, CellError):
        return expression
    pairs = len(args) // 2
    for index in range(pairs):
        candidate = context.first(context.evaluate(args[2 * index]))
        if isinstance(candidate, CellError):
            return candidate
        if compare(expression, candidate) == 0:
            return context.evaluate(args[2 * index + 1])
    if len(args) % 2:
        return context.evaluate(args[-1])
    return NA


@function("CHOOSE", R, LAZY, maximum=255, legacy_as_cell=True)
def CHOOSE(context: Context, index: Value, *choices: Node) -> Value:
    chosen = context.operand(index)
    if not isinstance(chosen, Array):
        position = int(context.number(chosen))
        if not 1 <= position <= len(choices):
            return VALUE
        return context.evaluate(choices[position - 1])
    # An array of positions picks item by item, each from the choice it names, lined up with the positions as an
    # operator lines two arrays up: CHOOSE({1,2},A1:A3,B1:B3) is the two columns side by side
    # (tests/fixtures/formula/).
    positions: list[list[int | CellError]] = [[_position(context, item, len(choices)) for item in row]
                                              for row in chosen.rows]
    picked = {one: context.array_of(context.evaluate(choices[one - 1]))
              for row in positions for one in row if isinstance(one, int)}
    height = max([chosen.height, *(array.height for array in picked.values())])
    width = max([chosen.width, *(array.width for array in picked.values())])
    rows: list[list[Scalar]] = []
    for row in range(height):
        items: list[Scalar] = []
        for column in range(width):
            where = chosen.at(row, column)
            one = _position(context, where, len(choices))
            items.append(one if isinstance(one, CellError) else picked[one].at(row, column))
        rows.append(items)
    return Array(rows)


def _position(context: Context, item: Scalar, count: int) -> int | CellError:
    """Which choice an item of CHOOSE's positions names, or the error it makes."""
    if isinstance(item, CellError):
        return item
    try:
        position = int(context.number(item))
    except ExcelError as error:
        return error.error
    return position if 1 <= position <= count else VALUE


__all__: list[str] = []
