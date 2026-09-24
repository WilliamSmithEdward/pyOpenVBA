"""LET, LAMBDA and the functions that call a LAMBDA over an array: MAP,
REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY, with ISOMITTED.

A file writes the names these bind with a prefix, ``_xlpm.x``, and they
are looked up before any defined name. LET binds its names in order, so a
later value can use an earlier name. A LAMBDA is a value that is a
function: called straight away, ``LAMBDA(x,x*2)(4)``, bound by LET and
called by name, or handed to MAP and the rest. Left in a cell uncalled it
is ``#CALC!``.

The functions that call a LAMBDA once per item want one value back each
time: a LAMBDA that gives an array of more than one item there makes that
item ``#CALC!``, Excel's error for an array nested in another.
"""

from __future__ import annotations

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.nodes import NameReference, Node, StructuredReference
from pyopenvba.formula._calc.registry import LAZY, A, V, function
from pyopenvba.formula._calc.values import (
    CALC,
    EMPTY,
    VALUE,
    Array,
    Empty,
    ExcelError,
    Lambda,
    Reference,
    Scalar,
    Scope,
    Value,
)
from pyopenvba.formula._calc.cells import CellError


#: How a file marks a LAMBDA's parameter a call may leave out, [y] on screen, and every other bound name.
_OPTIONAL = "_XLOP."
_BOUND = "_XLPM."


def _parameter(node: Node) -> tuple[str, bool]:
    """A name LET or LAMBDA binds, as it is looked up, and whether a call may leave it out: LAMBDA(x,[y],x), which
    a file writes LAMBDA(_xlpm.x,_xlop.y,_xlpm.x), the name being _xlpm.y where it is used (pyOpenVBA's
    tests/fixtures/bound_names.json)."""
    if isinstance(node, StructuredReference) and node.table is None and not node.items and node.first is not None \
            and node.first == node.last:
        return node.first.upper(), True
    if not isinstance(node, NameReference) or node.prefix is not None:
        raise ExcelError(VALUE)
    name = node.name.upper()
    if name.startswith(_OPTIONAL):
        return _BOUND + name[len(_OPTIONAL) :], True
    return name, False


@function("LET", LAZY, minimum=3, maximum=253)
def LET(context: Context, *nodes: Node) -> Value:
    if len(nodes) % 2 == 0:
        return VALUE
    scope = Scope({})
    context.scopes.append(scope)
    try:
        for position in range(0, len(nodes) - 1, 2):
            name, optional = _parameter(nodes[position])
            if optional:
                return VALUE
            scope.values[name] = context.evaluate(nodes[position + 1])
        return context.evaluate(nodes[-1])
    finally:
        context.scopes.pop()


@function("LAMBDA", LAZY, minimum=1, maximum=254)
def LAMBDA(context: Context, *nodes: Node) -> Value:
    bound = [_parameter(node) for node in nodes[:-1]]
    parameters = tuple(name for name, _ in bound)
    if len(set(parameters)) != len(parameters):
        return VALUE
    optional = frozenset(name for name, left_out in bound if left_out)
    return Lambda.make(parameters, nodes[-1], tuple(context.scopes), optional)


@function("ISOMITTED", LAZY)
def ISOMITTED(context: Context, node: Node) -> Value:
    """Whether a LAMBDA's parameter was left out of the call."""
    if not isinstance(node, NameReference) or node.prefix is not None:
        return False
    key = node.name.upper()
    for scope in reversed(context.scopes):
        if key in scope.values:
            return key in scope.omitted
    return False


def _function(value: Value) -> Lambda:
    if isinstance(value, Lambda):
        return value
    if isinstance(value, CellError):
        raise ExcelError(value)
    raise ExcelError(VALUE)


def _single(context: Context, value: Value) -> Scalar:
    """One item a LAMBDA gave where one is wanted."""
    if isinstance(value, Reference):
        area = value.area
        if area is None or not area.is_cell:
            return CALC
        return context.book.cell(area.sheet, area.top, area.left)
    if isinstance(value, Array):
        return value.rows[0][0] if value.height == 1 and value.width == 1 else CALC
    if isinstance(value, Lambda):
        return CALC
    return value


@function("MAP", A, maximum=254)
def MAP(context: Context, *args: Value) -> Value:
    if len(args) < 2:
        return VALUE
    call = _function(args[-1])
    grids = [matrix(context, arg) for arg in args[:-1]]
    height = max(grid.height for grid in grids)
    width = max(grid.width for grid in grids)
    return Array(
        [
            [_single(context, context.apply(call, [grid.at(row, column) for grid in grids])) for column in range(width)]
            for row in range(height)
        ]
    )


def _start(value: Value) -> Value:
    return EMPTY if isinstance(value, Empty) else value


@function("REDUCE", A, A, A)
def REDUCE(context: Context, initial: Value, array: Value, step: Value) -> Value:
    call = _function(step)
    total = _start(initial)
    for item in matrix(context, array).items():
        total = context.apply(call, [total, item])
    return total


@function("SCAN", A, A, A)
def SCAN(context: Context, initial: Value, array: Value, step: Value) -> Value:
    call = _function(step)
    grid = matrix(context, array)
    total = _start(initial)
    rows: list[list[Scalar]] = []
    for row in grid.rows:
        line: list[Scalar] = []
        for item in row:
            total = context.apply(call, [total, item])
            line.append(_single(context, total))
        rows.append(line)
    return Array(rows)


@function("BYROW", A, A, minimum=1)
def BYROW(context: Context, array: Value, step: Value = CALC) -> Value:
    # Given no function, BYROW and BYCOL are #CALC! (tests/fixtures/formula_refusals.json).
    call = _function(step)
    grid = matrix(context, array)
    return Array([[_single(context, context.apply(call, [Array([list(row)])]))] for row in grid.rows])


@function("BYCOL", A, A, minimum=1)
def BYCOL(context: Context, array: Value, step: Value = CALC) -> Value:
    call = _function(step)
    grid = matrix(context, array)
    columns = [Array([[row[column]] for row in grid.rows]) for column in range(grid.width)]
    return Array([[_single(context, context.apply(call, [column])) for column in columns]])


@function("MAKEARRAY", V, V, A)
def MAKEARRAY(context: Context, rows: Scalar, columns: Scalar, step: Value) -> Value:
    call = _function(step)
    height = context.integer(rows)
    width = context.integer(columns)
    if height < 1 or width < 1:
        return VALUE
    return Array(
        [
            [_single(context, context.apply(call, [float(row), float(column)])) for column in range(1, width + 1)]
            for row in range(1, height + 1)
        ]
    )


__all__: list[str] = []
