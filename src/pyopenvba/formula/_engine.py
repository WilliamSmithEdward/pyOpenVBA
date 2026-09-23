"""Evaluating a parsed formula against a grid.

The evaluator is pure: it asks a :class:`Grid` for what is in a cell and
never writes anything.  Which cells need recalculating, and in what
order, is the grid's business, and
:mod:`pyopenvba.apps.excel._calc` is the one that answers for a
workbook.

Errors travel as values.  A cell that cannot be computed holds
``#VALUE!`` rather than raising, so IFERROR can catch it and a
neighbouring cell can show it.
"""

from __future__ import annotations

import datetime as _dt
import math
from collections.abc import Callable
from dataclasses import dataclass
from typing import Final, Protocol

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_area
from pyopenvba.formula import _parse as P
from pyopenvba.formula._values import (
    BLANK,
    DIV0,
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
    snapped,
)

#: How deep a chain of formulas may go before it is called circular.
MAX_DEPTH: Final = 128


class Grid(Protocol):
    """What the evaluator needs from whatever holds the cells."""

    def cell_value(self, sheet: str, row: int, column: int) -> object:
        """One cell's value, already calculated, or BLANK if it is empty."""
        ...

    def block(self, sheet: str, area: Area) -> Matrix:
        """A whole area's values, clipped to what the sheet actually uses."""
        ...

    def named(self, name: str, sheet: str) -> object:
        """A defined name's value, area or formula (a parsed node), or None if there is no such name."""
        ...

    def sheet_exists(self, name: str) -> bool: ...

    def formula_at(self, sheet: str, row: int, column: int) -> str:
        """The formula text in one cell, for FORMULATEXT."""
        ...

    def used(self, sheet: str) -> tuple[int, int, int, int] | None:
        """A sheet's used block -- top, left, bottom, right -- or None when it has nothing."""
        ...

    def hidden_rows(self, sheet: str, every: bool) -> set[int]:
        """The rows SUBTOTAL passes over on a sheet: every hidden one, or with ``every`` false those a filter hid."""
        ...

    def now(self) -> _dt.datetime:
        """The clock the workbook is running on."""
        ...


@dataclass(slots=True)
class Context:
    """Where a formula is being evaluated from."""

    grid: Grid
    sheet: str
    row: int = 1
    column: int = 1
    depth: int = 0
    #: Worked out as an array formula is, as Evaluate works one: a block beside an operator stays whole.
    array: bool = False

    def at(self, sheet: str, row: int, column: int) -> Context:
        return Context(self.grid, sheet, row, column, self.depth + 1)

    def value(self, node: P.Node | None) -> object:
        """What one node comes to."""
        if node is None:
            return BLANK
        return evaluate(node, self)

    def area_of(self, node: P.Node | None) -> Area | None:
        """The block a node names, where it names one."""
        if isinstance(node, P.Reference):
            return self.resolve(node)
        if isinstance(node, P.NameNode):
            found = self.grid.named(node.name, node.sheet or self.sheet)
            return found if isinstance(found, Area) else None
        return None

    def resolve(self, node: P.Reference) -> Area:
        """A reference as an area on a named sheet."""
        sheet = node.sheet or self.sheet
        if node.sheet and not self.grid.sheet_exists(node.sheet):
            raise REF
        try:
            area = parse_area(node.text, sheet=sheet)
        except ValueError:
            if ":" not in node.text:
                try:
                    area = parse_area(f"{node.text}:{node.text}", sheet=sheet)
                except ValueError:
                    raise REF from None
            else:
                raise REF from None
        return Area(area.top, area.left, area.bottom, area.right, sheet)


def evaluate_formula(node: P.Node, context: Context) -> object:
    """A whole formula's value, as a cell or a defined name holds it.

    A formula that ends on a + or - whose answer all but cancels is 0, as
    Excel sets it (snapped); one in brackets, or inside a function, keeps
    its bits (tests/fixtures/zero_snap.json).
    """
    if isinstance(node, P.Binary) and node.op in ("+", "-") and not node.grouped:
        return _binary(node, context, snap=True)
    return evaluate(node, context)


def evaluate(node: P.Node, context: Context) -> object:
    """A parsed formula's value, errors included."""
    if context.depth > MAX_DEPTH:
        raise ExcelError("#CIRCULAR!")
    if isinstance(node, P.Literal):
        return BLANK if node.value is None else node.value
    if isinstance(node, P.Reference):
        return context.grid.block(context.sheet, context.resolve(node))
    if isinstance(node, P.NameNode):
        return _named(node, context)
    if isinstance(node, P.Call):
        from pyopenvba.formula._functions import call

        return call(node.name, node.args, context)
    if isinstance(node, P.Unary):
        return _unary(node, context)
    if isinstance(node, P.Binary):
        return _binary(node, context)
    if isinstance(node, P.ArrayLiteral):
        return Matrix([[single(evaluate(item, context)) for item in row] for row in node.rows])
    raise VALUE


def _named(node: P.NameNode, context: Context) -> object:
    found = context.grid.named(node.name, node.sheet or context.sheet)
    if found is None:
        raise NAME
    if isinstance(found, Area):
        return context.grid.block(context.sheet, found)
    if isinstance(found, P.Node):
        # A name standing for a formula is worked out where it is used.
        return evaluate_formula(found, context)
    return found


def _unary(node: P.Unary, context: Context) -> object:
    value = evaluate(node.operand, context) if node.operand is not None else BLANK
    if isinstance(value, Matrix) and value.single:
        # One cell is one value, as beside a binary operator: -A1 is a number, not a block of one.
        value = single(value)
    if node.op == "%":
        return _map(value, lambda one: as_number(one) / 100.0)
    return _map(value, lambda one: -as_number(one))


def _map(value: object, apply: Callable[[object], object]) -> object:
    """Apply an operation to a value, or to every element of a block."""
    if isinstance(value, Matrix):
        return Matrix([[_safely(apply, one) for one in row] for row in value.rows])
    return apply(value)


def _safely(apply: Callable[[object], object], value: object) -> object:
    try:
        return apply(value)
    except ExcelError as failure:
        return failure


def _binary(node: P.Binary, context: Context, *, snap: bool = False) -> object:
    if node.op in P.REFERENCE_OPS:
        from pyopenvba.exceptions import VBAUnsupportedError

        raise VBAUnsupportedError("a range between references, an intersection or a union of references in a "
                                  "formula is not implemented")
    left = _operand(node.left, context)
    right = _operand(node.right, context)
    if isinstance(left, Matrix) or isinstance(right, Matrix):
        return _over_blocks(node.op, left, right, snap=snap)
    return _apply(node.op, left, right, snap=snap)


def _operand(node: P.Node | None, context: Context) -> object:
    """One side of an operator, narrowed the way a cell narrows it.

    A reference spanning several cells used beside an operator is cut
    down to the one on the formula's own row or column, which is what
    Excel writes as @A1:A2 and what makes SUM(A1:A2*2) two rather than
    six.  An array written out is left whole, and so is every block when
    the formula is worked out as an array.
    """
    if node is None:
        return BLANK
    value = evaluate(node, context)
    if context.array or not isinstance(node, P.Reference) or not isinstance(value, Matrix) or value.single:
        return value
    area = context.resolve(node)
    if area.columns == 1 and area.rows > 1:
        offset = context.row - area.top
        return value.rows[offset][0] if 0 <= offset < value.height else VALUE
    if area.rows == 1 and area.columns > 1:
        offset = context.column - area.left
        return value.rows[0][offset] if 0 <= offset < value.width else VALUE
    down = context.row - area.top
    across = context.column - area.left
    if 0 <= down < value.height and 0 <= across < value.width:
        return value.rows[down][across]
    return VALUE


def _over_blocks(op: str, left: object, right: object, *, snap: bool = False) -> object:
    """An operator between blocks, which Excel applies element by element."""
    height = max(_height(left), _height(right))
    width = max(_width(left), _width(right))
    if height == 1 and width == 1:
        return _apply(op, single(left), single(right), snap=snap)
    rows: list[list[object]] = []
    for row in range(height):
        line: list[object] = []
        for column in range(width):
            try:
                line.append(_apply(op, _pick(left, row, column), _pick(right, row, column), snap=snap))
            except ExcelError as failure:
                line.append(failure)
        rows.append(line)
    return Matrix(rows)


def _height(value: object) -> int:
    return value.height if isinstance(value, Matrix) else 1


def _width(value: object) -> int:
    return value.width if isinstance(value, Matrix) else 1


def _pick(value: object, row: int, column: int) -> object:
    return value.at(row, column) if isinstance(value, Matrix) else value


def _apply(op: str, left: object, right: object, *, snap: bool = False) -> object:
    if op in ("=", "<>", "<", ">", "<=", ">="):
        return compare(op, left, right)
    if isinstance(left, ExcelError):
        raise left
    if isinstance(right, ExcelError):
        raise right
    if op == "&":
        return as_text(left) + as_text(right)
    first = as_number(left)
    second = as_number(right)
    if op == "+":
        result = first + second
    elif op == "-":
        result = first - second
    elif op == "*":
        result = first * second
    elif op == "/":
        if second == 0:
            raise DIV0
        result = first / second
    elif op == "^":
        try:
            power = first**second
        except (OverflowError, ValueError, ZeroDivisionError):
            raise NUM from None
        if isinstance(power, complex):
            raise NUM
        result = float(power)
    else:
        raise VALUE
    if not math.isfinite(result):
        # Past the largest double, as 1E+300*1E+300 is: #NUM! (tests/fixtures/evaluate.json).
        raise NUM
    return snapped(first, result) if snap and op in ("+", "-") else result


def clip(area: Area, used: tuple[int, int, int, int] | None) -> Area:
    """An area cut down to the cells a sheet actually holds.

    A whole-column reference covers a million rows; SUM(A:A) means the
    ones that are there, and walking the rest would take a minute and
    change nothing.
    """
    if area.bottom < MAX_ROWS and area.right < MAX_COLUMNS:
        return area
    if used is None:
        return Area(area.top, area.left, min(area.bottom, area.top), min(area.right, area.left), area.sheet)
    _, _, bottom, right = used
    return Area(
        max(area.top, 1),
        max(area.left, 1),
        min(area.bottom, max(bottom, area.top)),
        min(area.right, max(right, area.left)),
        area.sheet,
    )


def truth(value: object) -> bool:
    return as_bool(single(value))
