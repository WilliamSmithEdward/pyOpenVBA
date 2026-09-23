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
from dataclasses import dataclass, replace
from typing import Final, Protocol

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area, parse_area
from pyopenvba.formula import _parse as P
from pyopenvba.formula._structured import TableShape, area as structured_area
from pyopenvba.formula._values import (
    BLANK,
    DIV0,
    NAME,
    NULL,
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

    def table(self, name: str) -> TableShape | None:
        """The table a structured reference names, found in any case, or None."""
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
    #: Worked out as an array formula is, by Evaluate or for a defined name's formula: blocks stay whole.
    array: bool = False
    #: Inside an argument Excel works out as an array even in a cell, SUMPRODUCT's or INDEX's first: blocks stay
    #: whole as in an array formula, but IF, CHOOSE, IFERROR, IFNA, IFS and SWITCH work their own arguments out as a
    #: cell does.
    array_argument: bool = False

    def at(self, sheet: str, row: int, column: int) -> Context:
        return Context(self.grid, sheet, row, column, self.depth + 1)

    def value(self, node: P.Node | None) -> object:
        """What one node comes to."""
        if node is None:
            return BLANK
        return evaluate(node, self)

    def area_of(self, node: P.Node | None) -> Area | None:
        """The block a node names, where it names one block."""
        areas = self.areas_of(node)
        return areas[0] if areas is not None and len(areas) == 1 else None

    def areas_of(self, node: P.Node | None) -> list[Area] | None:
        """The blocks a node names, or None when it comes to a value.

        A reference, a name for cells, the range operator, an intersection
        or a union, and INDEX, OFFSET, INDIRECT, CHOOSE or IF landing on
        cells all name blocks (tests/fixtures/reference_forms.json).
        """
        if isinstance(node, P.Reference):
            return [self.resolve(node)]
        if isinstance(node, P.Structured):
            return [self.structured(node)]
        if isinstance(node, P.NameNode):
            found = self.grid.named(node.name, node.sheet or self.sheet)
            if isinstance(found, Area):
                return [found]
            return self.areas_of(found) if isinstance(found, P.Node) else None
        if isinstance(node, P.Binary) and node.op in P.REFERENCE_OPS:
            left, right = self.areas_of(node.left), self.areas_of(node.right)
            if left is None or right is None:
                raise VALUE
            return _joined(node.op, left, right)
        if isinstance(node, P.Call):
            from pyopenvba.formula._functions import REFERENCES

            found_references = REFERENCES.get(node.name.upper())
            return found_references(self, node.args) if found_references is not None else None
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

    def structured(self, node: P.Structured) -> Area:
        """The cells a structured reference names; a table the workbook has not got is #REF!."""
        table = self.grid.table(node.table) if node.table else None
        if table is None:
            raise REF
        return structured_area(node, table, self.row)


def evaluate_formula(node: P.Node, context: Context) -> object:
    """A whole formula's value, as a cell or a defined name holds it.

    A formula that ends on a + or - whose answer all but cancels is 0, as
    Excel sets it (snapped); one in brackets, or inside a function, keeps
    its bits (tests/fixtures/zero_snap.json).
    """
    if isinstance(node, P.Binary) and node.op in ("+", "-") and not node.grouped:
        return _binary(node, context, snap=True)
    return evaluate(node, context)


def cell_answer(node: P.Node, context: Context) -> object:
    """What a cell holding this formula shows: one value, cells cut to the cell's own row or column as @A1:A3 is."""
    return single(intersected(node, context) if _names_cells(node) else evaluate_formula(node, context))


def evaluate(node: P.Node, context: Context) -> object:
    """A parsed formula's value, errors included."""
    if context.depth > MAX_DEPTH:
        raise ExcelError("#CIRCULAR!")
    if isinstance(node, P.Literal):
        return BLANK if node.value is None else node.value
    if isinstance(node, P.Reference):
        return context.grid.block(context.sheet, context.resolve(node))
    if isinstance(node, P.Structured):
        return context.grid.block(context.sheet, context.structured(node))
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
        # A name standing for a formula is worked out where it is used, and as an array formula is: SUM(dbl) adds
        # all of $A$1:$A$3*2, and =dbl on its own shows the first (tests/fixtures/formula/probes.txt).
        return evaluate_formula(found, replace(context, array=True))
    return found


def _unary(node: P.Unary, context: Context) -> object:
    # A minus sign wants one value too: =ABS(-A1:A3) in H2 reads A2.
    value = intersected(node.operand, context) if node.operand is not None else BLANK
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


def _joined(op: str, left: list[Area], right: list[Area]) -> list[Area]:
    """Two references joined by a reference operator: the block spanning both, the cells they share, or both."""
    if op == ",":
        return left + right
    if op == ":":
        if len(left) != 1 or len(right) != 1 or left[0].sheet.lower() != right[0].sheet.lower():
            raise VALUE
        first, second = left[0], right[0]
        return [Area(min(first.top, second.top), min(first.left, second.left), max(first.bottom, second.bottom),
                     max(first.right, second.right), first.sheet)]
    shared: list[Area] = []
    for first in left:
        for second in right:
            top, bottom = max(first.top, second.top), min(first.bottom, second.bottom)
            start, end = max(first.left, second.left), min(first.right, second.right)
            if first.sheet.lower() == second.sheet.lower() and top <= bottom and start <= end:
                shared.append(Area(top, start, bottom, end, first.sheet))
    if not shared:
        # Blocks that do not meet, A1:A2 B3:B4.
        raise NULL
    return shared


def _referenced(areas: list[Area], context: Context) -> object:
    """What the blocks a reference operator names hold: one block, or several at once."""
    blocks = [context.grid.block(context.sheet, area) for area in areas]
    return blocks[0] if len(blocks) == 1 else Areas(tuple(blocks))


def _binary(node: P.Binary, context: Context, *, snap: bool = False) -> object:
    if node.op in P.REFERENCE_OPS:
        areas = context.areas_of(node)
        assert areas is not None
        return _referenced(areas, context)
    left = _operand(node.left, context)
    right = _operand(node.right, context)
    if isinstance(left, Matrix) or isinstance(right, Matrix):
        return _over_blocks(node.op, left, right, snap=snap)
    return _apply(node.op, left, right, snap=snap)


def _operand(node: P.Node | None, context: Context) -> object:
    """One side of an operator: one value, as intersected gives it."""
    return BLANK if node is None else intersected(node, context)


def intersected(node: P.Node, context: Context) -> object:
    """A node where a formula wants one value, narrowed the way a cell narrows it.

    Cells spanning several -- a reference, a name for cells, OFFSET or
    INDEX landing on cells -- are cut down to the one on the formula's own
    row or column, which is what Excel writes as @A1:A3: =LEN(A1:A3) in H2
    is LEN(A2), and in H5, beside none of them, #VALUE!. An array written
    out or answered by a function is left whole, and so are the cells when
    the formula, or the argument, is worked out as an array
    (tests/fixtures/formula/probes.txt).
    """
    if not _names_cells(node):
        return evaluate(node, context)
    areas = context.areas_of(node)
    if areas is None:
        return evaluate(node, context)
    if context.array or context.array_argument:
        return _referenced(areas, context)
    if len(areas) != 1:
        # Several areas where one value is wanted.
        return VALUE
    return _at_formula(areas[0], context)


def _names_cells(node: P.Node) -> bool:
    """Whether a node can come to cells rather than a value, which only areas_of can settle."""
    if isinstance(node, (P.Reference, P.Structured, P.NameNode)):
        return True
    if isinstance(node, P.Binary):
        return node.op in P.REFERENCE_OPS
    if isinstance(node, P.Call):
        from pyopenvba.formula._functions import REFERENCES

        return node.name.upper() in REFERENCES
    return False


def _at_formula(area: Area, context: Context) -> object:
    """The cell of ``area`` on the formula's own row, column or both, or #VALUE! when it has none there."""
    row, column = area.top, area.left
    if area.rows > 1 and area.columns == 1:
        row = context.row
    elif area.columns > 1 and area.rows == 1:
        column = context.column
    elif area.rows > 1:
        row, column = context.row, context.column
    if not (area.top <= row <= area.bottom and area.left <= column <= area.right):
        return VALUE
    return context.grid.cell_value(area.sheet or context.sheet, row, column)


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
