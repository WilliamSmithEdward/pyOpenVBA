"""Evaluating a formula's tree against a workbook.

A :class:`Context` is one formula being evaluated: the workbook it reads,
the cell it is in, and whether it is an array formula. The cell matters to
more than ROW(): a range where one value is wanted, ``=A1:A10*2`` in row 5,
means the one cell of it in the formula's row, A5. That is the *implicit
intersection*, and it is what an ordinary formula does. An array formula
instead works on every cell of the range and gives an array.

The context is also what a function implementation receives, for reading
values and turning them into the numbers, text and logicals it needs.

Reading a cell whose formula has not been calculated yet raises
:class:`PendingCellsError`, naming every such cell a range held, so the
calculation can do those first and try this formula again. That keeps a
chain of ten thousand formulas off Python's call stack.
"""

from __future__ import annotations

import datetime as dt
import math
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from fractions import Fraction
from typing import Protocol

from pyopenvba.formula._calc import registry
from pyopenvba.formula._calc.catalog import is_excel_function
from pyopenvba.formula._calc.nodes import (
    ALL,
    DATA,
    HEADERS,
    THIS_ROW,
    TOTALS,
    AreaReference,
    ArrayLiteral,
    AxisReference,
    Binary,
    Call,
    CellReference,
    ErrorLiteral,
    ErrorReference,
    Invoke,
    Logical,
    Missing,
    NameReference,
    Node,
    Number,
    Paren,
    Postfix,
    Prefix,
    StructuredReference,
    Text,
    Unary,
)
from pyopenvba.formula._calc.numbers import near_zero, normal
from pyopenvba.formula._calc.precise import add, divide, exp, extended, ln, multiply, square_root, subtract
from pyopenvba.formula._calc.values import (
    DIV0,
    EMPTY,
    MAX_TEXT,
    NA,
    NAME,
    NULL,
    NUM,
    REF,
    VALUE,
    Area,
    Array,
    Empty,
    ExcelError,
    Lambda,
    Reference,
    Scalar,
    Scope,
    Value,
    compare,
    scalar_text,
    text_to_number,
)
from pyopenvba.formula._calc.reference import MAX_COLUMN, MAX_ROW
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.cells import UnsupportedFormulaError

CellKey = tuple[str, int, int]


class PendingCellsError(Exception):
    """A formula read cells whose formulas are not calculated yet."""

    def __init__(self, cells: list[CellKey]) -> None:
        super().__init__(f"{len(cells)} cells to calculate first")
        self.cells = cells


@dataclass(frozen=True)
class TableShape:
    """What a structured reference needs to know of a table."""

    name: str
    sheet: str
    top: int
    left: int
    bottom: int
    right: int
    headers: int
    totals: int
    columns: tuple[str, ...]


class Book(Protocol):
    """The workbook as the evaluator reads it."""

    @property
    def epoch_1904(self) -> bool: ...

    @property
    def name(self) -> str:
        """The workbook's file name, as CELL shows it: ``Book1.xlsx``."""
        ...

    @property
    def folder(self) -> str | None:
        """The folder the workbook was saved in, with its separator at the
        end, or ``None`` for a workbook never saved."""
        ...

    def sheet_key(self, name: str) -> str | None:
        """The sheet's name as the workbook spells it, for a name in any
        case, or ``None`` when there is no such sheet."""
        ...

    def sheet_order(self) -> list[str]: ...

    def cell(self, sheet: str, row: int, column: int) -> Scalar: ...

    def cells(self, area: Area) -> Iterator[tuple[int, int, Scalar]]:
        """The cells of an area that hold something, row by row."""
        ...

    def defined_name(self, name: str, sheet: str | None) -> Node | None:
        """A name's formula, scoped to ``sheet`` or, with ``None``, to the
        workbook."""
        ...

    def table(self, name: str) -> TableShape | None: ...

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None: ...

    def formula_text(self, sheet: str, row: int, column: int) -> str | None: ...

    def used(self, sheet: str) -> tuple[int, int]:
        """The last row and column of a sheet that hold anything, or
        (0, 0) for an empty sheet."""
        ...

    def row_hidden(self, sheet: str, row: int) -> bool:
        """Whether a row is hidden, by hand or by a filter."""
        ...

    def row_filtered(self, sheet: str, row: int) -> bool:
        """Whether a row is hidden because a filter leaves it out."""
        ...

    def subtotal(self, sheet: str, row: int, column: int) -> bool:
        """Whether a cell's formula calls SUBTOTAL or AGGREGATE, which the
        two leave out of what they total."""
        ...

    def spill(self, sheet: str, row: int, column: int) -> Area | None:
        """The block a dynamic-array formula in a cell spilled into, as the
        file records it, or ``None`` when the cell holds no such formula."""
        ...


_ARITHMETIC = frozenset({"+", "-", "*", "/", "^"})
_COMPARISON: dict[str, Callable[[int], bool]] = {
    "=": lambda order: order == 0,
    "<>": lambda order: order != 0,
    "<": lambda order: order < 0,
    ">": lambda order: order > 0,
    "<=": lambda order: order <= 0,
    ">=": lambda order: order >= 0,
}
_GETTING_DATA = "#GETTING_DATA"
#: How deep LAMBDAs may call one another here: each level costs Python a
#: dozen frames, so a deeper recursion keeps the value Excel cached.
_MAX_DEPTH = 60


class Context:
    """One formula being evaluated, and the conversions its functions use."""

    def __init__(
        self,
        book: Book,
        sheet: str,
        row: int,
        column: int,
        *,
        array: bool = False,
        today: dt.date | None = None,
        now: dt.datetime | None = None,
    ) -> None:
        self.book = book
        self.sheet = sheet
        self.row = row
        self.column = column
        self.array = array
        #: A legacy formula, which Excel works out as it did before dynamic arrays (see registry). Every formula but
        #: an array formula is one: Range.Formula writes them, and a file keeps a formula that needs dynamic arrays
        #: as an array formula.
        self.legacy = not array
        self.now = now or dt.datetime.now()
        self.today = today or self.now.date()
        #: Names bound by LET and by a LAMBDA's call, innermost last.
        self.scopes: list[Scope] = []
        self._names: list[str] = []
        self._depth = 0

    @property
    def epoch_1904(self) -> bool:
        return self.book.epoch_1904

    # ------------------------------------------------------------------
    # Evaluation
    # ------------------------------------------------------------------

    def formula(self, node: Node) -> Value:
        """A whole formula's value, before it goes into a cell. Only here
        does a last ``+`` or ``-`` treat a rounding error as zero."""
        if isinstance(node, Binary) and node.op in ("+", "-"):
            return self.operate(node.op, self.evaluate(node.left), self.evaluate(node.right), final=True)
        return self.evaluate(node)

    def evaluate(self, node: Node) -> Value:
        # Most frequent first.
        if isinstance(node, CellReference):
            return self._cell(node)
        if isinstance(node, Call):
            return self._call(node)
        if isinstance(node, Binary):
            return self._binary(node)
        if isinstance(node, (Number, Text, Logical)):
            return node.value
        if isinstance(node, AreaReference):
            return self._area(node)
        if isinstance(node, NameReference):
            return self._name(node)
        if isinstance(node, Unary):
            return self._unary(node)
        if isinstance(node, Paren):
            return self.evaluate(node.inner)
        if isinstance(node, ErrorLiteral):
            # Measured: the literal #GETTING_DATA is #N/A once calculated.
            return NA if node.code == _GETTING_DATA else CellError(node.code)
        if isinstance(node, Postfix):
            return self._postfix(node)
        if isinstance(node, StructuredReference):
            return self._structured(node)
        if isinstance(node, AxisReference):
            return self._axis(node)
        if isinstance(node, ArrayLiteral):
            return Array([list(row) for row in node.rows])
        if isinstance(node, ErrorReference):
            return REF
        if isinstance(node, Invoke):
            return self._invoke(self.evaluate(node.target), node.args)
        return EMPTY

    def evaluate_array(self, node: Node) -> Value:
        """node evaluated as an array formula would evaluate it."""
        if self.array:
            return self.evaluate(node)
        self.array = True
        try:
            return self.evaluate(node)
        finally:
            self.array = False

    def _cell(self, node: CellReference) -> Value:
        sheets = self._sheets(node.prefix)
        if isinstance(sheets, CellError):
            return sheets
        ref = node.ref
        return Reference(tuple(Area(sheet, ref.row, ref.column, ref.row, ref.column) for sheet in sheets))

    def _area(self, node: AreaReference) -> Value:
        sheets = self._sheets(node.prefix)
        if isinstance(sheets, CellError):
            return sheets
        top, bottom = sorted((node.first.row, node.last.row))
        left, right = sorted((node.first.column, node.last.column))
        return Reference(tuple(Area(sheet, top, left, bottom, right) for sheet in sheets))

    def _axis(self, node: AxisReference) -> Value:
        sheets = self._sheets(node.prefix)
        if isinstance(sheets, CellError):
            return sheets
        axis = node.ref
        if axis.is_row:
            return Reference(tuple(Area(sheet, axis.low, 1, axis.high, MAX_COLUMN) for sheet in sheets))
        return Reference(tuple(Area(sheet, 1, axis.low, MAX_ROW, axis.high) for sheet in sheets))

    def _sheets(self, prefix: Prefix | None) -> list[str] | CellError:
        """The sheets a reference's prefix names, in the workbook's order."""
        if prefix is None:
            return [self.sheet]
        if prefix.book is not None:
            raise UnsupportedFormulaError("a reference to another workbook")
        if prefix.sheet is None:
            return REF
        first = self.book.sheet_key(prefix.sheet)
        if first is None:
            return REF
        if prefix.last_sheet is None:
            return [first]
        last = self.book.sheet_key(prefix.last_sheet)
        if last is None:
            return REF
        order = self.book.sheet_order()
        start, end = sorted((order.index(first), order.index(last)))
        return order[start : end + 1]

    def _name(self, node: NameReference) -> Value:
        key = node.name.upper()
        if node.prefix is None:
            for scope in reversed(self.scopes):
                if key in scope.values:
                    return scope.values[key]
        if node.prefix is not None and node.prefix.book is not None:
            raise UnsupportedFormulaError("a name defined in another workbook")
        sheet = None
        if node.prefix is not None:
            if node.prefix.sheet is None:
                return NAME
            sheet = self.book.sheet_key(node.prefix.sheet)
            if sheet is None:
                return REF
            found = self.book.defined_name(node.name, sheet)
        else:
            found = self.book.defined_name(node.name, self.sheet)
            if found is None:
                found = self.book.defined_name(node.name, None)
        if found is None:
            # A table's name alone is not a reference in a file, where Excel
            # writes Stock[] for it: measured, VLOOKUP(x,Stock,2) is #NAME?.
            return NAME
        if key in self._names or len(self._names) > 64:
            return NAME
        self._names.append(key)
        saved = self.array, self.legacy
        # A name standing for a formula is worked out where it is used, and as an array formula is, whatever formula
        # uses it: SUM(dbl) with dbl =$A$1:$A$3*2 adds all three, =dbl alone shows the first, and an IF in the name
        # picks item by item (tests/fixtures/formula/). As a formula of its own, its last sum snaps to zero as a
        # cell's does (tests/fixtures/zero_snap.json).
        self.array, self.legacy = True, False
        try:
            return self.formula(found)
        finally:
            self.array, self.legacy = saved
            self._names.pop()

    def _structured(self, node: StructuredReference) -> Value:
        table = (
            self.book.table_at(self.sheet, self.row, self.column)
            if node.table is None
            else self.book.table(node.table)
        )
        if table is None:
            return REF if node.table is None else NAME
        data_top = table.top + table.headers
        data_bottom = table.bottom - table.totals
        spans: list[tuple[int, int]] = []
        items = set(node.items) or {DATA}
        if ALL in items:
            spans.append((table.top, table.bottom))
        # A special item the table does not show leaves the rest: Table1[[#Data],[#Totals]] is the data of a table
        # without a totals row, and only with nothing left is it #REF! (tests/fixtures/structured_references/).
        if HEADERS in items and table.headers:
            spans.append((table.top, table.top))
        if DATA in items:
            spans.append((data_top, data_bottom))
        if TOTALS in items and table.totals:
            spans.append((table.bottom, table.bottom))
        if THIS_ROW in items:
            # This row is the formula's own row, on whatever sheet the formula is.
            if not data_top <= self.row <= data_bottom:
                return VALUE
            spans.append((self.row, self.row))
        if not spans:
            return REF
        top = min(span[0] for span in spans)
        bottom = max(span[1] for span in spans)
        left, right = table.left, table.right
        if node.first is not None:
            names = [column.casefold() for column in table.columns]
            try:
                first = names.index(node.first.casefold())
                last = names.index((node.last or node.first).casefold())
            except ValueError:
                return REF
            first, last = sorted((first, last))
            left, right = table.left + first, table.left + last
        return Reference.of(Area(table.sheet, top, left, bottom, right))

    def _unary(self, node: Unary) -> Value:
        value = self.evaluate(node.operand)
        if node.op == "+":
            return value
        if node.op == "@":
            return self.implicit(value) if isinstance(value, Reference) else self.first(value)
        return self.map(value, self._negate)

    def _negate(self, value: Scalar) -> Scalar:
        if isinstance(value, CellError):
            return value
        number = self.coerce_number(value)
        if isinstance(number, CellError):
            return number
        return -number + 0.0

    def _postfix(self, node: Postfix) -> Value:
        value = self.evaluate(node.operand)
        if node.op == "#":
            return self.spilled(value)
        return self.map(value, self._percent)

    def spilled(self, value: Value) -> Value:
        """``A1#``: the block A1's dynamic-array formula spilled into, as
        the file records it. Anything else is ``#REF!``."""
        if isinstance(value, CellError):
            return value
        if not isinstance(value, Reference) or value.area is None or not value.area.is_cell:
            return REF
        area = value.area
        block = self.book.spill(area.sheet, area.top, area.left)
        return REF if block is None else Reference.of(block)

    def _percent(self, value: Scalar) -> Scalar:
        if isinstance(value, CellError):
            return value
        number = self.coerce_number(value)
        if isinstance(number, CellError):
            return number
        return normal(divide(number, 100.0))

    def _binary(self, node: Binary) -> Value:
        if node.op in (":", " ", ","):
            return self._reference_operator(node)
        return self.operate(node.op, self.evaluate(node.left), self.evaluate(node.right))

    def _reference_operator(self, node: Binary) -> Value:
        left = self.evaluate(node.left)
        right = self.evaluate(node.right)
        for side in (left, right):
            if isinstance(side, CellError):
                return side
        if not isinstance(left, Reference) or not isinstance(right, Reference):
            return VALUE
        if node.op == ",":
            return Reference(left.areas + right.areas)
        if node.op == " ":
            common = [
                found
                for first in left.areas
                for second in right.areas
                if (found := first.intersection(second)) is not None
            ]
            return Reference(tuple(common)) if common else NULL
        areas = left.areas + right.areas
        sheet = areas[0].sheet
        if any(area.sheet != sheet for area in areas):
            return VALUE
        return Reference.of(
            Area(
                sheet,
                min(area.top for area in areas),
                min(area.left for area in areas),
                max(area.bottom for area in areas),
                max(area.right for area in areas),
            )
        )

    def _call(self, node: Call) -> Value:
        entry = registry.FUNCTIONS.get(node.function)
        if entry is None:
            if is_excel_function(node.name):
                raise UnsupportedFormulaError(f"the function {node.function}")
            # A name LET bound to a LAMBDA, or a defined name holding one.
            found = self._name(NameReference(node.function))
            if isinstance(found, Lambda):
                return self._invoke(found, node.args)
            return NAME
        if not entry.minimum <= len(node.args) <= entry.maximum:
            return VALUE
        return registry.call(entry, self, node.args)

    def _invoke(self, target: Value, nodes: tuple[Node, ...]) -> Value:
        """A call of what an expression gave: a LAMBDA, with its arguments
        evaluated here."""
        if not isinstance(target, Lambda):
            return target if isinstance(target, CellError) else VALUE
        args = [EMPTY if isinstance(arg, Missing) else self.evaluate(arg) for arg in nodes]
        return self.apply(target, args)

    def apply(self, function: Lambda, args: list[Value]) -> Value:
        """A LAMBDA called with values: its parameters bound to them, in the
        scope it was made in. Parameters left out read as blank, and
        ISOMITTED says which they were; more arguments than parameters are
        ``#VALUE!``."""
        if len(args) > len(function.parameters) or function.body is None:
            return VALUE
        if self._depth >= _MAX_DEPTH:
            raise UnsupportedFormulaError(f"a LAMBDA calling itself more than {_MAX_DEPTH} deep")
        values = dict(zip(function.parameters, args, strict=False))
        omitted = frozenset(function.parameters[len(args) :])
        for name in omitted:
            values[name] = EMPTY
        saved = self.scopes
        self.scopes = [*function.closure, Scope(values, omitted)]
        self._depth += 1
        try:
            return self.evaluate(function.body)
        finally:
            self._depth -= 1
            self.scopes = saved

    # ------------------------------------------------------------------
    # Operators
    # ------------------------------------------------------------------

    def operate(self, op: str, left: Value, right: Value, *, final: bool = False) -> Value:
        """``left op right`` for an arithmetic, comparison or ``&``
        operator, item by item when either side is an array."""
        first = self.operand(left)
        second = self.operand(right)
        if isinstance(first, Array) or isinstance(second, Array):
            return self.broadcast(lambda a, b: self._scalar(op, a, b, final), first, second)
        return self._scalar(op, first, second, final)

    def _scalar(self, op: str, left: Scalar, right: Scalar, final: bool) -> Scalar:
        if op in _ARITHMETIC:
            if isinstance(left, CellError):
                return left
            a = self.coerce_number(left)
            if isinstance(a, CellError):
                return a
            if isinstance(right, CellError):
                return right
            b = self.coerce_number(right)
            if isinstance(b, CellError):
                return b
            return arithmetic(op, a, b, final)
        if isinstance(left, CellError):
            return left
        if isinstance(right, CellError):
            return right
        if op == "&":
            joined = scalar_text(left) + scalar_text(right)
            return VALUE if len(joined) > MAX_TEXT else joined
        return _COMPARISON[op](compare(left, right))

    def broadcast(self, operation: Callable[[Scalar, Scalar], Scalar], left: Scalar | Array, right: Scalar | Array) -> Array:
        """``operation`` item by item over two arrays, or an array and a
        scalar. A single row or column stretches to meet the other; past
        the end of a shorter side, the item is ``#N/A``."""
        first = left if isinstance(left, Array) else Array([[left]])
        second = right if isinstance(right, Array) else Array([[right]])
        height = max(first.height, second.height)
        width = max(first.width, second.width)
        return Array(
            [[operation(first.at(row, column), second.at(row, column)) for column in range(width)] for row in range(height)]
        )

    def map(self, value: Value, operation: Callable[[Scalar], Scalar]) -> Value:
        """``operation`` on a value, or on each item of an array."""
        operand = self.operand(value)
        if isinstance(operand, Array):
            return Array([[operation(item) for item in row] for row in operand.rows])
        return operation(operand)

    # ------------------------------------------------------------------
    # Ranges into values
    # ------------------------------------------------------------------

    def operand(self, value: Value) -> Scalar | Array:
        """A value where one value is wanted: a range is intersected with
        the formula's row or column, or in an array formula read whole."""
        if isinstance(value, Reference):
            if self.array:
                return self.grid(value)
            return self.implicit(value)
        return value

    def implicit(self, reference: Reference) -> Scalar:
        """The implicit intersection: the cell of the range in the
        formula's row, for a column, or column, for a row."""
        area = reference.area
        if area is None:
            return VALUE
        if area.is_cell:
            return self.book.cell(area.sheet, area.top, area.left)
        row, column = area.top, area.left
        if area.height > 1:
            if not area.top <= self.row <= area.bottom:
                return VALUE
            row = self.row
        if area.width > 1:
            if not area.left <= self.column <= area.right:
                return VALUE
            column = self.column
        return self.book.cell(area.sheet, row, column)

    def grid(self, reference: Reference) -> Scalar | Array:
        """Every value of a one-area range, as an array; a single cell as
        its value."""
        area = reference.area
        if area is None:
            return VALUE
        if area.is_cell:
            return self.book.cell(area.sheet, area.top, area.left)
        rows: list[list[Scalar]] = [[EMPTY] * area.width for _ in range(area.height)]
        for row, column, value in self.book.cells(area):
            rows[row - area.top][column - area.left] = value
        return Array(rows)

    def array_of(self, value: Value) -> Array:
        """Any value as an array: a range read whole, a scalar as 1x1."""
        if isinstance(value, Reference):
            found = self.grid(value)
            return found if isinstance(found, Array) else Array([[found]])
        if isinstance(value, Array):
            return value
        return Array([[value]])

    def scalars(self, value: Value) -> Iterator[tuple[Scalar, bool]]:
        """Every value a range, an array or a scalar holds, each with
        whether it came from a cell. A range yields only the cells that
        hold something."""
        if isinstance(value, Reference):
            for area in value.areas:
                for _, _, item in self.book.cells(area):
                    yield item, True
        elif isinstance(value, Array):
            for item in value.items():
                yield item, False
        else:
            yield value, False

    def first(self, value: Value) -> Scalar:
        """The one scalar a value comes down to where only one fits."""
        if isinstance(value, Reference):
            return self.implicit(value)
        if isinstance(value, Array):
            return value.rows[0][0] if value.rows and value.rows[0] else VALUE
        return value

    # ------------------------------------------------------------------
    # Coercion
    # ------------------------------------------------------------------

    def coerce_number(self, value: Scalar) -> float | CellError:
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, float):
            return value
        if isinstance(value, str):
            number = text_to_number(value, self.today, epoch_1904=self.epoch_1904)
            return VALUE if number is None else number
        if isinstance(value, Empty):
            return 0.0
        return value

    def number(self, value: Scalar) -> float:
        """A scalar as a number, raising the error it is or makes."""
        found = self.coerce_number(value)
        if isinstance(found, CellError):
            raise ExcelError(found)
        return found

    def text(self, value: Scalar) -> str:
        """A scalar as text, raising the error it is."""
        if isinstance(value, CellError):
            raise ExcelError(value)
        return scalar_text(value)

    def logical(self, value: Scalar) -> bool:
        """A scalar as TRUE or FALSE: a number is TRUE unless 0, and text
        must say TRUE or FALSE."""
        if isinstance(value, bool):
            return value
        if isinstance(value, CellError):
            raise ExcelError(value)
        if isinstance(value, Empty):
            return False
        if isinstance(value, str):
            upper = value.upper()
            if upper == "TRUE":
                return True
            if upper == "FALSE":
                return False
            raise ExcelError(VALUE)
        return value != 0.0

    def integer(self, value: Scalar) -> int:
        """A scalar as a whole number, cut toward zero."""
        number = self.number(value)
        if not math.isfinite(number) or abs(number) > 2.0**53:
            raise ExcelError(NUM)
        return int(number)


def arithmetic(op: str, a: float, b: float, final: bool = False) -> Scalar:
    """One arithmetic operation on two numbers, the Excel way: in an x87
    register, rounded to 64 bits and then to a double."""
    if op == "+":
        result = add(a, b)
        if final:
            result = near_zero(a, b, result)
    elif op == "-":
        result = subtract(a, b)
        if final:
            result = near_zero(a, -b, result)
    elif op == "*":
        result = multiply(a, b)
    elif op == "/":
        if b == 0.0:
            return DIV0
        result = divide(a, b)
    else:
        return power(a, b)
    if not math.isfinite(result):
        return NUM
    if op in ("+", "-"):
        # A sum or difference too small to be normal keeps its bits: 2^-1000 less the double after it is -2^-1052
        # (pyOpenVBA's tests/fixtures/zero_snap.json).
        return result + 0.0
    return normal(result) + 0.0


def power(base: float, exponent: float) -> Scalar:
    """``base ^ exponent``, computed as Excel computes it, to the last bit
    where the corpus could tell.

    A whole exponent is square-and-multiply, which is why ``1.1^100`` is
    ``13780.612339822364`` where the C library's pow gives ...238. A
    positive one is EXP of ``exponent * LN(base)``, each as the x87 gives
    it, so ``8^(1/3)`` is ``1.9999999999999998`` rather than 2, and 0.5 is
    a square root. A negative one is the reciprocal of the positive power,
    so ``8^(-1/3)`` is exactly 0.5. Measured, 370 of 370 in the corpus and
    999 of 1000 random pairs: the other lands within a 64-bit unit of a
    tie between two doubles, where the processor's F2XM1 decides. A
    negative base takes an odd root, ``(-8)^(1/3)`` being about -2, and an
    even one is ``#NUM!``.
    """
    if base == 0.0:
        if exponent == 0.0:
            return NUM
        return DIV0 if exponent < 0 else 0.0
    # Callers pass a count as an int, which a float annotation admits, and
    # int has no is_integer before Python 3.12.
    if float(exponent).is_integer():
        result = _whole_power(base, int(exponent))
    else:
        negative = False
        if base < 0.0:
            root = 1 / exponent
            whole = round(root)
            # The root, read to fifteen digits, must be an odd whole number.
            if abs(root - whole) > abs(root) * 1e-14 or whole % 2 == 0:
                return NUM
            base, negative = -base, True
        try:
            result = _positive_power(base, abs(exponent))
        except OverflowError:
            return NUM
        if exponent < 0:
            if result == 0.0:
                return NUM
            result = extended(1 / Fraction(result)) if math.isfinite(result) else 0.0
        if negative:
            result = -result
    if not math.isfinite(result):
        return NUM
    return normal(result) + 0.0


def _positive_power(base: float, exponent: float) -> float:
    if exponent == 0.5:
        return square_root(base)
    return exp(multiply(exponent, ln(base)))


def _whole_power(base: float, exponent: int) -> float:
    """``base`` to a whole power by repeated squaring, the reciprocal for a
    negative one."""
    result = 1.0
    square = base
    remaining = abs(exponent)
    while remaining:
        if remaining & 1:
            result = multiply(result, square)
        remaining >>= 1
        if remaining:
            square = multiply(square, square)
    if exponent < 0:
        return math.inf if result == 0.0 else divide(1.0, result)
    return result


__all__ = ["Book", "CellKey", "Context", "PendingCellsError", "TableShape", "UnsupportedFormulaError", "arithmetic", "power"]
