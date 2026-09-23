"""The model's workbook as the formula engine reads it: the Book protocol of :mod:`pyopenvba.formula._calc`.

The engine asks for cells, defined names, tables and hidden rows; the
calculator answers with what it has worked out, working a formula out
first when one the engine reads is stale. Values cross both ways: the
engine's CellError is the model's ExcelError, its EMPTY a blank cell, and
a Date its serial.

Two rules the model measured stay the model's own. SUBTOTAL's 1 to 11
pass over every hidden row while the sheet is in filter mode, whatever
hid it (scripts/measure_subtotal.py). And a defined name whose formula
has a relative reference counts it from the cell the name was defined at,
which the file does not record, so such a name reports itself.
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterator
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area as ModelArea
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc import registry
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.evaluator import TableShape
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import (
    AreaReference,
    AxisReference,
    Call,
    CellReference,
    NameReference,
    Node,
    Prefix,
    StructuredReference,
    walk,
)
from pyopenvba.formula._calc.parser import parse
from pyopenvba.formula._calc.values import EMPTY, Area, Empty, Scalar
from pyopenvba.formula._values import BLANK, ExcelError
from pyopenvba.interpreter._values import EMPTY as NOTHING_THERE, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._calc import Calculator
    from pyopenvba.apps.excel._model import Worksheet

#: The functions whose cells SUBTOTAL and AGGREGATE leave out.
_SUBTOTALS = frozenset({"SUBTOTAL", "AGGREGATE"})
#: Functions whose answer can change with no cell they name changing: a reference worked out as they run, the
#: clock, chance, or rows hidden. Excel works them out after every edit.
_VOLATILE = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO",
                       "SUBTOTAL"})
#: One side of a reference that names its cells outright: $A$1, a whole column $A or a whole row $1.
_ABSOLUTE = re.compile(r"\$[A-Za-z]{1,3}(?:\$[0-9]+)?|\$[0-9]+")
#: VBA's CVErr numbers for Excel's errors.
_ERROR_NAMES = {2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?", 2036: "#NUM!",
                2042: "#N/A"}


def scalar(value: object) -> Scalar:
    """A value the model holds, as the engine computes with it."""
    if value is None or value is BLANK or value is NOTHING_THERE:
        return EMPTY
    if isinstance(value, bool) or isinstance(value, str):
        return value
    if isinstance(value, ExcelError):
        return CellError(value.name)
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (int, float, VBAInt)):
        return float(value)
    if isinstance(value, VBAErrorValue):
        return CellError(_ERROR_NAMES.get(value.number, "#VALUE!"))
    return EMPTY


def model_value(value: Scalar) -> object:
    """What the engine worked out, as a cell of the model holds it: a formula that comes to nothing shows 0."""
    if isinstance(value, Empty):
        return 0.0
    if isinstance(value, CellError):
        return ExcelError(value.code)
    return value


class EngineBook:
    """The Book protocol over a workbook's calculator."""

    def __init__(self, calculator: Calculator) -> None:
        self.calculator = calculator
        self.book = calculator.book

    @property
    def epoch_1904(self) -> bool:
        return False

    @property
    def name(self) -> str:
        return self.book.name

    @property
    def folder(self) -> str | None:
        return self.book.path + os.sep if self.book.path else None

    def _sheet(self, name: str) -> Worksheet:
        found = self.book.sheet_named(name)
        assert found is not None, name
        return found

    def sheet_key(self, name: str) -> str | None:
        return next((sheet.name for sheet in self.book.sheets_ if sheet.name.casefold() == name.casefold()), None)

    def sheet_order(self) -> list[str]:
        return [sheet.name for sheet in self.book.sheets_]

    def table_names(self) -> list[str]:
        """Every table of the workbook, which the file form of a formula spells its references to with []."""
        return [table.name for sheet in self.book.sheets_ for table in sheet.tables]

    def read(self, formula: str) -> Node:
        """A formula as the model keeps it, Range.Formula's spelling with every table named, read into the engine's
        tree: spelled as a file spells it, Table1[] and [#This Row], with a column's @ kept escaped."""
        from pyopenvba.formula._structured import in_file

        return parse(in_file(formula, self.table_names(), at_escaped=True))

    def cell(self, sheet: str, row: int, column: int) -> Scalar:
        return scalar(self.calculator.value_of(sheet, row, column))

    def cells(self, area: Area) -> Iterator[tuple[int, int, Scalar]]:
        """The cells of an area that hold something, row by row."""
        owner = self._sheet(area.sheet)
        for (top, left), block in list(owner.array_formulas.items()):
            if block.top <= area.bottom and area.top <= block.bottom and block.left <= area.right \
                    and area.left <= block.right:
                # An array formula's other cells are there once it is worked out.
                self.calculator.value_of(area.sheet, top, left)
        held = owner.cells_
        found: list[tuple[int, int, Scalar]] = []
        for row, column in sorted(key for key in held if area.contains(*key)):
            value = self.cell(area.sheet, row, column)
            if not isinstance(value, Empty):
                found.append((row, column, value))
        return iter(found)

    def defined_name(self, name: str, sheet: str | None) -> Node | None:
        """A name's formula, scoped to ``sheet`` or, with None, to the workbook: that scope and no other."""
        from pyopenvba._a1 import split_sheet

        wanted = (sheet or "").casefold()
        for entry in self.book.names_.entries:
            scope, bare = split_sheet(entry.name)
            if bare.casefold() != name.casefold() or scope.casefold() != wanted:
                continue
            text = entry.refers_to.lstrip("=")
            _check_absolute(text)
            try:
                return self.read("=" + text)
            except FormulaSyntaxError:
                return None
        return None

    def table(self, name: str) -> TableShape | None:
        found = self.calculator.table(name)
        if found is None:
            return None
        area = found.area
        return TableShape(found.name, area.sheet, area.top, area.left, area.bottom, area.right, int(found.headers),
                          int(found.totals), found.columns)

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None:
        for table in self._sheet(sheet).tables:
            if table.area.contains(row, column):
                return self.table(table.name)
        return None

    def formula_text(self, sheet: str, row: int, column: int) -> str | None:
        text = self.calculator.formula_at(sheet, row, column)
        return text[1:] if text.startswith("=") else None

    def used(self, sheet: str) -> tuple[int, int]:
        bounds = self._sheet(sheet).used_bounds()
        return (0, 0) if bounds is None else (bounds[2], bounds[3])

    def row_hidden(self, sheet: str, row: int) -> bool:
        return row in self.calculator.hidden_rows(sheet, True)

    def row_filtered(self, sheet: str, row: int) -> bool:
        return row in self.calculator.hidden_rows(sheet, False)

    def subtotal(self, sheet: str, row: int, column: int) -> bool:
        cell = self._sheet(sheet).cells_.get((row, column))
        if cell is None or not cell.formula:
            return False
        try:
            node = parse(cell.formula)
        except FormulaSyntaxError:
            return False
        return any(isinstance(one, Call) and one.function in _SUBTOTALS for one in walk(node))

    def spill(self, sheet: str, row: int, column: int) -> Area | None:
        """The model has no dynamic-array formulas, so no block a formula spilled into."""
        return None

    # --- what a formula reads -----------------------------------------------------------------

    def precedents(self, node: Node, sheet: str, row: int, column: int, names: frozenset[str] = frozenset()) \
            -> list[ModelArea]:
        """The blocks a formula reads, as far as its text says: its references, those of the names it uses, and the
        whole of a table it names. A reference worked out as it runs, OFFSET's or INDIRECT's, is not among them; such a
        formula is volatile instead."""
        found: list[ModelArea] = []
        for one in walk(node):
            if isinstance(one, (CellReference, AreaReference, AxisReference)):
                found.extend(_block(one, name) for name in self._sheets_of(one.prefix, sheet))
            elif isinstance(one, NameReference) and one.name.casefold() not in names:
                formula = self._name_formula(one, sheet)
                if formula is not None:
                    found.extend(self.precedents(formula, sheet, row, column, names | {one.name.casefold()}))
            elif isinstance(one, StructuredReference):
                table = self.table(one.table) if one.table is not None else self.table_at(sheet, row, column)
                if table is not None:
                    # All of the table: a reference to this row reads one row of it, but which depends on the cell.
                    found.append(ModelArea(table.top, table.left, table.bottom, table.right, table.sheet))
        return found

    def _name_formula(self, node: NameReference, sheet: str) -> Node | None:
        """The formula of the defined name a name in a formula stands for, found as the engine finds it: the
        sheet's own name first, then the workbook's."""
        try:
            if node.prefix is not None:
                scope = self.sheet_key(node.prefix.sheet) if node.prefix.sheet and node.prefix.book is None else None
                return self.defined_name(node.name, scope) if scope is not None else None
            return self.defined_name(node.name, sheet) or self.defined_name(node.name, None)
        except VBAUnsupportedError:
            return None

    def _sheets_of(self, prefix: Prefix | None, sheet: str) -> list[str]:
        """The sheets a reference's prefix names, in the workbook's order; none for another workbook's."""
        if prefix is None:
            return [sheet]
        if prefix.book is not None or prefix.sheet is None:
            return []
        first = self.sheet_key(prefix.sheet)
        last = self.sheet_key(prefix.last_sheet) if prefix.last_sheet is not None else first
        if first is None or last is None:
            return []
        order = self.sheet_order()
        start, end = sorted((order.index(first), order.index(last)))
        return order[start:end + 1]


def _block(node: CellReference | AreaReference | AxisReference, sheet: str) -> ModelArea:
    if isinstance(node, CellReference):
        return ModelArea(node.ref.row, node.ref.column, node.ref.row, node.ref.column, sheet)
    if isinstance(node, AreaReference):
        top, bottom = sorted((node.first.row, node.last.row))
        left, right = sorted((node.first.column, node.last.column))
        return ModelArea(top, left, bottom, right, sheet)
    axis = node.ref
    if axis.is_row:
        return ModelArea(axis.low, 1, axis.high, MAX_COLUMNS, sheet)
    return ModelArea(1, axis.low, MAX_ROWS, axis.high, sheet)


def volatile(node: Node) -> bool:
    """Whether a formula has to be worked out after every edit, whatever it names."""
    for one in walk(node):
        if isinstance(one, Call):
            entry = registry.FUNCTIONS.get(one.function)
            if one.function in _VOLATILE or (entry is not None and entry.volatile):
                return True
    return False


def _check_absolute(text: str) -> None:
    """A defined name whose formula has a relative reference reports itself, a lone reference such as Sheet1!A1
    too."""
    try:
        tokens = P.tokenize(text)
    except P.FormulaError:
        return
    for token in tokens:
        sides = P.split_sheet(token.text)[1].split(":") if token.kind == "ref" else []
        if not all(_ABSOLUTE.fullmatch(side) for side in sides):
            raise VBAUnsupportedError("a defined name whose formula has a relative reference is not implemented")


__all__ = ["EngineBook", "model_value", "scalar", "volatile"]
