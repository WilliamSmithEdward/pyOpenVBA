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

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.evaluator import TableShape
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import Call, Node, walk
from pyopenvba.formula._calc.parser import parse
from pyopenvba.formula._calc.values import EMPTY, Area, Empty, Scalar
from pyopenvba.formula._values import BLANK, ExcelError
from pyopenvba.interpreter._values import EMPTY as NOTHING_THERE, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._calc import Calculator
    from pyopenvba.apps.excel._model import Worksheet

#: The functions whose cells SUBTOTAL and AGGREGATE leave out.
_SUBTOTALS = frozenset({"SUBTOTAL", "AGGREGATE"})
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
        from pyopenvba.formula._structured import in_file

        wanted = (sheet or "").casefold()
        for entry in self.book.names_.entries:
            scope, bare = split_sheet(entry.name)
            if bare.casefold() != name.casefold() or scope.casefold() != wanted:
                continue
            text = entry.refers_to.lstrip("=")
            _check_absolute(text)
            try:
                return parse(in_file("=" + text, self.table_names()))
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


def _check_absolute(text: str) -> None:
    """A defined name that stands for more than one reference, with a relative reference in it, reports itself."""
    try:
        tokens = P.tokenize(text)
    except P.FormulaError:
        return
    if len(tokens) == 1 and tokens[0].kind == "ref":
        return
    for token in tokens:
        sides = P.split_sheet(token.text)[1].split(":") if token.kind == "ref" else []
        if not all(_ABSOLUTE.fullmatch(side) for side in sides):
            raise VBAUnsupportedError("a defined name whose formula has a relative reference is not implemented")


__all__ = ["EngineBook", "model_value", "scalar"]
