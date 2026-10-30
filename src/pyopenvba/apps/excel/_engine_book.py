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

import dataclasses
import os
import re
import typing
from collections.abc import Iterator
from typing import TYPE_CHECKING

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area as ModelArea
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._locale import Locale
from pyopenvba.formula import _parse as P
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc import registry
from pyopenvba.formula._calc.cells import CellError, UnsupportedFormulaError
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
    from pyopenvba.apps.excel._external_links import LinkedBook
    from pyopenvba.apps.excel._pivot_reports import PivotReport
    from pyopenvba.apps.excel._calc import Calculator
    from pyopenvba.apps.excel._model import Worksheet
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.values import Value

#: The functions whose cells SUBTOTAL and AGGREGATE leave out.
_SUBTOTALS = frozenset({"SUBTOTAL", "AGGREGATE"})
#: Functions whose answer can change with no cell they name changing: a reference worked out as they run, the
#: clock, chance, or rows hidden. Excel works them out after every edit.
_VOLATILE = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO",
                       "SUBTOTAL"})
#: One side of a reference that names its cells outright: $A$1, a whole column $A or a whole row $1.
_ABSOLUTE = re.compile(r"\$[A-Za-z]{1,3}(?:\$[0-9]+)?|\$[0-9]+")
#: VBA's CVErr numbers for the errors a cell holds, and what it answers when VBA asks a cell holding one for its
#: Value. A cell takes CVErr of these numbers and refuses any other with error 1004 (tests/fixtures/excel_model/).
ERROR_NUMBERS: dict[str, int] = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029,
                                 "#NUM!": 2036, "#N/A": 2042, "#GETTING_DATA": 2043, "#SPILL!": 2045,
                                 "#CONNECT!": 2046, "#BLOCKED!": 2047, "#UNKNOWN!": 2048, "#FIELD!": 2049,
                                 "#CALC!": 2050}
_ERROR_NAMES = {number: name for name, number in ERROR_NUMBERS.items()}
#: The cells of a link saved with SaveLinkValues False, which keeps none: each reads #REF! once worked out, a
#: reference keeping its shape (tests/fixtures/workbook_links.json, Never.xlsx).
_NO_VALUES: dict[tuple[int, int], object] = {}


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

    pivot_evaluate = False

    def __init__(self, calculator: Calculator) -> None:
        self.calculator = calculator
        self.book = calculator.book

    @property
    def epoch_1904(self) -> bool:
        return self.book.epoch_1904

    @property
    def locale(self) -> Locale:
        """The Regional format TEXT, DOLLAR and FIXED show values in."""
        return self.book.application.locale

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
        from pyopenvba.apps.excel._names import same_sheet_name

        return next((sheet.name for sheet in self.book.sheets_ if same_sheet_name(sheet.name, name)), None)

    def sheet_order(self) -> list[str]:
        return [sheet.name for sheet in self.book.sheets_]

    def _linked(self, book: str) -> LinkedBook | None:
        """The linked workbook a formula names by its link's number, as a file spells it, or by its file's name."""
        books = self.book.linked_books
        if book.isdigit():
            return books.get(int(book))
        return next((linked for linked in books.values() if linked.file.casefold() == book.casefold()), None)

    def external_sheets(self, book: str, first: str, last: str | None) -> list[str] | None:
        """The sheets of a linked workbook a reference names, as the engine keys them, [1]Sheet1, in that workbook's
        order; none where it has no such sheet, and None where the workbook has no such link. Their cells are the
        ones the link keeps (see _external_links)."""
        linked = self._linked(book)
        if linked is None:
            return None
        names = [sheet.casefold() for sheet in linked.sheets]
        try:
            start = names.index(first.casefold())
            end = start if last is None else names.index(last.casefold())
        except ValueError:
            return []
        low, high = sorted((start, end))
        return [f"[{linked.number}]{sheet}" for sheet in linked.sheets[low:high + 1]]

    def external_name(self, book: str, name: str) -> Node | None:
        """A linked workbook's own name, as a formula reading that workbook, or None where its link has no such name."""
        linked = self._linked(book)
        if linked is None:
            raise UnsupportedFormulaError("a name defined in another workbook the file keeps no copy of")
        refers = next((text for key, text in linked.names.items() if key.casefold() == name.casefold()), None)
        if refers is None:
            return None
        try:
            return _in_book(parse(refers.removeprefix("=")), str(linked.number))
        except FormulaSyntaxError:
            return None

    def _kept(self, sheet: str) -> dict[tuple[int, int], object] | None:
        """The cells a link keeps of a linked workbook's sheet keyed [1]Sheet1, or None for one of this workbook."""
        if not sheet.startswith("["):
            return None
        number, _, name = sheet[1:].partition("]")
        linked = self.book.linked_books.get(int(number)) if number.isdigit() else None
        if linked is None:
            return {}
        if not linked.kept:
            return _NO_VALUES
        index = next((at for at, one in enumerate(linked.sheets) if one.casefold() == name.casefold()), None)
        return {} if index is None else linked.cells[index]

    def pivot_reports(self, sheet: str) -> list[PivotReport]:
        reports: list[PivotReport] = []
        for table in self._sheet(sheet).pivot_tables:
            report = table.report()
            if report is not None:
                reports.append(report)
        return reports

    def table_names(self) -> list[str]:
        """Every table of the workbook, which the file form of a formula spells its references to with []."""
        return [table.name for sheet in self.book.sheets_ for table in sheet.tables]

    def read(self, formula: str) -> Node:
        """A formula as the model keeps it, Range.Formula's spelling with every table named, read into the engine's
        tree: spelled as a file spells it, Table1[] and [#This Row], with a column's @ kept escaped, and each name a
        LET or a LAMBDA binds _xlpm.x, x1 among them, which the engine would otherwise read as a cell."""
        from pyopenvba.formula._prefixes import in_file as prefixed
        from pyopenvba.formula._structured import in_file

        return parse(prefixed(in_file(formula, self.table_names(), at_escaped=True)))

    def cell(self, sheet: str, row: int, column: int) -> Scalar:
        kept = self._kept(sheet)
        if kept is _NO_VALUES:
            return CellError("#REF!")
        if kept is not None:
            # A cell the link does not keep is blank.
            return scalar(kept.get((row, column), NOTHING_THERE))
        return scalar(self.calculator.value_of(sheet, row, column))

    def user_function(self, context: Context, name: str, args: tuple[Node, ...]) -> Value | None:
        """A function the engine does not know, worked out by the project's VBA Function of the name, if any."""
        from pyopenvba.apps.excel._udf import call_user_function

        return call_user_function(self.book, context, name, args)

    def cells(self, area: Area) -> Iterator[tuple[int, int, Scalar]]:
        """The cells of an area that hold something, row by row, each worked out as it is reached: a lookup that
        stops early reads no further, as Excel's does, and so closes no loop through the cells past its match
        (tests/fixtures/self_reference.json)."""
        kept = self._kept(area.sheet)
        if kept is _NO_VALUES:
            if (area.bottom - area.top + 1) * (area.right - area.left + 1) > MAX_ROWS:
                raise UnsupportedFormulaError("more than a column's cells of a link that keeps no values")
            for row in range(area.top, area.bottom + 1):
                for column in range(area.left, area.right + 1):
                    yield row, column, CellError("#REF!")
            return
        if kept is not None:
            yield from [(row, column, scalar(kept[(row, column)]))
                        for row, column in sorted(key for key in kept if area.contains(*key))]
            return
        owner = self._sheet(area.sheet)
        for (top, left), block in list(owner.array_formulas.items()):
            if block.top <= area.bottom and area.top <= block.bottom and block.left <= area.right \
                    and area.left <= block.right:
                # An array formula's other cells are there once it is worked out.
                self.calculator.value_of(area.sheet, top, left)
        for row, column in sorted(key for key in owner.cells_ if area.contains(*key)):
            value = self.cell(area.sheet, row, column)
            if not isinstance(value, Empty):
                yield row, column, value

    def defined_name(self, name: str, sheet: str | None) -> Node | None:
        """A name's formula, scoped to ``sheet`` or, with None, to the workbook: that scope and no other, and a name
        word sort cannot tell from ``name`` (tests/fixtures/name_equality.json)."""
        from pyopenvba._a1 import split_sheet
        from pyopenvba.apps.excel._names import same_name, same_sheet_name

        for entry in self.book.names_.entries:
            scope, bare = split_sheet(entry.name)
            if not same_name(bare, name) or not same_sheet_name(scope, sheet or ""):
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

    # A link keeps a linked sheet's values and nothing else: no tables, formulas, hidden rows or spills.

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None:
        if self._kept(sheet) is not None:
            return None
        for table in self._sheet(sheet).tables:
            if table.area.contains(row, column):
                return self.table(table.name)
        return None

    def formula_text(self, sheet: str, row: int, column: int) -> str | None:
        if self._kept(sheet) is not None:
            return None
        text = self.calculator.formula_at(sheet, row, column)
        return text[1:] if text.startswith("=") else None

    def used(self, sheet: str) -> tuple[int, int]:
        kept = self._kept(sheet)
        if kept is not None:
            return max((row for row, _ in kept), default=0), max((column for _, column in kept), default=0)
        bounds = self._sheet(sheet).used_bounds()
        return (0, 0) if bounds is None else (bounds[2], bounds[3])

    def row_hidden(self, sheet: str, row: int) -> bool:
        return self._kept(sheet) is None and row in self.calculator.hidden_rows(sheet, True)

    def row_filtered(self, sheet: str, row: int) -> bool:
        return self._kept(sheet) is None and row in self.calculator.hidden_rows(sheet, False)

    def subtotal(self, sheet: str, row: int, column: int) -> bool:
        if self._kept(sheet) is not None:
            return False
        cell = self._sheet(sheet).cells_.get((row, column))
        if cell is None or not cell.formula:
            return False
        try:
            node = parse(cell.formula)
        except FormulaSyntaxError:
            return False
        return any(isinstance(one, Call) and one.function in _SUBTOTALS for one in walk(node))

    def spill(self, sheet: str, row: int, column: int) -> Area | None:
        """The block a dynamic-array formula in a cell spilled into, the formula worked out first; None for a cell
        holding no such formula, or one whose answer is kept out (see _spills)."""
        if self._kept(sheet) is not None:
            return None
        owner = self._sheet(sheet)
        if (row, column) not in owner.spills:
            return None
        self.calculator.value_of(owner.name, row, column)
        spill = owner.spills.get((row, column))
        if spill is None or spill.blocked:
            return None
        return Area(owner.name, spill.area.top, spill.area.left, spill.area.bottom, spill.area.right)

    # --- what a formula reads -----------------------------------------------------------------

    def precedents(self, node: Node, sheet: str, row: int, column: int, names: frozenset[str] = frozenset()) \
            -> list[ModelArea]:
        """The blocks a formula reads, as far as its text says: its references, those of the names it uses, and the
        whole of a table it names. A reference worked out as it runs, OFFSET's or INDIRECT's, is not among them; such a
        formula is volatile instead."""
        found: list[ModelArea] = []
        for one in walk(node):
            if isinstance(one, (CellReference, AreaReference, AxisReference)):
                found.extend(block(one, name) for name in self._sheets_of(one.prefix, sheet))
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

    def visibility_sensitive(self, node: Node, sheet: str, names: frozenset[str] = frozenset()) -> bool:
        """SUBTOTAL and AGGREGATE also depend on row visibility, including through defined names."""
        for one in walk(node):
            if isinstance(one, Call) and one.function in _SUBTOTALS:
                return True
            if isinstance(one, NameReference) and one.name.casefold() not in names:
                formula = self._name_formula(one, sheet)
                if formula is not None and self.visibility_sensitive(formula, sheet, names | {one.name.casefold()}):
                    return True
        return False

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


def block(node: CellReference | AreaReference | AxisReference, sheet: str) -> ModelArea:
    """The cells a reference names on ``sheet``, a whole row or column to the sheet's edge."""
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


def calls_user(node: Node) -> bool:
    """Whether a formula calls a function the engine does not have, which a VBA Function may answer."""
    from pyopenvba.formula._calc.catalog import is_excel_function

    return any(isinstance(one, Call) and one.function not in registry.FUNCTIONS and not is_excel_function(one.name)
               for one in walk(node))


_NODES = typing.get_args(Node)


def _in_book(node: Node, book: str) -> Node:
    """A formula of a linked workbook, as one defining a name there, with each reference and name in it made to name
    that workbook, as pyOfficeEditor's engine does it (its commit 691adba)."""
    changes: dict[str, object] = {}
    for item in dataclasses.fields(node):
        value = getattr(node, item.name)
        if isinstance(value, Prefix):
            if value.book is None:
                changes[item.name] = Prefix(value.sheet, value.last_sheet, book)
        elif isinstance(node, NameReference) and item.name == "prefix" and value is None:
            changes[item.name] = Prefix(book=book)
        elif isinstance(value, _NODES):
            changes[item.name] = _in_book(typing.cast("Node", value), book)
        elif isinstance(value, tuple):
            parts = typing.cast("tuple[object, ...]", value)
            if parts and all(isinstance(part, _NODES) for part in parts):
                changes[item.name] = tuple(_in_book(typing.cast("Node", part), book) for part in parts)
    return dataclasses.replace(node, **changes) if changes else node


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


__all__ = ["ERROR_NUMBERS", "EngineBook", "calls_user", "model_value", "scalar", "volatile"]
