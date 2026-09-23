"""Recalculation: which cells are stale, and what they come to.

A cell's formula is worked out by :mod:`pyopenvba.formula._calc`, which
reads the workbook through :mod:`pyopenvba.apps.excel._engine_book`.
This is the bookkeeping around it: which cells hold formulas, which
cells feed which, and what has to be worked out again after a macro
writes somewhere. Evaluate, WorksheetFunction, a validation's formula, a
control's link and the format a formula gives its cell go through the
same engine.

Calculation is on demand.  Reading a stale cell computes it, and
computing it computes whatever it reads, so nothing is ordered up front
and a formula that reaches through INDIRECT is no harder than one that
does not.  In manual mode nothing is computed at all: the cell keeps the
value it last had, which is what Excel shows until F9.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.apps.excel._arrays import array_at
from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS, EngineBook, model_value, volatile
from pyopenvba.formula._calc.evaluator import Context as EngineContext
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.values import Array, ExcelError as EngineError, Scalar
from pyopenvba.formula._deep import deep
from pyopenvba.formula._structured import TableShape
from pyopenvba.formula._values import BLANK, REF, ExcelError
from pyopenvba.interpreter._values import EMPTY, VBACurrency, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet

CellKey = tuple[str, int, int]

#: How many cells deep one cell is worked out inside another before the next is put off (see Calculator.value_of).
_DEEPEST = 24


class _Deeper(BaseException):
    """A cell too deep to work out inside the others, for the cell read from the top to work out first. A
    BaseException, so no ``except Exception`` a formula's function runs through takes it for the formula failing."""

    def __init__(self, sheet: str, row: int, column: int) -> None:
        super().__init__(sheet, row, column)
        self.sheet = sheet
        self.row = row
        self.column = column


@dataclass(slots=True)
class Compiled:
    """One cell's formula, parsed once; no node for one the parser cannot read, which keeps the value it has."""

    node: Node | None
    precedents: list[Area] = field(default_factory=lambda: [])
    volatile: bool = False
    #: Why the parser could not read the formula.
    unread: str = ""


class Calculator:
    """A workbook's formulas, and what they come to."""

    def __init__(self, book: Workbook) -> None:
        self.book = book
        #: The workbook as the formula engine reads it.
        self.engine_book = EngineBook(self)
        self.compiled: dict[CellKey, Compiled] = {}
        #: Which formulas read each cell a formula names on its own, and the blocks of cells each other formula reads,
        #: so that a change finds what reads it without going through every formula.
        self._readers: dict[CellKey, set[CellKey]] = {}
        self._blocks: dict[CellKey, list[Area]] = {}
        #: Cells that could not be worked out because they feed
        #: themselves.  Excel shows zero in these and warns separately.
        self.circular: set[CellKey] = set()
        self._built = False
        #: The cells being worked out, one inside another, and how many.
        self._working: set[CellKey] = set()
        self._depth = 0
        #: The cells put off until a cell too deep to work out inside them is worked out.
        self._waiting: set[CellKey] = set()

    # --- keeping track ------------------------------------------------------------

    def build(self) -> None:
        """Parse every formula in the workbook, once."""
        if self._built:
            return
        self._built = True
        for sheet in self.book.sheets_:
            for (row, column), cell in list(sheet.cells_.items()):
                if cell.formula:
                    self.remember(sheet.name, row, column, cell)

    def remember(self, sheet: str, row: int, column: int, cell: Cell) -> None:
        """Take note of a cell's formula, or forget it if there is none."""
        key = (sheet.lower(), row, column)
        if not cell.formula:
            self._keep(key, None)
            return
        formula = cell.formula
        try:
            compiled = deep(lambda: self._compiled(formula, sheet, row, column))
        except FormulaSyntaxError as failure:
            # One formula the parser cannot read stops no other from being worked out.
            compiled = Compiled(node=None, unread=str(failure))
        self._keep(key, compiled)

    def _compiled(self, formula: str, sheet: str, row: int, column: int) -> Compiled:
        node = self.engine_book.read(formula)
        precedents = self.engine_book.precedents(node, sheet, row, column)
        return Compiled(node=node, precedents=precedents, volatile=volatile(node))

    def forget(self, sheet: str, row: int, column: int) -> None:
        """A cell is gone: forget its formula."""
        self._keep((sheet.lower(), row, column), None)

    def _keep(self, key: CellKey, compiled: Compiled | None) -> None:
        """Keep ``compiled`` as the formula at ``key``, or none, and note which cells it reads."""
        old = self.compiled.pop(key, None)
        if old is not None:
            for area in old.precedents:
                readers = self._readers.get((area.sheet.lower(), area.top, area.left))
                if readers is not None:
                    readers.discard(key)
            self._blocks.pop(key, None)
        if compiled is None:
            return
        self.compiled[key] = compiled
        blocks: list[Area] = []
        for area in compiled.precedents:
            if area.top == area.bottom and area.left == area.right:
                self._readers.setdefault((area.sheet.lower(), area.top, area.left), set()).add(key)
            else:
                blocks.append(area)
        if blocks:
            self._blocks[key] = blocks

    def _reading(self, name: str, row: int, column: int) -> set[CellKey]:
        """The formulas that read a cell."""
        found = set(self._readers.get((name, row, column), ()))
        found.update(key for key, blocks in self._blocks.items()
                     if any(area.sheet.lower() == name and area.contains(row, column) for area in blocks))
        return found

    def rebuild(self) -> None:
        """Forget every parsed formula and mark them all for another look.

        Deleting a row moves what every formula was pointing at, and
        working out which ones moved is the same work as reading them
        all again.
        """
        self.compiled.clear()
        self._readers.clear()
        self._blocks.clear()
        self.circular.clear()
        self._built = False
        for sheet in self.book.sheets_:
            for cell in sheet.cells_.values():
                if cell.formula:
                    cell.stale = True

    def wrote(self, sheet: str, row: int, column: int) -> None:
        """A cell changed: mark whatever reads it as needing another look."""
        self.build()
        self._spoil({(sheet.lower(), row, column)}, set())

    def wrote_area(self, sheet: str, area: Area) -> None:
        """A block of cells changed at once, as a copy changes it: forget the formulas compiled there, and mark
        whatever reads the block, and what reads that, as needing another look."""
        self.build()
        name = sheet.lower()
        for key in [key for key in self.compiled if key[0] == name and area.contains(key[1], key[2])]:
            self._keep(key, None)
        reading: set[CellKey] = set()
        for key, compiled in self.compiled.items():
            if any(one.sheet.lower() == name and one.top <= area.bottom and area.top <= one.bottom
                   and one.left <= area.right and area.left <= one.right for one in compiled.precedents):
                cell = self._cell(key)
                if cell is not None and not cell.stale:
                    cell.stale = True
                    reading.add(key)
                    reading.update(self._members(key))
        if reading:
            self._spoil(reading, set(reading))

    def _spoil(self, changed: set[CellKey], seen: set[CellKey]) -> None:
        """Mark stale what reads ``changed``, and what reads that, a wave at a time."""
        while changed:
            following: set[CellKey] = set()
            for name, row, column in changed:
                for key in self._reading(name, row, column) - seen:
                    cell = self._cell(key)
                    if cell is not None and not cell.stale:
                        cell.stale = True
                        following.add(key)
                        # What reads a cell of an array formula follows the array.
                        following.update(self._members(key))
                    seen.add(key)
            changed = following

    def _cell(self, key: CellKey) -> Cell | None:
        for sheet in self.book.sheets_:
            if sheet.name.lower() == key[0]:
                return sheet.cells_.get((key[1], key[2]))
        return None

    def _sheet_named(self, name: str) -> Worksheet | None:
        for sheet in self.book.sheets_:
            if sheet.name.lower() == name.lower():
                return sheet
        return None

    # --- calculating ---------------------------------------------------------------

    @property
    def automatic(self) -> bool:
        """Whether the application recalculates as things change."""
        return self.book.application.calculation != -4135  # xlCalculationManual

    def calculate_all(self) -> None:
        """Work out every stale cell, as F9 does."""
        self.build()
        for sheet in self.book.sheets_:
            for (row, column), cell in list(sheet.cells_.items()):
                if cell.formula and cell.stale:
                    self.value_of(sheet.name, row, column, force=True)

    def recalculated(self) -> list[Worksheet]:
        """Work out every formula an edit left stale, and the volatile ones, as Excel does after the edit; the
        sheets that had formulas worked out, in order. Nothing, in manual calculation."""
        self.build()
        if not self.automatic:
            return []
        worked: list[Worksheet] = []
        for sheet in self.book.sheets_:
            did = False
            for (row, column), cell in list(sheet.cells_.items()):
                if not cell.formula:
                    continue
                compiled = self.compiled.get((sheet.name.lower(), row, column))
                if cell.stale or (compiled is not None and compiled.volatile):
                    self.value_of(sheet.name, row, column, force=True)
                    did = True
            if did:
                worked.append(sheet)
        return worked

    def value_of(self, sheet: str, row: int, column: int, *, force: bool = False) -> object:
        """A cell's value, worked out first if it needs to be; a cell of an array formula, by working out the array.

        A cell read from outside any formula is worked out from here. A cell
        more than :data:`_DEEPEST` cells down is put off: it is worked out on
        its own, and the cell that wanted it tried again. A running total
        thousands of rows long comes out as it does in Excel, and no chain of
        cells runs Python out of stack.
        """
        if self._depth:
            return self._value(sheet, row, column, force)
        wanted = [(sheet, row, column, force)]
        waiting: list[CellKey] = []
        try:
            while True:
                try:
                    value = self._value(*wanted[-1])
                except _Deeper as deeper:
                    top = wanted[-1]
                    waiting.append((top[0].lower(), top[1], top[2]))
                    self._waiting.add(waiting[-1])
                    wanted.append((deeper.sheet, deeper.row, deeper.column, False))
                    continue
                wanted.pop()
                if not wanted:
                    return value
                self._waiting.discard(waiting.pop())
        finally:
            self._waiting.difference_update(waiting)

    def _value(self, sheet: str, row: int, column: int, force: bool) -> object:
        self.build()
        key = (sheet.lower(), row, column)
        owner = self._sheet_named(sheet)
        if owner is not None and owner.array_formulas:
            found = array_at(owner, row, column)
            if found is not None and found[0] != (row, column):
                self._value(sheet, *found[0], force)
                member = owner.cells_.get((row, column))
                return EMPTY if member is None else member.value
        cell = self._cell(key)
        if cell is None:
            return EMPTY
        if not cell.formula:
            return cell.value
        compiled = self.compiled.get(key)
        if compiled is None:
            self.remember(sheet, row, column, cell)
            compiled = self.compiled[key]
        if not (cell.stale or compiled.volatile or force):
            return cell.value
        if not self.automatic and not force:
            return cell.value
        if key in self._working or key in self._waiting:
            self.circular.add(key)
            raise ExcelError("#CIRCULAR!")
        if self._depth >= _DEEPEST:
            raise _Deeper(sheet, row, column)
        self._working.add(key)
        self._depth += 1
        try:
            value = self._computed(compiled, sheet, row, column)
        finally:
            self._depth -= 1
            self._working.discard(key)
        cell.value = value
        cell.stale = False
        from pyopenvba.apps.excel._controls import cell_changed

        owner = next(one for one in self.book.sheets_ if one.name.lower() == sheet.lower())
        # A control's cell is read from the top, so that nothing put off stops the control short.
        depth, self._depth = self._depth, 0
        try:
            cell_changed(owner, row, column, value)
        finally:
            self._depth = depth
        return value

    def _computed(self, compiled: Compiled, sheet: str, row: int, column: int) -> object:
        if compiled.node is None:
            raise VBAUnsupportedError(f"working out a formula the model cannot read is not implemented: "
                                      f"{compiled.unread}")
        owner = self._worksheet(sheet)
        block = owner.array_formulas.get((row, column))
        now = self.now()
        node = compiled.node

        def context() -> EngineContext:
            return EngineContext(self.engine_book, owner.name, row, column, array=block is not None, today=now.date(),
                                 now=now)

        def one() -> Scalar:
            working = context()
            return working.first(working.formula(node))

        def whole() -> Array:
            working = context()
            return working.array_of(working.formula(node))

        try:
            # Worked out afresh on a deeper stack if the formula's tree is deeper than Python's.
            if block is None:
                return model_value(deep(one))
            return _spread(owner, block, deep(whole))
        except EngineError as failure:
            return model_value(failure.error) if block is None else _spread(owner, block, Array([[failure.error]]))
        except ExcelError as failure:
            if failure.name != "#CIRCULAR!":
                raise
            # Excel leaves a zero in a cell that feeds itself.
            self.circular.add((sheet.lower(), row, column))
            return 0.0
        except RecursionError:
            self.circular.add((sheet.lower(), row, column))
            return 0.0

    def _members(self, key: CellKey) -> set[CellKey]:
        """The other cells of the array formula whose first cell is ``key``, if it is one."""
        owner = self._sheet_named(key[0])
        block = owner.array_formulas.get((key[1], key[2])) if owner is not None else None
        if block is None:
            return set()
        return {(key[0], row, column) for row in range(block.top, block.bottom + 1)
                for column in range(block.left, block.right + 1)} - {key}

    # --- what the engine's book asks of the workbook ---------------------------------------

    def table(self, name: str) -> TableShape | None:
        from pyopenvba.apps.excel._tables import shape

        wanted = name.lower()
        for sheet in self.book.sheets_:
            for table in sheet.tables:
                if table.name.lower() == wanted:
                    return shape(table)
        return None

    def _worksheet(self, name: str) -> Worksheet:
        """The sheet a reference names; one the workbook has not got is #REF!."""
        found = self._sheet_named(name)
        if found is None:
            raise REF
        return found

    def hidden_rows(self, sheet: str, every: bool) -> set[int]:
        """The rows SUBTOTAL passes over on a sheet.

        Measured (scripts/measure_subtotal.py): 101 to 111 pass over every
        hidden row, 1 to 11 only the rows a filter hid -- and while the
        sheet is in filter mode Excel takes every hidden row for one,
        whatever hid it, inside the filter's range or not.
        """
        target = self._worksheet(sheet)
        dims = target.dims
        if dims.zero_height:
            raise VBAUnsupportedError("SUBTOTAL over a sheet whose rows are hidden by default is not implemented")
        found = target.auto_filter
        if not every and (found is None or not found.fields):
            return set()
        return {row for row, record in dims.rows.items() if record.hidden}

    def formula_at(self, sheet: str, row: int, column: int) -> str:
        from pyopenvba.apps.excel._model import shown_formula

        cell = self._cell((sheet.lower(), row, column))
        owner = self._sheet_named(sheet)
        if cell is None or not cell.formula or owner is None:
            return cell.formula if cell is not None else ""
        return shown_formula(owner, row, column, cell.formula)

    def now(self) -> _dt.datetime:
        interpreter = self.book.application.interpreter
        if interpreter is not None:
            return interpreter.clock()
        return _dt.datetime.now()


# --- moving values across the boundary ------------------------------------------------------


def from_vba(value: object) -> object:
    """A cell's stored value as the formula engine sees it."""
    if value is EMPTY:
        return BLANK
    if isinstance(value, bool):
        return value
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, VBAInt):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, VBAErrorValue):
        for name, number in ERROR_NUMBERS.items():
            if number == value.number:
                return ExcelError(name)
        return ExcelError("#VALUE!")
    return value


def _spread(sheet: Worksheet, block: Area, answer: Array) -> object:
    """An array formula's answer laid over its block, as Array.at lines it up; the first cell's item, for it to keep.

    One row or column repeats across the block, and past the end of the
    answer a cell shows #N/A.
    """
    first: object = 0.0
    for row in range(block.top, block.bottom + 1):
        for column in range(block.left, block.right + 1):
            value = model_value(answer.at(row - block.top, column - block.left))
            if (row, column) == (block.top, block.left):
                first = value
                continue
            cell = sheet.cell(row, column, create=True)
            assert cell is not None
            cell.value, cell.stale = value, False
    return first


def as_vba(value: object, cell: Cell | None = None) -> object:
    """A cell's value as VBA reads it.

    An error becomes a Variant/Error. A cell holds every number as a
    Double, and its format decides what Range.Value makes of one: a date
    format a Date, which is what makes 44259 the fourth of March, and a
    dollar sign a Currency. A number the Date or Currency cannot hold is
    error 6, as Range.Value raises it, a block's read included. With no
    cell -- Value2 -- a number stays a Double.
    """
    if isinstance(value, ExcelError):
        return VBAErrorValue(ERROR_NUMBERS.get(value.name, 2015))
    if isinstance(value, VBADate):
        value = value.serial
    if cell is None or isinstance(value, bool) or not isinstance(value, (int, float)):
        return value
    from pyopenvba.apps.excel._typing import is_currency_format, is_date_format

    code = cell.number_format
    if is_date_format(code):
        return VBADate(float(value))
    if is_currency_format(code):
        return VBACurrency(value)
    return float(value)


def as_python(value: object, cell: Cell | None = None) -> object:
    """A cell's value as ``as_vba`` reads it, but the Double where its Date or Currency cannot hold it."""
    try:
        return as_vba(value, cell)
    except VBARuntimeError:
        return as_vba(value)
