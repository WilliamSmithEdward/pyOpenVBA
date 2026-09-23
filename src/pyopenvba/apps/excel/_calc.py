"""Recalculation: which cells are stale, and what they come to.

A cell's formula is worked out by :mod:`pyopenvba.formula._calc`, which
reads the workbook through :mod:`pyopenvba.apps.excel._engine_book`.
This is the bookkeeping around it: which cells hold formulas, which
cells feed which, and what has to be worked out again after a macro
writes somewhere. Evaluate, WorksheetFunction, a validation's formula, a
control's link and the format a formula gives its cell still go through
the older engine in :mod:`pyopenvba.formula`, for which this is the grid.

Calculation is on demand.  Reading a stale cell computes it, and
computing it computes whatever it reads, so nothing is ordered up front
and a formula that reaches through INDIRECT is no harder than one that
does not.  In manual mode nothing is computed at all: the cell keeps the
value it last had, which is what Excel shows until F9.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.apps.excel._arrays import array_at
from pyopenvba.apps.excel._engine_book import EngineBook, model_value, volatile
from pyopenvba.formula._calc.evaluator import Context as EngineContext
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.values import Array, ExcelError as EngineError
from pyopenvba.formula._engine import clip
from pyopenvba.formula._structured import TableShape, area as structured_area
from pyopenvba.formula._values import BLANK, REF, ExcelError, Matrix
from pyopenvba.interpreter._values import EMPTY, VBACurrency, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet

#: What a cell holding an error answers when VBA asks for its Value.
#: These are Excel's own CVErr numbers.
ERROR_NUMBERS: dict[str, int] = {
    "#NULL!": 2000,
    "#DIV/0!": 2007,
    "#VALUE!": 2015,
    "#REF!": 2023,
    "#NAME?": 2029,
    "#NUM!": 2036,
    "#N/A": 2042,
}

CellKey = tuple[str, int, int]


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
        #: Cells that could not be worked out because they feed
        #: themselves.  Excel shows zero in these and warns separately.
        self.circular: set[CellKey] = set()
        self._built = False
        self._working: set[CellKey] = set()

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
            self.compiled.pop(key, None)
            return
        try:
            node = self.engine_book.read(cell.formula)
        except FormulaSyntaxError as failure:
            # One formula the parser cannot read stops no other from being worked out.
            self.compiled[key] = Compiled(node=None, unread=str(failure))
            return
        precedents = self.engine_book.precedents(node, sheet, row, column)
        self.compiled[key] = Compiled(node=node, precedents=precedents, volatile=volatile(node))

    def rebuild(self) -> None:
        """Forget every parsed formula and mark them all for another look.

        Deleting a row moves what every formula was pointing at, and
        working out which ones moved is the same work as reading them
        all again.
        """
        self.compiled.clear()
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
            del self.compiled[key]
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
        following: set[CellKey] = set()
        for key, compiled in self.compiled.items():
            if key in seen:
                continue
            if any(
                area.sheet.lower() == name and area.contains(row, column)
                for name, row, column in changed
                for area in compiled.precedents
            ):
                cell = self._cell(key)
                if cell is not None and not cell.stale:
                    cell.stale = True
                    following.add(key)
                    # What reads a cell of an array formula follows the array.
                    following.update(self._members(key))
                seen.add(key)
        if following:
            self._spoil(following, seen)

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
        """A cell's value, worked out first if it needs to be; a cell of an array formula, by working out the array."""
        self.build()
        key = (sheet.lower(), row, column)
        owner = self._sheet_named(sheet)
        if owner is not None and owner.array_formulas:
            found = array_at(owner, row, column)
            if found is not None and found[0] != (row, column):
                self.value_of(sheet, *found[0], force=force)
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
        if key in self._working:
            self.circular.add(key)
            raise ExcelError("#CIRCULAR!")
        self._working.add(key)
        try:
            value = self._computed(compiled, sheet, row, column)
        finally:
            self._working.discard(key)
        cell.value = value
        cell.stale = False
        from pyopenvba.apps.excel._controls import cell_changed

        owner = next(one for one in self.book.sheets_ if one.name.lower() == sheet.lower())
        cell_changed(owner, row, column, value)
        return value

    def _computed(self, compiled: Compiled, sheet: str, row: int, column: int) -> object:
        if compiled.node is None:
            raise VBAUnsupportedError(f"working out a formula the model cannot read is not implemented: "
                                      f"{compiled.unread}")
        owner = self._worksheet(sheet)
        block = owner.array_formulas.get((row, column))
        now = self.now()
        context = EngineContext(self.engine_book, owner.name, row, column, array=block is not None, today=now.date(),
                                now=now)
        try:
            value = context.formula(compiled.node)
            if block is None:
                return model_value(context.first(value))
            return _spread(owner, block, context.array_of(value))
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

    # --- what the engine asks of a grid ------------------------------------------------

    def cell_value(self, sheet: str, row: int, column: int) -> object:
        return from_vba(self.value_of(sheet, row, column))

    def block(self, sheet: str, area: Area) -> Matrix:
        target = self._worksheet(area.sheet or sheet)
        bounded = clip(area, target.used_bounds())
        rows: list[list[object]] = []
        for row in range(bounded.top, bounded.bottom + 1):
            line: list[object] = []
            for column in range(bounded.left, bounded.right + 1):
                line.append(self.cell_value(target.name, row, column))
            rows.append(line)
        return Matrix(rows if rows else [[BLANK]])

    def named(self, name: str, sheet: str) -> object:
        found = self.book.names_.find(name, scope=self.book.sheet_named(sheet))
        if found is None:
            # A table's name on its own stands for its data, as Table1[] does.
            table = self.table(name)
            return None if table is None else structured_area(P.Structured(table=table.name), table, 0)
        text = found.entry.refers_to.lstrip("=")
        from pyopenvba._a1 import parse_area

        try:
            if not _one_reference(text):
                # Sheet1!$A$1-Sheet1!$B$1 would read as $B$1 on a sheet called Sheet1!$A$1-Sheet1.
                raise ValueError(text)
            area = parse_area(text, sheet=sheet)
        except ValueError:
            node = P.parse(text)
            if isinstance(node, P.Literal):
                return node.value
            from pyopenvba.apps.excel._control_refs import binding

            try:
                _, area = binding(self.book.sheet_named(sheet), found.entry.name)
            except (ValueError, VBAUnsupportedError):
                area = None
            # A formula that does not land on cells, IF($A$1:$A$3>1,1,0), is worked out where the name is used.
            return _named_formula(node, text) if area is None else area
        return area

    def table(self, name: str) -> TableShape | None:
        from pyopenvba.apps.excel._tables import shape

        wanted = name.lower()
        for sheet in self.book.sheets_:
            for table in sheet.tables:
                if table.name.lower() == wanted:
                    return shape(table)
        return None

    def sheet_exists(self, name: str) -> bool:
        return self._sheet_named(name) is not None

    def _worksheet(self, name: str) -> Worksheet:
        """The sheet a reference names; one the workbook has not got is #REF!."""
        found = self._sheet_named(name)
        if found is None:
            raise REF
        return found

    def used(self, sheet: str) -> tuple[int, int, int, int] | None:
        return self._worksheet(sheet).used_bounds()

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


def _one_reference(text: str) -> bool:
    """Whether a defined name's formula is a reference and nothing else; text the tokenizer cannot read may be."""
    try:
        tokens = P.tokenize(text)
    except P.FormulaError:
        return True
    return len(tokens) == 1 and tokens[0].kind == "ref"


#: One side of a reference that names its cells outright: $A$1, a whole column $A or a whole row $1.
_ABSOLUTE = re.compile(r"\$[A-Za-z]{1,3}(?:\$[0-9]+)?|\$[0-9]+")


def _named_formula(node: P.Node, text: str) -> P.Node:
    """A defined name that stands for a formula, for the engine to work out where the name is used.

    A relative reference in one is counted from the cell the name was
    defined at, which the name does not record here, so it reports itself
    rather than answering from the wrong cells.
    """
    for token in P.tokenize(text):
        sides = P.split_sheet(token.text)[1].split(":") if token.kind == "ref" else []
        if not all(_ABSOLUTE.fullmatch(side) for side in sides):
            raise VBAUnsupportedError("a defined name whose formula has a relative reference is not implemented")
    return node


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
