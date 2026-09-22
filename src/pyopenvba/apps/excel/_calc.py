"""Recalculation: which cells are stale, and what they come to.

The engine in :mod:`pyopenvba.formula` evaluates one formula against a
grid.  This is the grid, and the bookkeeping around it: which cells hold
formulas, which cells feed which, and what has to be worked out again
after a macro writes somewhere.

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
from pyopenvba.formula import _parse as P
from pyopenvba.formula._engine import Context, clip, evaluate
from pyopenvba.formula._values import BLANK, ExcelError, Matrix, single
from pyopenvba.interpreter._values import EMPTY, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Workbook

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
    """One cell's formula, parsed once."""

    node: P.Node
    precedents: list[Area] = field(default_factory=lambda: [])
    volatile: bool = False


class Calculator:
    """A workbook's formulas, and what they come to."""

    def __init__(self, book: Workbook) -> None:
        self.book = book
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
        node = P.parse(cell.formula)
        context = Context(self, sheet, row, column)
        areas: list[Area] = []
        for reference in P.references(node):
            try:
                areas.append(context.resolve(reference))
            except ExcelError:
                continue
        for named in P.names(node):
            found = self.named(named.name, named.sheet or sheet)
            if isinstance(found, Area):
                areas.append(found)
        self.compiled[key] = Compiled(node=node, precedents=areas, volatile=P.is_volatile(node))

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
                seen.add(key)
        if following:
            self._spoil(following, seen)

    def _cell(self, key: CellKey) -> Cell | None:
        for sheet in self.book.sheets_:
            if sheet.name.lower() == key[0]:
                return sheet.cells_.get((key[1], key[2]))
        return None

    def _sheet_named(self, name: str) -> object:
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

    def value_of(self, sheet: str, row: int, column: int, *, force: bool = False) -> object:
        """A cell's value, worked out first if it needs to be."""
        self.build()
        key = (sheet.lower(), row, column)
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
        if cell.number_format in ("General", ""):
            from pyopenvba.formula._functions import result_format

            wanted = result_format(compiled.node)
            if wanted:
                cell.number_format = wanted
        return value

    def _computed(self, compiled: Compiled, sheet: str, row: int, column: int) -> object:
        try:
            answer = single(evaluate(compiled.node, Context(self, sheet, row, column)))
        except ExcelError as failure:
            if failure.name == "#CIRCULAR!":
                # Excel leaves a zero in a cell that feeds itself.
                self.circular.add((sheet.lower(), row, column))
                return 0.0
            return failure
        except RecursionError:
            self.circular.add((sheet.lower(), row, column))
            return 0.0
        return _to_cell(answer)

    # --- what the engine asks of a grid ------------------------------------------------

    def cell_value(self, sheet: str, row: int, column: int) -> object:
        return from_vba(self.value_of(sheet, row, column))

    def block(self, sheet: str, area: Area) -> Matrix:
        target = self._sheet_named(area.sheet or sheet)
        if target is None:
            from pyopenvba.formula._values import REF

            raise REF
        bounded = clip(area, target.used_bounds())  # type: ignore[attr-defined]
        rows: list[list[object]] = []
        for row in range(bounded.top, bounded.bottom + 1):
            line: list[object] = []
            for column in range(bounded.left, bounded.right + 1):
                line.append(self.cell_value(target.name, row, column))  # type: ignore[attr-defined]
            rows.append(line)
        return Matrix(rows if rows else [[BLANK]])

    def named(self, name: str, sheet: str) -> object:
        found = self.book.names_.find(name, scope=self.book.sheet_named(sheet))
        if found is None:
            return None
        text = found.entry.refers_to.lstrip("=")
        from pyopenvba._a1 import parse_area

        try:
            area = parse_area(text, sheet=sheet)
        except ValueError:
            node = P.parse(text)
            if isinstance(node, P.Literal):
                return node.value
            from pyopenvba.apps.excel._control_refs import binding
            from pyopenvba.exceptions import VBAUnsupportedError

            try:
                _, area = binding(self.book.sheet_named(sheet), found.entry.name)
            except (ValueError, VBAUnsupportedError):
                return None
            return area
        return area

    def sheet_exists(self, name: str) -> bool:
        return self._sheet_named(name) is not None

    def formula_at(self, sheet: str, row: int, column: int) -> str:
        cell = self._cell((sheet.lower(), row, column))
        return cell.formula if cell is not None else ""

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


def _to_cell(value: object) -> object:
    """A computed value as the cell stores it."""
    if value is BLANK:
        # A formula that comes to nothing shows a zero: =A1 on an empty
        # cell is 0, not empty.
        return 0.0
    if isinstance(value, Matrix):
        return _to_cell(value.first())
    return value


def as_vba(value: object, cell: Cell | None = None) -> object:
    """A cell's value as VBA reads it.

    An error becomes a Variant/Error, and a number in a date-formatted
    cell becomes a Date, which is how VBA sees one: the format is what
    makes 44259 into the fourth of March.
    """
    if isinstance(value, ExcelError):
        return VBAErrorValue(ERROR_NUMBERS.get(value.name, 2015))
    if cell is not None and isinstance(value, (int, float)) and not isinstance(value, bool):
        from pyopenvba.apps.excel._io import is_date_format

        if is_date_format(cell.number_format):
            return VBADate(float(value))
    return value
