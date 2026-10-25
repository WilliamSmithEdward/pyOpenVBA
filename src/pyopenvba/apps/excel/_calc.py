"""Recalculation: which cells are stale, and what they come to.

A cell's formula is worked out by :mod:`pyopenvba.formula._calc`, which
reads the workbook through :mod:`pyopenvba.apps.excel._engine_book`.
This is the bookkeeping around it: which cells hold formulas, which
cells feed which, and what has to be worked out again after a macro
writes somewhere. Evaluate, WorksheetFunction, a validation's formula, a
control's link and the format a formula gives its cell go through the
same engine.

Calculation retains learned dependency order and resolves new references
on demand, including INDIRECT. Manual mode keeps caches after input edits;
explicit formula assignments obtain initial values. Pending UDF references
can defer a provisional call until the remaining scheduled cells are ready.
A data table's cell is worked out by trying its values in the table's
input cells (see Calculator._tried), and in semi-automatic calculation
only by a calculation a macro asked for.
"""

from __future__ import annotations

import datetime as _dt
import re
from collections import deque
from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.apps.excel._arrays import array_at
from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS, EngineBook, calls_user, model_value, volatile
from pyopenvba.formula._calc.evaluator import Context as EngineContext
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.values import Array, ExcelError as EngineError, Scalar
from pyopenvba.formula._deep import deep
from pyopenvba.formula._structured import TableShape
from pyopenvba.formula._values import BLANK, REF, ExcelError
from pyopenvba.interpreter._values import EMPTY, VBACurrency, VBADate, VBAErrorValue, VBAInt

if TYPE_CHECKING:
    from pyopenvba.apps.excel._data_tables import DataTable
    from pyopenvba.apps.excel._model import Cell, Workbook, Worksheet

CellKey = tuple[str, int, int]

#: How many cells deep one cell is worked out inside another before the next is put off (see Calculator.value_of).
_DEEPEST = 24
#: What in a name's formula may be another name (see Calculator.names_changed).
_WORDS = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]*")


class _Deeper(BaseException):
    """A cell too deep to work out inside the others, for the cell read from the top to work out first. A
    BaseException, so no ``except Exception`` a formula's function runs through takes it for the formula failing."""

    def __init__(self, sheet: str, row: int, column: int) -> None:
        super().__init__(sheet, row, column)
        self.sheet = sheet
        self.row = row
        self.column = column


class _UDFWaiting(_Deeper):
    """A provisional UDF call ran before its pending UDF precedent was ready."""


class _Loop(BaseException):  # noqa: N818 - how a calculation stops, not an error
    """A formula asked for a cell that is being worked out under it, which closes a loop, or for a cell held in a
    loop: it cannot be worked out, and nor can each formula asking for it, back to the first. ``origin`` is the cell
    asked for again until the cells back to it, the ones on the loop, are held (see Calculator._hold). A
    BaseException, as _Deeper is."""

    def __init__(self, origin: CellKey | None) -> None:
        super().__init__(origin)
        self.origin = origin


@dataclass(slots=True)
class Compiled:
    """One cell's formula, parsed once; no node for one the parser cannot read, which keeps the value it has."""

    node: Node | None
    precedents: list[Area] = field(default_factory=lambda: [])
    volatile: bool = False
    #: The formula calls a function Excel does not have, a VBA Function maybe (see _udf).
    calls_user: bool = False
    #: Volatility declared by a VBA UDF while this formula runs; reads still use the cache.
    udf_volatile: bool = False
    #: New UDF formulas get their initial value immediately, then join the next calculation.
    udf_pending: bool = False
    visibility_sensitive: bool = False
    #: Why the parser could not read the formula.
    unread: str = ""
    #: A dynamic-array formula, worked out whole and spilled (see _spills).
    dynamic: bool = False
    #: The data table the cell is one of, which works it out by trying values in its input cells (see _tried); its
    #: precedents are the formula it works out and the cells holding the values it tries.
    table: DataTable | None = None


class Calculator:
    """A workbook's formulas, and what they come to."""

    def __init__(self, book: Workbook) -> None:
        self.book = book
        #: The workbook as the formula engine reads it.
        self.engine_book = EngineBook(self)
        self.compiled: dict[CellKey, Compiled] = {}
        #: The formulas that call a function Excel does not have, which are worked out as they are written.
        self._user_calls: set[CellKey] = set()
        self._volatile_udfs: set[CellKey] = set()
        self._volatile_formulas: set[CellKey] = set()
        self._visibility_formulas: set[CellKey] = set()
        #: Which formulas read each cell a formula names on its own, and the blocks of cells each other formula reads,
        #: so that a change finds what reads it without going through every formula.
        self._readers: dict[CellKey, set[CellKey]] = {}
        self._blocks: dict[CellKey, list[Area]] = {}
        #: Cells that could not be worked out because they feed
        #: themselves.  Excel shows zero in these and warns separately.
        self.circular: set[CellKey] = set()
        #: The formulas held in loops, True for those on one and False for those reading one, found as formulas
        #: are worked out; and the cells that closed a loop, latest last (see _hold).
        self._loop: dict[CellKey, bool] = {}
        self._closed: dict[CellKey, None] = {}
        #: Loop cells being worked out from what the loop holds, and so not let go of when they come to a value.
        self._marked: set[CellKey] = set()
        #: Whether loops are being worked out over and over now (see _iterate).
        self._iterating = 0
        self._built = False
        #: The cells being worked out, one inside another, and how many.
        self._working: set[CellKey] = set()
        self._depth = 0
        self._calculation_depth = 0
        self._learning_depth = 0
        self._scopes: list[tuple[str | None, tuple[Area, ...] | None]] = []
        self._calculated_sheets: list[set[Worksheet]] = []
        self.pending_notifications: set[Worksheet] = set()
        self._held_caches: set[CellKey] = set()
        self.automatic_pending = False
        self._entries: set[CellKey] = set()
        self._active_formulas: list[CellKey] = []
        self._udf_waiting: dict[CellKey, CellKey] = {}
        self._calculation_order: dict[CellKey, None] = {}
        self._full_calculation_depth = 0
        self._udf_depth = 0
        self._provisional_disabled = 0
        self._provisional_errors: list[bool] = []
        self._cached_inputs: set[CellKey] = set()
        #: The cells put off until a cell too deep to work out inside them is worked out.
        self._waiting: set[CellKey] = set()
        #: How many calculations a macro asked for are running, which alone work data tables out in semi-automatic
        #: calculation, and the formulas being worked out with a value tried in a data table's input cell, which are
        #: put back after it (see _tried).
        self._asked = 0
        self._trying: list[set[CellKey]] = []
        #: The data tables that kept what they held while semi-automatic calculation waited, so that what read them
        #: then is worked out again with them by the next calculation that works them out.
        self._waited: set[CellKey] = set()

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
        if not cell.dynamic:
            self._unspill(sheet, row, column)
        if not cell.formula:
            self._keep(key, None)
            return
        tabled = self._sheet_named(sheet)
        if tabled is not None and tabled.data_tables:
            from pyopenvba.apps.excel._data_tables import table_at

            table = table_at(tabled, row, column)
            if table is not None:
                worked_out, tried = table.trial(row, column)
                self._keep(key, Compiled(node=None, table=table, precedents=[
                    Area(*one, *one, tabled.name) for one in [worked_out, *(value for _, value in tried)]]))
                return
        formula = cell.formula
        try:
            compiled = deep(lambda: self._compiled(formula, sheet, row, column))
        except FormulaSyntaxError as failure:
            # One formula the parser cannot read stops no other from being worked out.
            compiled = Compiled(node=None, unread=str(failure))
        compiled.dynamic = cell.dynamic
        self._keep(key, compiled)
        owner = self._sheet_named(sheet)
        if cell.dynamic and owner is not None and (row, column) not in owner.spills:
            # Known to the sheet before it is first worked out, so that reading where it may spill works it out.
            from pyopenvba.apps.excel._spills import Spill

            owner.spills[(row, column)] = Spill(Area(row, column, row, column), blocked=True)

    def settle_spills(self) -> None:
        """Work out every stale dynamic-array formula, so what each spills is in place: Excel spills after each
        edit."""
        self.build()
        for owner in self.book.sheets_:
            if owner.spills:
                self._settle(owner)

    def _settle(self, owner: Worksheet) -> None:
        """Work out the stale dynamic-array formulas of a sheet, so what each spills is in place."""
        if not self.automatic and not self._calculation_depth:
            return
        name = owner.name.lower()
        for row, column in list(owner.spills):
            cell = owner.cells_.get((row, column))
            key = (name, row, column)
            if cell is not None and cell.formula and cell.stale and key not in self._working \
                    and key not in self._waiting:
                self.value_of(owner.name, row, column)

    def _unspill(self, sheet: str, row: int, column: int) -> None:
        """A cell that holds no dynamic-array formula now: take back what its old one spilled."""
        owner = self._sheet_named(sheet)
        if owner is None or (row, column) not in owner.spills:
            return
        from pyopenvba.apps.excel._spills import remove

        self.spilled_changed(sheet, remove(owner, (row, column)))

    def spilled_changed(self, sheet: str, positions: set[tuple[int, int]]) -> None:
        """Cells a spill changed: mark whatever reads them as needing another look."""
        if positions:
            name = sheet.lower()
            self._spoil({(name, row, column) for row, column in positions}, set())

    def _compiled(self, formula: str, sheet: str, row: int, column: int) -> Compiled:
        node = self.engine_book.read(formula)
        precedents = self.engine_book.precedents(node, sheet, row, column)
        user = calls_user(node)
        return Compiled(node=node, precedents=precedents, volatile=volatile(node), calls_user=user,
                        visibility_sensitive=self.engine_book.visibility_sensitive(node, sheet))

    def forget(self, sheet: str, row: int, column: int) -> None:
        """A cell is gone: forget its formula, and take back what it spilled."""
        self._unspill(sheet, row, column)
        self._keep((sheet.lower(), row, column), None)

    def _keep(self, key: CellKey, compiled: Compiled | None) -> None:
        """Keep ``compiled`` as the formula at ``key``, or none, and note which cells it reads."""
        old = self.compiled.pop(key, None)
        self._loop.pop(key, None)
        self._closed.pop(key, None)
        self._user_calls.discard(key)
        self._volatile_udfs.discard(key)
        self._volatile_formulas.discard(key)
        self._visibility_formulas.discard(key)
        self._entries.discard(key)
        self._held_caches.discard(key)
        self._calculation_order.pop(key, None)
        if old is not None:
            for area in old.precedents:
                readers = self._readers.get((area.sheet.lower(), area.top, area.left))
                if readers is not None:
                    readers.discard(key)
            self._blocks.pop(key, None)
        if compiled is None:
            return
        self.compiled[key] = compiled
        if compiled.volatile:
            self._volatile_formulas.add(key)
        if compiled.visibility_sensitive:
            self._visibility_formulas.add(key)
        if compiled.calls_user:
            self._user_calls.add(key)
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

    def rebuild(self, *, keep_pending: bool = False) -> None:
        """Forget every parsed formula and mark them all for another look.

        Deleting a row moves what every formula was pointing at, and
        working out which ones moved is the same work as reading them
        all again.
        """
        from pyopenvba.apps.excel._spills import reconcile

        pending = {key: compiled.udf_pending for key, compiled in self.compiled.items()} if keep_pending else {}
        order = self._calculation_order.copy() if keep_pending else {}

        for owner in self.book.sheets_:
            if owner.spills or any(cell.spilled_from is not None for cell in owner.cells_.values()):
                reconcile(owner)
        self.compiled.clear()
        self._user_calls.clear()
        self._volatile_udfs.clear()
        self._volatile_formulas.clear()
        self._visibility_formulas.clear()
        self._entries.clear()
        self._held_caches.clear()
        self._calculation_order.clear()
        self._readers.clear()
        self._blocks.clear()
        self.circular.clear()
        self._loop.clear()
        self._closed.clear()
        self._built = False
        for sheet in self.book.sheets_:
            for cell in sheet.cells_.values():
                if cell.formula:
                    cell.stale = True
                    cell.save_dirty = None
        if keep_pending:
            self.build()
            for key, compiled in self.compiled.items():
                compiled.udf_pending = pending.get(key, False)
            self._calculation_order.update((key, None) for key in order if key in self.compiled)

    def wrote(self, sheet: str, row: int, column: int) -> None:
        """A cell changed: mark whatever reads it as needing another look, and a dynamic-array formula whose answer
        wants the cell, which may now spill there or be kept out."""
        self.build()
        changed = {(sheet.lower(), row, column)}
        owner = self._sheet_named(sheet)
        if owner is not None and not owner.enable_calculation:
            written = owner.cells_.get((row, column))
            if written is not None and written.formula:
                # A formula written where nothing is worked out reads 0, whatever the cell held (over_constant).
                written.value = 0.0
        if owner is not None and owner.spills:
            from pyopenvba.apps.excel._spills import block_cached, wanted

            for anchor in wanted(owner, row, column):
                cell = owner.cells_.get(anchor)
                if cell is not None:
                    cell.stale = True
                    changed.add((sheet.lower(), *anchor))
                    changed.update(self._members((sheet.lower(), *anchor)))
                    if not self.automatic:
                        changed.update((sheet.lower(), *position) for position in block_cached(owner, anchor))
        self._spoil(changed, set())

    def wrote_area(self, sheet: str, area: Area) -> None:
        """A block of cells changed at once, as a copy changes it: forget the formulas compiled there, and mark
        whatever reads the block, and what reads that, as needing another look."""
        self.build()
        name = sheet.lower()
        for key in [key for key in self.compiled if key[0] == name and area.contains(key[1], key[2])]:
            self._keep(key, None)
        reading: set[CellKey] = set()
        owner = self._sheet_named(sheet)
        if owner is not None:
            for (row, column), cell in list(owner.cells_.items()):
                if area.contains(row, column) and cell.dynamic and cell.formula:
                    self.remember(sheet, row, column, cell)
        if owner is not None and owner.spills:
            from pyopenvba.apps.excel._spills import reconcile

            reconcile(owner)
            # A bulk write can obstruct an existing spill just as a single-cell write can.
            # Recalculate its anchor before newly copied formulas try to spill over its old output.
            for anchor, spill in owner.spills.items():
                if (area.top <= spill.area.bottom and spill.area.top <= area.bottom
                        and area.left <= spill.area.right and spill.area.left <= area.right):
                    cell = owner.cells_.get(anchor)
                    if cell is not None:
                        cell.stale = True
                        key = (name, *anchor)
                        reading.add(key)
                        reading.update(self._members(key))
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

    def names_changed(self, changed: set[str]) -> None:
        """Defined names added, redefined, renamed or deleted, by their names in lower case: each formula naming
        one, or naming a name that names one, read again and worked out again with what reads it; every other
        formula keeps what it holds (tests/fixtures/calculation_state.json, name_*)."""
        from pyopenvba._a1 import split_sheet
        from pyopenvba.formula._calc.nodes import Call, NameReference, walk

        self.build()
        reached = set(changed)
        growing = True
        while growing:
            growing = False
            for entry in self.book.names_.entries:
                bare = split_sheet(entry.name)[1].lower()
                if bare not in reached and any(word.lower() in reached for word in _WORDS.findall(entry.refers_to)):
                    reached.add(bare)
                    growing = True
        marked: set[CellKey] = set()
        for key, compiled in list(self.compiled.items()):
            if compiled.node is None or not any(
                    isinstance(one, NameReference) and one.name.lower() in reached
                    or isinstance(one, Call) and one.name.lower() in reached for one in walk(compiled.node)):
                continue
            cell = self._cell(key)
            owner = self._sheet_named(key[0])
            if cell is not None and owner is not None:
                self.remember(owner.name, key[1], key[2], cell)
                cell.stale = True
                marked.add(key)
        self._spoil(marked, set(marked))

    def dirtied(self, sheet: str, area: Area) -> None:
        """Range.Dirty: the formulas in ``area``, what reads a cell of it, and what reads those, to work out again,
        the cells keeping what they hold until then (tests/fixtures/calculation_state.json)."""
        self.build()
        name = sheet.lower()
        marked: set[CellKey] = set()
        for key, compiled in self.compiled.items():
            if key[0] == name and area.contains(key[1], key[2]) or any(
                    one.sheet.lower() == name and one.top <= area.bottom and area.top <= one.bottom
                    and one.left <= area.right and area.left <= one.right for one in compiled.precedents):
                cell = self._cell(key)
                if cell is not None:
                    cell.stale = True
                    marked.add(key)
        self._spoil(marked, set(marked))

    def _spoil(self, changed: set[CellKey], seen: set[CellKey]) -> None:
        """Mark stale what reads ``changed``, and what reads that, a wave at a time."""
        while changed:
            following: set[CellKey] = set()
            for name, row, column in changed:
                for key in self._reading(name, row, column) - seen:
                    cell = self._cell(key)
                    if cell is not None:
                        cell.save_dirty = None
                    if cell is not None and not cell.stale:
                        cell.stale = True
                        following.add(key)
                        # What reads a cell of an array formula follows the array.
                        following.update(self._members(key))
                    seen.add(key)
            changed = following

    def thawed(self, sheet: Worksheet) -> None:
        """EnableCalculation True again: every formula on the sheet and what reads it to work out, at once under
        automatic calculation, at the next calculation under manual (calc_back_on, manual_back_on)."""
        self.build()
        keys = {(sheet.name.lower(), row, column) for (row, column), cell in sheet.cells_.items() if cell.formula}
        for _name, row, column in keys:
            sheet.cells_[(row, column)].stale = True
        self._spoil(set(keys), set(keys))
        if self.automatic:
            for _name, row, column in sorted(keys):
                self.value_of(sheet.name, row, column)

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

    def calculate_all(self, only: Worksheet | None = None, *, volatile: bool = True,
                      settled: set[CellKey] | None = None, udf_volatile: bool | None = None,
                      full: bool = False, background: bool = False,
                      areas: tuple[Area, ...] | None = None, force: bool = False,
                      asked: bool = False) -> list[Worksheet]:
        """Work out stale and volatile formulas, even in manual mode; report sheets actually calculated.

        only limits the work to one sheet, as Worksheet.Calculate does;
        volatile False leaves volatile formulas be, which is how pending
        formulas are settled without an edit (see Excel object vba_sink_attached).
        asked for Application.Calculate's and Worksheet.Calculate's, which
        work data tables out in semi-automatic calculation, as Range.Calculate's
        does not (tests/fixtures/data_tables.json).
        """
        self.build()
        if only is None:
            self.automatic_pending = False
        if not self.compiled:
            return []
        if self._waited and not (self.book.application.calculation == 2 and not (asked or self._asked)):
            waited, self._waited = self._waited, set()
            self._spoil(waited, set(waited))
        volatile_users = volatile if udf_volatile is None else udf_volatile
        roots: set[CellKey] = set()
        for key, compiled in self.compiled.items():
            if only is not None and key[0] != only.name.lower():
                continue
            if areas is not None and not any(area.contains(key[1], key[2]) for area in areas):
                continue
            if (force or full or compiled.udf_pending or (volatile and compiled.volatile)
                    or (volatile_users and compiled.udf_volatile)) and key not in (settled or ()):
                cell = self._cell(key)
                if cell is not None:
                    cell.stale = True
                    roots.add(key)
        self._spoil(roots, set(roots))
        # With Application.Iteration on, the formulas held in loops are left to the iteration after the rest.
        iteration = self.book.application.iteration
        held = set(self._loop) if iteration else set[CellKey]()
        worked: set[Worksheet] = set()
        self._scopes.append((None if only is None else only.name.lower(), areas))
        self._calculated_sheets.append(worked)
        self._calculation_depth += 1
        self._full_calculation_depth += int(full)
        self._asked += int(asked)
        self._provisional_errors.append(False)
        try:
            owners = {sheet.name.lower(): sheet for sheet in self.book.sheets_ if only is None or sheet is only}
            # One chain spans the workbook, so a deferred UDF can wait for a
            # precedent on a later tab rather than retrying on its own tab.
            new = [(name, row, column) for name, sheet in owners.items()
                   for (row, column), cell in reversed(list(sheet.cells_.items()))
                   if cell.formula and (name, row, column) not in self._calculation_order]
            known = [key for key in self._calculation_order if key[0] in owners]
            queue = deque(new + known)
            deferred: set[CellKey] = set()
            while queue:
                key = queue.popleft()
                name, row, column = key
                sheet = owners[name]
                cell = sheet.cells_.get((row, column))
                if cell is None or not cell.formula:
                    continue
                if areas is not None and not any(area.contains(row, column) for area in areas):
                    continue
                if key in held:
                    continue
                if cell.stale:
                    try:
                        self.value_of(sheet.name, row, column, force=True)
                    except _UDFWaiting:
                        if key not in deferred:
                            deferred.add(key)
                            queue.append(key)
                            continue
                        # No progress through a deferred cycle: use the
                        # ordinary reference path and circular detection.
                        self._provisional_disabled += 1
                        try:
                            self.value_of(sheet.name, row, column, force=True)
                        finally:
                            self._provisional_disabled -= 1
                    except VBAUnsupportedError:
                        compiled = self.compiled.get(key)
                        if not background or (compiled is not None and self._user_unsupported(compiled)):
                            raise
                        continue
                    deferred.clear()
            if iteration:
                self._iterate(frozenset(
                    key for key in self._loop if key[0] in owners
                    and (areas is None or any(area.contains(key[1], key[2]) for area in areas))))
        finally:
            self._scopes.pop()
            self._calculated_sheets.pop()
            self._calculation_depth -= 1
            self._full_calculation_depth -= int(full)
            self._asked -= int(asked)
            if self._provisional_errors.pop():
                interpreter = self.book.application.interpreter
                if interpreter is not None:
                    interpreter.err.number = 40040
                    interpreter.err.description = 'Application-defined or object-defined error'
                if self._provisional_errors:
                    self._provisional_errors[-1] = True
            if only is not None:
                self._held_caches.update(key for key in self.compiled
                                         if (cell := self._cell(key)) is not None and cell.stale)
        return [sheet for sheet in self.book.sheets_ if sheet in worked]

    @property
    def has_volatile_udfs(self) -> bool:
        return bool(self._volatile_udfs)

    @property
    def has_volatile_formulas(self) -> bool:
        return bool(self._volatile_formulas or self._volatile_udfs)

    @property
    def has_pending_udfs(self) -> bool:
        return any(self.compiled[key].udf_pending for key in self._user_calls)

    @property
    def has_held_caches(self) -> bool:
        return bool(self._held_caches)

    def settle_automatic_cache(self) -> None:
        """Materialize automatic work deferred while nobody observed its events."""
        if self.automatic and self.automatic_pending:
            self.calculate_all(volatile=False, background=True)

    def udf_source_changed(self) -> None:
        """Keep cached results, but reconsider user functions at the next calculation."""
        self.build()
        for key in self._user_calls:
            self.compiled[key].udf_pending = True

    def set_udf_volatile(self, sheet: str, row: int, column: int, enabled: bool) -> None:
        key = (sheet.lower(), row, column)
        compiled = self.compiled.get(key)
        if compiled is not None:
            compiled.udf_volatile = enabled
            if enabled:
                self._volatile_udfs.add(key)
            else:
                self._volatile_udfs.discard(key)

    def settle_entries(self) -> set[CellKey]:
        """Give explicit formula assignments their initial result, including in manual mode, as Excel works a
        formula out when it is written: what a loop's cells hold when a later formula closes it is what they came
        to then, and a formula closing one is found then (tests/fixtures/iteration.json).

        Copied caches are not entries. Existing stale UDFs also settle here in
        automatic mode, before the edit's events are dispatched.
        """
        worked: set[CellKey] = set()
        entries, self._entries = self._entries, set()
        candidates = self._user_calls | entries if self.automatic else entries
        for name, row, column in sorted(candidates):
            cell = self._cell((name, row, column))
            owner = self._sheet_named(name)
            compiled = self.compiled.get((name, row, column))
            if cell is not None and cell.stale and owner is not None and compiled is not None \
                    and (self.automatic or (name, row, column) in entries) \
                    and (compiled.calls_user or not self.automatic or (name, row, column) in entries):
                try:
                    self._learning_depth += int((name, row, column) not in entries)
                    try:
                        self.value_of(owner.name, row, column, force=True)
                    finally:
                        self._learning_depth -= int((name, row, column) not in entries)
                except VBAUnsupportedError:
                    if self._user_unsupported(compiled):
                        raise
                    # Preserve valid formula text the engine cannot yet evaluate.
                    continue
                if self.automatic and not compiled.calls_user:
                    # What it came to as it was written stays for a loop a later formula closes; a read works it out
                    # again, as the model works formulas out when they are read.
                    cell.stale = True
                    continue
                worked.add((name, row, column))
                if not self.automatic:
                    # Initial manual values are available immediately, but
                    # still participate in the next requested calculation.
                    cell.stale = True
        return worked

    def _user_unsupported(self, compiled: Compiled) -> bool:
        if not compiled.calls_user:
            return False
        from pyopenvba.apps.excel._udf import has_user_function

        return has_user_function(self.book, compiled.node)

    def entered(self, sheet: str, row: int, column: int) -> None:
        """An explicit formula assignment needs its initial result even in manual mode."""
        key = (sheet.lower(), row, column)
        self._entries.add(key)
        compiled = self.compiled.get(key)
        if compiled is not None and compiled.calls_user and not self.automatic:
            compiled.udf_pending = True

    @contextmanager
    def user_call(self) -> Generator[None, None, None]:
        self._udf_depth += 1
        try:
            yield
        finally:
            self._udf_depth -= 1

    def after_udf(self, *, error: bool = False) -> None:
        if self._active_formulas:
            pending = self._udf_waiting.pop(self._active_formulas[-1], None)
            if pending is not None:
                if error and self._provisional_errors:
                    self._provisional_errors[-1] = True
                raise _UDFWaiting(*pending)

    def invalidate_visibility(self) -> None:
        """Row visibility is an input to subtotal formulas even when no cell value changed."""
        self.build()
        roots = self._visibility_formulas.copy()
        for key in self._visibility_formulas:
            cell = self._cell(key)
            if cell is not None:
                cell.stale = True
                roots.update(self._members(key))
        self._spoil(roots, set(roots))

    def recalculated(self, *, settled: set[CellKey] | None = None) -> list[Worksheet]:
        """Calculate after an edit, unless the application is in manual mode."""
        return self.calculate_all(settled=settled, background=True) if self.automatic else []

    # --- loops of formulas -------------------------------------------------------------

    def loop_cells(self) -> frozenset[CellKey]:
        """The formulas held in loops. Excel finds a loop as it works formulas out, among the cells a formula reads
        then: ROWS(A1:A10) in A5 reads no cell and closes none, nor does a MATCH there that stops at A3, where
        SUM(A1:A10) does (tests/fixtures/self_reference.json). With Application.Iteration off none of them is
        worked out: a formula just written there is 0, and the others keep what they hold until the loop is
        broken. With it on every calculation works them out over and over (see _iterate)."""
        return frozenset(self._loop)

    @property
    def iterates(self) -> bool:
        """Whether every calculation works loops out, as a volatile formula is worked out every time."""
        return self.book.application.iteration and bool(self._loop)

    def _reads_loops(self) -> bool:
        """Whether a formula reaching a loop reads what its cells hold and is worked out, as in an iteration and with
        Application.Iteration on in manual mode, where Excel works a formula out once as it is written."""
        return bool(self._iterating) or (self.book.application.iteration and not self.automatic)

    def circular_cell(self, sheet: Worksheet) -> tuple[int, int] | None:
        """Worksheet.CircularReference's cell: of the cells on loops on the sheet, the latest that closed one, or
        else the first, row by row; none with Application.Iteration on. Of two cells of one loop on a sheet Excel
        names either, run to run (tests/fixtures/iteration.json)."""
        if self.book.application.iteration:
            return None
        name = sheet.name.lower()
        for key in [key for key, on in self._loop.items() if on and key[0] == name]:
            cell = self._cell(key)
            if cell is not None and cell.stale:
                # Still going round?
                self.value_of(sheet.name, key[1], key[2])
        here = [key for key, on in self._loop.items() if on and key[0] == name]
        if not here:
            return None
        found = next((key for key in reversed(self._closed) if key in here), min(here))
        return found[1], found[2]

    def _held_value(self, cell: Cell) -> object:
        """A cell held in a loop, asked for: what it holds, to a macro and to a formula reading the loop; any other
        formula asking for it cannot be worked out either (see _hold)."""
        if self._active_formulas and not self._udf_depth:
            if not self._reads_loops():
                raise _Loop(None)
            asking = self._active_formulas[-1]
            self._loop.setdefault(asking, False)
            self._marked.add(asking)
        return 0.0 if cell.value is EMPTY else cell.value

    def _closing(self, key: CellKey, cell: Cell) -> object:
        """A formula asked for ``key``, which is being worked out under it: a loop. Read where formulas read loops,
        the cells from ``key`` up to the one asking are on it and are worked out from what it holds, a formula just
        written starting at 0; anywhere else none of them can be worked out (see _hold), a macro a formula runs and
        reads its own cell stopping where it reads it."""
        self.circular.add(key)
        if self._udf_depth or not self._reads_loops():
            raise _Loop(key)
        stack = self._active_formulas
        for one in stack[stack.index(key):] if key in stack else stack:
            self._loop[one] = True
            self._marked.add(one)
        self._closed.pop(key, None)
        self._closed[key] = None
        return 0.0 if cell.value is EMPTY else cell.value

    def _hold(self, key: CellKey, cell: Cell, loop: _Loop) -> object:
        """A formula that cannot be worked out, as it is on a loop or reads a cell held in one: it keeps what it holds,
        a formula just written 0, and so does each formula asking for it, back to the first. The cells from the one
        asked for again up to the one asking are on the loop and the one asked for closed it; those asking for them
        only read it (tests/fixtures/iteration.json)."""
        if cell.value is EMPTY:
            cell.value = 0.0
        self._loop[key] = loop.origin is not None or self._loop.get(key, False)
        self._marked.discard(key)
        cell.stale = False
        self.circular.add(key)
        if key == loop.origin:
            self._closed.pop(key, None)
            self._closed[key] = None
            loop.origin = None
        if self._active_formulas and not self._udf_depth:
            raise loop
        return cell.value

    def _iterate(self, keys: frozenset[CellKey]) -> None:
        """Work loops out over and over, as Excel does with Application.Iteration on. A pass works each formula held
        in one out once, a sheet at a time in the order the sheets were made, whatever their tabs' order, and row by
        row, each reading what the others hold by then, a formula just written starting at 0. Passes go on until no
        number changes by MaxChange or more, text and logicals changing nothing, or MaxIterations of them are done
        (tests/fixtures/iteration.json)."""
        if not keys:
            return
        application = self.book.application
        made = {sheet.name.lower(): sheet.sheet_id for sheet in self.book.sheets_}
        cells: list[tuple[CellKey, Cell, Compiled]] = []
        for key in sorted(keys, key=lambda key: (made.get(key[0], 0), key[1], key[2])):
            cell = self._cell(key)
            compiled = self.compiled.get(key)
            if cell is not None and compiled is not None:
                if cell.value is EMPTY:
                    cell.value = 0.0
                cells.append((key, cell, compiled))
        for _ in range(application.max_iterations):
            change = 0.0
            for key, cell, compiled in cells:
                before = cell.value
                after = self._step(key, cell, compiled)
                if isinstance(before, float | int) and isinstance(after, float | int) \
                        and not isinstance(before, bool) and not isinstance(after, bool):
                    change = max(change, abs(after - before))
            if change < application.max_change:
                break
        from pyopenvba.apps.excel._controls import cell_changed

        for key, cell, _ in cells:
            owner = self._sheet_named(key[0])
            if owner is not None:
                cell_changed(owner, key[1], key[2], cell.value)

    def _step(self, key: CellKey, cell: Cell, compiled: Compiled) -> object:
        """A loop's formula worked out once, reading what the loop's cells hold."""
        owner = self._sheet_named(key[0])
        if owner is None:
            return cell.value
        self._iterating += 1
        self._working.add(key)
        self._active_formulas.append(key)
        try:
            value = self._computed(compiled, owner.name, key[1], key[2])
        finally:
            self._iterating -= 1
            self._working.discard(key)
            self._active_formulas.pop()
        cell.value = value
        cell.stale = False
        for sheets in self._calculated_sheets:
            sheets.add(owner)
        return value

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
                except _UDFWaiting:
                    raise
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

    def is_working(self, sheet: str, row: int, column: int) -> bool:
        """Whether this exact worksheet formula is currently being evaluated."""
        return (sheet.lower(), row, column) in self._working

    def _value(self, sheet: str, row: int, column: int, force: bool) -> object:
        self.build()
        key = (sheet.lower(), row, column)
        owner = self._sheet_named(sheet)
        if self._scopes:
            scope, areas = self._scopes[-1]
            outside = scope is not None and scope != key[0]
            outside |= areas is not None and not any(area.contains(row, column) for area in areas)
            if outside:
                cell = self._cell(key)
                if cell is not None and cell.stale and self._active_formulas:
                    caller = self._active_formulas[-1]
                    # A sheet-only UDF call settles against its foreign
                    # cached argument; ordinary expressions stay dirty.
                    if areas is not None or not self.compiled[caller].calls_user:
                        self._cached_inputs.add(caller)
                return EMPTY if cell is None else cell.value
        if owner is not None and not owner.enable_calculation:
            # A sheet whose EnableCalculation is False keeps every formula as it is, through a full calculation too,
            # and a sheet reading one reads what it keeps (tests/fixtures/worksheet_settings.json, calc_*).
            frozen = self._cell(key)
            return EMPTY if frozen is None else frozen.value
        if owner is not None and owner.array_formulas:
            found = array_at(owner, row, column)
            if found is not None and found[0] != (row, column):
                self._value(sheet, *found[0], force)
                member = owner.cells_.get((row, column))
                return EMPTY if member is None else member.value
        cell = self._cell(key)
        if (cell is None or not cell.formula) and owner is not None and owner.spills:
            # A dynamic-array formula to work out first may spill here, or have spilled here and shrunk.
            self._settle(owner)
            cell = self._cell(key)
        if cell is None:
            return EMPTY
        if not cell.formula:
            return cell.value
        compiled = self.compiled.get(key)
        if compiled is None:
            self.remember(sheet, row, column, cell)
            compiled = self.compiled[key]
        if key in self._loop and (self._iterating or not (cell.stale or force)):
            return self._held_value(cell)
        if self._calculation_depth and self._udf_depth and not self._provisional_disabled \
                and self._active_formulas and cell.stale and compiled.calls_user \
                and (compiled.udf_pending or self._full_calculation_depth) \
                and key not in self._working and key not in self._waiting:
            caller = self._active_formulas[-1]
            calling = self.compiled[caller]
            if calling.calls_user and (calling.udf_pending or self._full_calculation_depth):
                self._udf_waiting[caller] = key
                return EMPTY
        if not (cell.stale or force):
            return cell.value
        if compiled.table is not None and self.book.application.calculation == 2 and not self._asked:
            # In semi-automatic calculation a data table keeps what it holds until a calculation asked for
            # (tests/fixtures/data_tables.json).
            self._waited.add(key)
            return cell.value
        if key in self._held_caches and not force and not self._calculation_depth:
            return cell.value
        if key in self._working or key in self._waiting:
            # A loop, found before manual mode's cache stands in for the cell: Excel finds it as the formula is
            # written then too (tests/fixtures/iteration.json).
            return self._closing(key, cell)
        if not self.automatic and not force and not self._calculation_depth:
            if self._active_formulas:
                self._cached_inputs.add(self._active_formulas[-1])
            return cell.value
        if self._depth >= _DEEPEST:
            raise _Deeper(sheet, row, column)
        self._working.add(key)
        self._depth += 1
        self._active_formulas.append(key)
        self._cached_inputs.discard(key)
        cached = False
        caught: _Loop | None = None
        value: object = EMPTY
        try:
            value = self._computed(compiled, sheet, row, column)
        except _Loop as loop:
            caught = loop
        finally:
            cached = key in self._cached_inputs
            self._cached_inputs.discard(key)
            self._depth -= 1
            self._working.discard(key)
            self._active_formulas.pop()
            self._udf_waiting.pop(key, None)
        if caught is not None:
            return self._hold(key, cell, caught)
        if key in self._marked:
            # Worked out from what a loop holds, it stays on the loop.
            self._marked.discard(key)
        else:
            self._loop.pop(key, None)
        cell.value = value
        cell.save_dirty = None
        self._held_caches.discard(key)
        cell.stale = cached
        if owner is not None:
            for sheets in self._calculated_sheets:
                sheets.add(owner)
        if self._calculation_depth:
            compiled.udf_pending = False
        if (self._calculation_depth or self._learning_depth) and (not self._scopes or self._scopes[-1][1] is None):
            self._calculation_order.setdefault(key, None)
        if any(key in reading for reading in self._trying):
            # A value worked out with a value tried in a data table's input cell is put back, unseen (see _tried).
            return value
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
        if compiled.table is not None:
            return self._tried(compiled.table, sheet, row, column)
        if compiled.node is None:
            raise VBAUnsupportedError(f"working out a formula the model cannot read is not implemented: "
                                      f"{compiled.unread}")
        owner = self._worksheet(sheet)
        block = owner.array_formulas.get((row, column))
        now = self.now()
        node = compiled.node
        # A dynamic-array formula is worked out whole, as an array formula is, and spills.
        whole_formula = block is not None or compiled.dynamic

        def context() -> EngineContext:
            return EngineContext(self.engine_book, owner.name, row, column, array=whole_formula, today=now.date(),
                                 now=now)

        def one() -> Scalar:
            working = context()
            return working.first(working.formula(node))

        def whole() -> Array:
            working = context()
            return working.array_of(working.formula(node))

        from pyopenvba.apps.excel._udf import evaluation_origin

        try:
            # The origin surrounds actual evaluation, including deep scalar/array retries.
            with evaluation_origin(self.book.application, "worksheet"):
                if compiled.dynamic:
                    return self._spilling(owner, row, column, deep(whole))
                if block is None:
                    return model_value(deep(one))
                return _spread(owner, block, deep(whole))
        except EngineError as failure:
            if compiled.dynamic:
                return self._spilling(owner, row, column, Array([[failure.error]]))
            return model_value(failure.error) if block is None else _spread(owner, block, Array([[failure.error]]))
        except RecursionError:
            self.circular.add((sheet.lower(), row, column))
            return 0.0

    def _tried(self, table: DataTable, sheet: str, row: int, column: int) -> object:
        """A data table's cell: the formula it works out, worked out with the values it tries put in the table's input
        cells and every formula reading them, on any sheet, worked out again with them; then each of those cells is
        as it was, and the input cells hold what they held, a formula included. A blank formula comes to 0, and an
        input cell taken away to #REF! (tests/fixtures/data_tables.json)."""
        if table.broken:
            return ExcelError("#REF!")
        from pyopenvba.apps.excel._model import Cell

        owner = self._worksheet(sheet)
        worked_out, tried = table.trial(row, column)
        # Each value tried is what its cell comes to before any is tried.
        values = [(at, self.value_of(owner.name, *source)) for at, source in tried]
        name = owner.name.lower()
        inputs = {(name, *at) for at, _ in values}
        reading = self._closure(inputs) - inputs
        if any((one := self.compiled.get(key)) is not None and one.dynamic for key in reading):
            raise VBAUnsupportedError("a data table whose input cell a dynamic-array formula reads is not implemented")
        saved = {key: None if (cell := self._cell(key)) is None else (cell.value, cell.stale, cell.save_dirty)
                 for key in reading}
        pending = {key for key in reading if (one := self.compiled.get(key)) is not None and one.udf_pending}
        held = self._held_caches & reading
        originals = {at: owner.cells_.get(at) for at, _ in values}
        for key in reading:
            cell = self._cell(key)
            if cell is not None and cell.formula:
                cell.stale = True
        self._trying.append(reading)
        # Worked out whatever the mode, sheet or range a calculation keeps to, from the top.
        self._scopes.append((None, None))
        self._calculation_depth += 1
        depth, self._depth = self._depth, 0
        try:
            for at, value in values:
                original = originals[at]
                owner.cells_[at] = Cell(value=value) if original is None else replace(
                    original, value=value, formula="", stale=False, shared=None, dynamic=False, spilled_from=None)
            answer = self.value_of(owner.name, *worked_out)
        finally:
            self._depth = depth
            self._calculation_depth -= 1
            self._scopes.pop()
            self._trying.pop()
            for at, original in originals.items():
                if original is None:
                    owner.cells_.pop(at, None)
                else:
                    owner.cells_[at] = original
            for key, state in saved.items():
                if state is None:
                    made = self._sheet_named(key[0])
                    if made is not None:
                        made.cells_.pop((key[1], key[2]), None)
                    continue
                cell = self._cell(key)
                if cell is not None:
                    cell.value, cell.stale, cell.save_dirty = state
            for key in pending:
                if key in self.compiled:
                    self.compiled[key].udf_pending = True
            self._held_caches |= held
        # The formula as it stands, so that what the table cell reads is worked out under it, and what changes it
        # next reaches the cell through it.
        depth, self._depth = self._depth, 0
        try:
            self.value_of(owner.name, *worked_out)
        finally:
            self._depth = depth
        return 0.0 if answer is EMPTY else answer

    def _closure(self, keys: set[CellKey]) -> set[CellKey]:
        """The formulas reading ``keys``, those reading them, and so on, with every cell of each array formula among
        them."""
        found: set[CellKey] = set()
        wave = set(keys)
        while wave:
            following: set[CellKey] = set()
            for name, row, column in wave:
                for key in self._reading(name, row, column) - found:
                    found.add(key)
                    following.add(key)
                    for member in self._members(key) - found:
                        found.add(member)
                        following.add(member)
            wave = following
        return found

    def _spilling(self, owner: Worksheet, row: int, column: int, answer: Array) -> object:
        """A dynamic-array formula's answer spilled from its cell; the value the cell keeps."""
        from pyopenvba.apps.excel._spills import place

        value, changed = place(owner, (row, column), answer)
        self.spilled_changed(owner.name, changed)
        return value

    def _members(self, key: CellKey) -> set[CellKey]:
        """The other cells of the array formula whose first cell is ``key``, or of the block its dynamic-array formula
        spilled into, if it is one."""
        owner = self._sheet_named(key[0])
        if owner is None:
            return set()
        block = owner.array_formulas.get((key[1], key[2]))
        spill = owner.spills.get((key[1], key[2]))
        if block is None and spill is not None and not spill.blocked:
            block = spill.area
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
        whatever hid it, inside the filter's range or not. A table's filter
        with criteria counts as the sheet's would, whatever FilterMode says
        (tests/fixtures/table_autofilter/).
        """
        from pyopenvba.apps.excel._table_filters import criteria_anywhere

        target = self._worksheet(sheet)
        dims = target.dims
        if dims.zero_height:
            raise VBAUnsupportedError("SUBTOTAL over a sheet whose rows are hidden by default is not implemented")
        if not every and not criteria_anywhere(target):
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


def as_vba(value: object, cell: Cell | None = None, *, epoch_1904: bool = False) -> object:
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
        return VBADate(float(value) + (1462 if epoch_1904 and not 0 <= value < 1 else 0))
    if is_currency_format(code):
        return VBACurrency(value)
    return float(value)


def as_python(value: object, cell: Cell | None = None, *, epoch_1904: bool = False) -> object:
    """A cell's value as ``as_vba`` reads it, but the Double where its Date or Currency cannot hold it."""
    try:
        return as_vba(value, cell, epoch_1904=epoch_1904)
    except VBARuntimeError:
        return as_vba(value)
