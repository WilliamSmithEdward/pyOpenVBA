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
"""

from __future__ import annotations

import datetime as _dt
from collections import deque
from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass, field
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


class _UDFWaiting(_Deeper):
    """A provisional UDF call ran before its pending UDF precedent was ready."""


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
                      areas: tuple[Area, ...] | None = None, force: bool = False) -> list[Worksheet]:
        """Work out stale and volatile formulas, even in manual mode; report sheets actually calculated.

        only limits the work to one sheet, as Worksheet.Calculate does;
        volatile False leaves volatile formulas be, which is how pending
        formulas are settled without an edit (see Excel object vba_sink_attached).
        """
        self.build()
        if only is None:
            self.automatic_pending = False
        if not self.compiled:
            return []
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
        worked: set[Worksheet] = set()
        self._scopes.append((None if only is None else only.name.lower(), areas))
        self._calculated_sheets.append(worked)
        self._calculation_depth += 1
        self._full_calculation_depth += int(full)
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
        finally:
            self._scopes.pop()
            self._calculated_sheets.pop()
            self._calculation_depth -= 1
            self._full_calculation_depth -= int(full)
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
        """Give explicit formula assignments their initial result, including in manual mode.

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
                    and (compiled.calls_user or not self.automatic):
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
        if key in self._held_caches and not force and not self._calculation_depth:
            return cell.value
        if not self.automatic and not force and not self._calculation_depth:
            if self._active_formulas:
                self._cached_inputs.add(self._active_formulas[-1])
            return cell.value
        if key in self._working or key in self._waiting:
            self.circular.add(key)
            raise ExcelError("#CIRCULAR!")
        if self._depth >= _DEEPEST:
            raise _Deeper(sheet, row, column)
        self._working.add(key)
        self._depth += 1
        self._active_formulas.append(key)
        self._cached_inputs.discard(key)
        cached = False
        try:
            value = self._computed(compiled, sheet, row, column)
        finally:
            cached = key in self._cached_inputs
            self._cached_inputs.discard(key)
            self._depth -= 1
            self._working.discard(key)
            self._active_formulas.pop()
            self._udf_waiting.pop(key, None)
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
        except ExcelError as failure:
            if failure.name != "#CIRCULAR!":
                raise
            # Excel leaves a zero in a cell that feeds itself.
            self.circular.add((sheet.lower(), row, column))
            return 0.0
        except RecursionError:
            self.circular.add((sheet.lower(), row, column))
            return 0.0

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
