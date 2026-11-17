"""The events Excel raises in a sheet's and the workbook's own modules.

Measured in live Excel (scripts/measure_events.py and measure_excel_event_sinks.py).
Each event reaches worksheet sinks, Worksheet_Change, workbook sinks,
Workbook_SheetChange and application sinks, in that order, and none while
Application.EnableEvents is False. The sheet is first among workbook and
application handler arguments.
A handler that edits a sheet raises that edit's events inside its own, as
Excel does.

* Change: a write of Value, Value2, Formula, FormulaR1C1 or FormulaArray,
  ClearContents and Clear, whatever the cells held before, with the range
  written as Target, several areas included; inserting and deleting rows,
  columns or cells, with the rows, columns or cells as Target; Copy, with
  the cells copied to; FillDown, with the whole range; AutoFill, with the
  cells filled, after it selects the destination; Replace, one cell at a
  time. Sort, Merge and formats raise none.
* Calculate: after an edit that worked formulas out, for each sheet
  whose formulas it worked out -- before the Change, except after an
  insert, which changes first, and after Replace, which calculates last.
  Disabled calculation notifications wait for an enabled calculation cycle.
  Manual range calculation retains its new notification for a later cycle.
* SelectionChange: a Select that moves the selection.
* Deactivate, then Activate: another sheet made active; a new sheet is
  NewSheet first, and deactivates the sheet it was placed beside, or the
  active one.

tests/fixtures/more_events.json measures the rest:

* BeforeDelete: for each worksheet of a delete Excel does not refuse, last tab
  first, before the first one named deactivates, active or not.
* NewChart: a chart made, or a copy by Duplicate; not one a sheet copy carries.
* WindowDeactivate, Deactivate, Activate, WindowActivate: another workbook's
  window to the front, a workbook closing among them, the left and the coming.
* WindowResize: once for a width or height set, twice when it changes; once for
  a state that changes.
* BeforePrint: before a print, its sheet active by then if it was not.
* FollowHyperlink: after the link has gone to its place.
* NewWorkbook: a workbook made, before its window comes to the front.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._runtime import Slot
from pyopenvba.interpreter._values import to_bool

if TYPE_CHECKING:
    from pyopenvba.apps.excel._chart_sheets import ChartSheet
    from pyopenvba.apps.excel._model import Application, Range, Workbook, Worksheet
    from pyopenvba.apps.excel._pivots import PivotTable


def _on(book: Workbook) -> bool:
    return bool(book.application.enable_events)


def _handle(owner: VBAObject, procedure: str, args: list[object]) -> None:
    """Run ``procedure`` of ``owner``'s own module, if it has one; a module without it is simply not listening."""
    from pyopenvba.apps.excel._model import Workbook, Worksheet
    from pyopenvba.apps.excel._vbide import VBProject

    book = owner if isinstance(owner, Workbook) else owner.book if isinstance(owner, Worksheet) else None
    if book is not None and isinstance(owner, (Workbook, Worksheet)) and isinstance(book.vbide_project, VBProject):
        entry = next((entry for entry in book.vbide_project.components.entries
                      if entry.name.casefold() == owner.code_name.casefold()), None)
        if entry is not None and entry.pending:
            old = owner.vba_document
            has_handler = old is not None and old.module.procedure(procedure) is not None
            has_handler |= any(row[0].casefold() == procedure.casefold() for row in entry.code_module.procedures())
            if has_handler:
                book.vbide_project.compile_pending()
    document = owner.vba_document
    if document is None:
        return
    found = document.module.procedure(procedure)
    if found is not None:
        document.interpreter.call_event_handler(found, document.module, args, me=document)


def _sinks(book: Workbook, source: VBAObject, event: str, args: list[object]) -> None:
    interpreter = book.application.interpreter
    if interpreter is not None:
        interpreter.raise_event(source, event, args)


def _workbook(book: Workbook, event: str, args: list[object], on_application: str | None = None) -> None:
    if not _on(book):
        return
    _sinks(book, book, event, args)
    _handle(book, f"Workbook_{event}", args)
    if _on(book):
        _sinks(book, book.application, on_application or f"Workbook{event}", [book, *args])


def before_save(book: Workbook, save_as_ui: bool = False) -> bool:
    cancel = Slot(False, "Boolean")
    _workbook(book, "BeforeSave", [save_as_ui, cancel])
    return not to_bool(cancel.get())


def after_save(book: Workbook, success: bool) -> None:
    _workbook(book, "AfterSave", [success])


def before_close(book: Workbook) -> bool:
    cancel = Slot(False, "Boolean")
    _workbook(book, "BeforeClose", [cancel])
    return not to_bool(cancel.get())


def opened(book: Workbook) -> None:
    """Workbook_Open precedes application WorkbookOpen and activation."""
    _workbook(book, "Open", [])


def _window(book: Workbook, event: str) -> None:
    from pyopenvba.apps.excel._windows import window_of

    _workbook(book, event, [window_of(book)], event)


def book_left(book: Workbook) -> None:
    _window(book, "WindowDeactivate")
    _workbook(book, "Deactivate", [])


def book_entered(book: Workbook) -> None:
    _workbook(book, "Activate", [])
    _window(book, "WindowActivate")


def book_activated(book: Workbook, before: Workbook | None) -> None:
    """Changing workbook windows does not re-activate their active sheet."""
    if before is book:
        return
    if before is not None:
        book_left(before)
    book_entered(book)


def window_resized(book: Workbook) -> None:
    _window(book, "WindowResize")


def new_workbook(book: Workbook) -> None:
    if _on(book):
        _sinks(book, book.application, "NewWorkbook", [book])


def new_chart(book: Workbook, chart: VBAObject) -> None:
    _workbook(book, "NewChart", [chart])


def before_print(book: Workbook) -> bool:
    cancel = Slot(False, "Boolean")
    _workbook(book, "BeforePrint", [cancel])
    return not to_bool(cancel.get())


def _both(sheet: Worksheet | ChartSheet, event: str, args: list[object], *, application: bool = True) -> None:
    """One event, in the sheet's module and then in its workbook's, then the application's unless it has none to raise.
    A chart sheet raises its own as Chart_ (tests/fixtures/chart_sheet_members.json)."""
    from pyopenvba.apps.excel._chart_sheets import ChartSheet

    if not _on(sheet.book):
        return
    _sinks(sheet.book, sheet, event, args)
    _handle(sheet, f"{'Chart' if isinstance(sheet, ChartSheet) else 'Worksheet'}_{event}", args)
    if not _on(sheet.book):
        return
    combined = [sheet, *args]
    _sinks(sheet.book, sheet.book, f"Sheet{event}", combined)
    _handle(sheet.book, f"Workbook_Sheet{event}", combined)
    if application and _on(sheet.book):
        _sinks(sheet.book, sheet.book.application, f"Sheet{event}", combined)


def changed(target: Range) -> None:
    _both(target.sheet, "Change", [target])


def before_delete(sheet: Worksheet) -> None:
    _both(sheet, "BeforeDelete", [])


def sheet_left(sheet: Worksheet) -> None:
    _both(sheet, "Deactivate", [])


def sheet_entered(sheet: Worksheet) -> None:
    _both(sheet, "Activate", [])


def followed(sheet: Worksheet, link: VBAObject) -> None:
    _both(sheet, "FollowHyperlink", [link])


def pivot_updated(table: PivotTable, *, refreshed: bool = False) -> None:
    """PivotTableUpdate, with PivotTableChangeSync before it on a refresh and after it on any other edit."""
    if refreshed:
        pivot_synced(table)
    _both(table.sheet, "PivotTableUpdate", [table])
    if not refreshed:
        pivot_synced(table)


def pivot_synced(table: PivotTable) -> None:
    """PivotTableChangeSync, in the sheet's module and its workbook's, never the application's: Excel raises it on every edit of a
    pivot table, before PivotTableUpdate on a refresh and after it otherwise, and alone when a value written into the table is refused
    with 1004 (tests/fixtures/workbook_macro_events.json)."""
    _both(table.sheet, "PivotTableChangeSync", [table], application=False)


def pivot_refresh_calculated(sheet: Worksheet) -> None:
    """A completed report refresh notifies Calculate even in manual mode. A notification the sheet holds stays for
    the calculation that ends the refresh, or in manual calculation the next one (tests/fixtures/
    calculate_order.json refresh_held_automatic, refresh_held_manual)."""
    _both(sheet, "Calculate", [])


def refresh_calculated(book: Workbook) -> None:
    """The calculation that ends a refresh in automatic calculation: formulas reading the tables and volatile ones
    are worked out, and every sheet that worked one out or holds a notification raises Calculate (tests/fixtures/
    calculate_order.json refresh_*)."""
    calculator = book.calculator
    if calculator.automatic:
        calculated_all(book, calculator.calculate_all(background=True))


def selected(sheet: Worksheet, target: Range) -> None:
    _both(sheet, "SelectionChange", [target])


def after_calculate(application: Application) -> None:
    """Application.AfterCalculate, once a calculation is over: after Application.Calculate, CalculateFull and
    CalculateFullRebuild, Range.Calculate and the calculation an edit brings in automatic calculation, formulas or none,
    and never after Worksheet.Calculate (tests/fixtures/after_calculate.json)."""
    interpreter = application.interpreter
    if application.enable_events and interpreter is not None:
        interpreter.raise_event(application, "AfterCalculate", [])


def recalculated(book: Workbook, *, settled: set[tuple[str, int, int]] | None = None,
                 already_calculated: set[Worksheet] | None = None, announce: bool = True) -> None:
    """Work out what an edit left to work out, and raise Calculate for each sheet it worked out formulas on, then
    AfterCalculate unless ``announce`` is False.

    The model works a formula out when something reads it; Excel works
    them out after each edit, which only shows through this event, so the
    work is done here only when a module is listening for it.
    """
    calculator = book.calculator
    if not calculator.automatic:
        return
    _recalculated(book, settled, already_calculated)
    if announce:
        after_calculate(book.application)


def _recalculated(book: Workbook, settled: set[tuple[str, int, int]] | None,
                  already_calculated: set[Worksheet] | None) -> None:
    calculator = book.calculator
    if book.application.full_calculation:
        # Every workbook held to full calculation: each edit works every formula out again (see _calculation_state).
        from pyopenvba.apps.excel._calculation_state import worked_in_full

        worked_in_full(book.application)
        return
    calculator.automatic_pending = True
    prior = already_calculated or set()
    if not _on(book):
        # Ordinary formulas can stay lazy, but native Excel calculated
        # them and retains a notification for the next enabled cycle.
        calculator.pending_notifications.update(prior)
        calculator.pending_notifications.update(sheet for sheet in book.sheets_
            if any(cell.formula and cell.stale for cell in sheet.cells_.values()))
    # With Application.Iteration on, a loop of formulas is worked out at every calculation, as a volatile formula is.
    if not (calculator.has_volatile_formulas or calculator.has_pending_udfs or calculator.iterates
            or calculator.has_held_caches or calculator.pending_notifications) \
            and (not _on(book) or not _listening_for_calculate(book)):
        return
    if not _on(book) and not (calculator.has_volatile_formulas or calculator.has_pending_udfs or calculator.iterates):
        return
    worked = calculator.recalculated(settled=settled)
    calculated_all(book, list(set(worked) | prior))


def visibility_changed(book: Workbook) -> None:
    """Visibility edits calculate subtotal formulas, their dependents and volatile formulas, without Change."""
    book.calculator.invalidate_visibility()
    if book.calculator.automatic:
        calculated_all(book, book.calculator.calculate_all(background=True))


def calculated_all(book: Workbook, worked: list[Worksheet]) -> None:
    """Application.Calculate: events for sheets with stale or volatile formulas, in the order the sheets were made,
    whatever their tabs' order (tests/fixtures/calculate_order.json)."""
    pending = book.calculator.pending_notifications
    pending.update(worked)
    pending.intersection_update(book.sheets_)
    for sheet in sorted(book.sheets_, key=lambda sheet: sheet.sheet_id):
        if sheet in pending and _on(book):
            pending.remove(sheet)
            _both(sheet, "Calculate", [])


def after_edit(target: Range, *, calculate_first: bool = True, announced_last: bool = False) -> None:
    """An edit's Change and the Calculate it brings, in the order Excel raises them, AfterCalculate after Change where
    ``announced_last`` says so, as ClearContents raises it (tests/fixtures/after_calculate.json). What a dynamic-array
    formula spills is in place after the edit, as Excel spills it, events or none."""
    from pyopenvba.apps.excel._hidden_names import edited

    book = target.sheet.book
    edited(target)
    book.calculator.settle_spills()
    settled = book.calculator.settle_entries()
    settled_names = {key[0] for key in settled}
    calculated = {sheet for sheet in book.sheets_ if sheet.name.lower() in settled_names}
    if not _on(book):
        recalculated(book, settled=settled, already_calculated=calculated)
        return
    if calculate_first:
        recalculated(book, settled=settled, already_calculated=calculated, announce=not announced_last)
        changed(target)
        if announced_last and book.calculator.automatic:
            after_calculate(book.application)
    else:
        changed(target)
        recalculated(book, settled=settled, already_calculated=calculated)


def quietly_edited(book: Workbook) -> None:
    """An edit that raises no Change, as Subtotal and RemoveSubtotal make: what it left to work out, and the
    Calculate that brings (tests/fixtures/subtotals/)."""
    book.calculator.settle_spills()
    settled = book.calculator.settle_entries()
    settled_names = {key[0] for key in settled}
    recalculated(book, settled=settled,
                 already_calculated={sheet for sheet in book.sheets_ if sheet.name.lower() in settled_names})


def activated(book: Workbook, before: Worksheet | ChartSheet | None, after: Worksheet | ChartSheet) -> None:
    """Another sheet made active: the one left deactivates, then the new one activates."""
    if not _on(book) or before is after:
        return
    if before is not None:
        _both(before, "Deactivate", [])
    _both(after, "Activate", [])


def new_sheet(book: Workbook, sheet: Worksheet | ChartSheet, left: Worksheet | ChartSheet | None) -> None:
    """A sheet added: NewSheet, then the sheet it left deactivates and the new one activates."""
    if not _on(book):
        return
    _sinks(book, book, "NewSheet", [sheet])
    _handle(book, "Workbook_NewSheet", [sheet])
    _sinks(book, book.application, "WorkbookNewSheet", [book, sheet])
    activated(book, left, sheet)


def _listening_for_calculate(book: Workbook) -> bool:
    from pyopenvba.apps.excel._vbide import VBProject

    interpreter = book.application.interpreter
    if interpreter is not None:
        if interpreter.listens_for_event(book, "SheetCalculate") or interpreter.listens_for_event(book.application, "SheetCalculate"):
            return True
        if any(interpreter.listens_for_event(sheet, "Calculate") for sheet in book.sheets_):
            return True
    for owner in [book, *book.sheets_]:
        if isinstance(book.vbide_project, VBProject):
            entry = next((entry for entry in book.vbide_project.components.entries
                          if entry.name.casefold() == owner.code_name.casefold()), None)
            name = 'Workbook_SheetCalculate' if owner is book else 'Worksheet_Calculate'
            if entry is not None and entry.pending and any(row[0].casefold() == name.casefold()
                                                           for row in entry.code_module.procedures()):
                return True
        document = owner.vba_document
        if document is not None and (document.module.procedure("Worksheet_Calculate") is not None
                                     or document.module.procedure("Workbook_SheetCalculate") is not None):
            return True
    return False
