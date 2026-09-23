"""The events Excel raises in a sheet's and the workbook's own modules.

Measured in live Excel (scripts/measure_events.py, tests/fixtures/events.json).
Each event runs its handler in the sheet's module first -- Worksheet_Change
-- and then the workbook's -- Workbook_SheetChange, the sheet first among
its arguments -- and none of them while Application.EnableEvents is False.
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
  Application.Calculate raises it for every sheet that holds a formula.
* SelectionChange: a Select that moves the selection.
* Deactivate, then Activate: another sheet made active; a new sheet is
  NewSheet first, and deactivates the sheet it was placed beside, or the
  active one.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.interpreter._objects import VBAObject

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Workbook, Worksheet


def _on(book: Workbook) -> bool:
    return bool(book.application.enable_events)


def _handle(owner: VBAObject, procedure: str, args: list[object]) -> None:
    """Run ``procedure`` of ``owner``'s own module, if it has one; a module without it is simply not listening."""
    document = owner.vba_document
    if document is None:
        return
    found = document.module.procedure(procedure)
    if found is not None:
        document.interpreter.call(found, document.module, args, {}, me=document)


def _both(sheet: Worksheet, event: str, args: list[object]) -> None:
    """One event, in the sheet's module and then in its workbook's."""
    if not _on(sheet.book):
        return
    _handle(sheet, f"Worksheet_{event}", args)
    _handle(sheet.book, f"Workbook_Sheet{event}", [sheet, *args])


def changed(target: Range) -> None:
    _both(target.sheet, "Change", [target])


def selected(sheet: Worksheet, target: Range) -> None:
    _both(sheet, "SelectionChange", [target])


def recalculated(book: Workbook) -> None:
    """Work out what an edit left to work out, and raise Calculate for each sheet it worked out formulas on.

    The model works a formula out when something reads it; Excel works
    them out after each edit, which only shows through this event, so the
    work is done here only when a module is listening for it.
    """
    if not _on(book) or not _listening_for_calculate(book):
        return
    for sheet in book.calculator.recalculated():
        _both(sheet, "Calculate", [])


def calculated_all(book: Workbook) -> None:
    """Application.Calculate: Calculate for every sheet that holds a formula."""
    if not _on(book):
        return
    for sheet in list(book.sheets_):
        if any(cell.formula for cell in sheet.cells_.values()):
            _both(sheet, "Calculate", [])


def after_edit(target: Range, *, calculate_first: bool = True) -> None:
    """An edit's Change and the Calculate it brings, in the order Excel raises them."""
    book = target.sheet.book
    if not _on(book):
        return
    if calculate_first:
        recalculated(book)
        changed(target)
    else:
        changed(target)
        recalculated(book)


def activated(book: Workbook, before: Worksheet | None, after: Worksheet) -> None:
    """Another sheet made active: the one left deactivates, then the new one activates."""
    if not _on(book) or before is after:
        return
    if before is not None:
        _both(before, "Deactivate", [])
    _both(after, "Activate", [])


def new_sheet(book: Workbook, sheet: Worksheet, left: Worksheet | None) -> None:
    """A sheet added: NewSheet, then the sheet it left deactivates and the new one activates."""
    if not _on(book):
        return
    _handle(book, "Workbook_NewSheet", [sheet])
    activated(book, left, sheet)


def _listening_for_calculate(book: Workbook) -> bool:
    for owner in [book, *book.sheets_]:
        document = owner.vba_document
        if document is not None and (document.module.procedure("Worksheet_Calculate") is not None
                                     or document.module.procedure("Workbook_SheetCalculate") is not None):
            return True
    return False
