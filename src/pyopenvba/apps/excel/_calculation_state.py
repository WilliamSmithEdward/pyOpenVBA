"""How far calculation has got, and what steers it: Application.CalculationState, CalculationVersion and
CalculationInterruptKey, Workbook.CalculationVersion and ForceFullCalculation, Range.Dirty and
CalculateRowMajorOrder, and the files Excel works out again as it opens them (scripts/measure_calculation_state.py,
tests/fixtures/calculation_state.json).

- CalculationState is xlCalculating while a formula's VBA Function runs, and xlPending while a workbook has a
  formula to work out under manual calculation: one entered, one reading a cell written, every formula on a sheet
  a row is put into, one Range.Dirty marked, one naming a defined name added, redefined or deleted; a value
  nothing reads, a format, a name no formula names, a name's comment or a sheet added leave it xlDone, as
  automatic calculation always does.
- A file whose calcId is lower than 191028, or that has none or says fullCalcOnLoad, Excel works out in full as it
  opens it, keeping it saved, and from then on gives it the calcId 191029, its own; under manual calculation it
  keeps the file's values and calcId and is pending until the next Application.Calculate, which leaves it unsaved.
  Any other file keeps its values and its calcId, a newer one too, and saves it again.
- A workbook's ForceFullCalculation set True, or a file opened that says forceFullCalc, holds every workbook to
  full calculation until Excel quits, whatever ForceFullCalculation says after: each edit works out every formula
  of every workbook, each calculation does, a workbook it works out goes unsaved, a file opened is worked out as
  it opens, and CalculationState is xlPending for good.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Final

from pyopenvba._a1 import Area
from pyopenvba.interpreter._values import NULL, VBAInt, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Range, Workbook

#: The calculation engine Excel reports, and writes as a workbook's calcId.
VERSION: Final = 191029
#: The lowest calcId whose values Excel takes as a file has them; a file of a lower one, or of none, it works out.
LOWEST_KEPT: Final = 191028
_DONE, _CALCULATING, _PENDING = 0, 1, 2
_MANUAL: Final = -4135


def state(application: Application) -> object:
    """Application.CalculationState."""
    if application.udf_callers:
        return VBAInt(_CALCULATING, "Long")
    if application.full_calculation:
        return VBAInt(_PENDING, "Long")
    if application.calculation == _MANUAL and any(pending(book) for book in application.workbooks_.books):
        return VBAInt(_PENDING, "Long")
    return VBAInt(_DONE, "Long")


def pending(book: Workbook) -> bool:
    """Whether a workbook has a formula to work out at the next calculation."""
    return book.full_pending or any(cell.formula and cell.stale for sheet in book.sheets_
                                    for cell in sheet.cells_.values())


def set_interrupt_key(application: Application, value: object) -> None:
    """CalculationInterruptKey: xlNoKey, xlEscKey or xlAnyKey, taken as a Long is; any other number is error 1004."""
    if value is NULL:
        raise error(94)
    key = int(to_integer(value, "Long"))
    if not 0 <= key <= 2:
        raise error(1004, "Unable to set the CalculationInterruptKey property of the Application class")
    application.calculation_interrupt_key = key


def set_force(book: Workbook, value: object) -> None:
    """ForceFullCalculation, taken as a Boolean is; True holds every workbook to full calculation from now on."""
    if value is NULL:
        raise error(94)
    book.force_full_calculation = to_bool(value)
    if book.force_full_calculation:
        book.application.full_calculation = True


def opened(book: Workbook, calc_id: int, full_on_load: bool, forced: bool) -> None:
    """A workbook read from a file: the calcId it has, the full calculation it may want as it opens."""
    application = book.application
    book.calculation_version = calc_id
    book.force_full_calculation = forced
    application.full_calculation |= forced
    if not (calc_id < LOWEST_KEPT or full_on_load or application.full_calculation):
        return
    formulas = [cell for sheet in book.sheets_ for cell in sheet.cells_.values() if cell.formula]
    if application.calculation == _MANUAL:
        book.full_pending = True
        return
    for cell in formulas:
        cell.stale = True
    if calc_id < LOWEST_KEPT:
        book.calculation_version = VERSION
    if application.full_calculation and formulas:
        book.saved = False


def in_full(book: Workbook) -> bool:
    """Whether the next calculation asked for works ``book`` out in full: every workbook while one has been held
    to full calculation, and one opened wanting it under manual calculation."""
    return book.application.full_calculation or book.full_pending


def calculated(book: Workbook, full: bool) -> None:
    """After a calculation asked for: a workbook it worked out in full that has formulas goes unsaved, and one
    waiting since it was opened takes Excel's calcId."""
    if not full:
        return
    if book.full_pending or book.application.full_calculation:
        if any(cell.formula for sheet in book.sheets_ for cell in sheet.cells_.values()):
            book.saved = False
    if book.full_pending:
        book.full_pending = False
        if book.calculation_version < LOWEST_KEPT:
            book.calculation_version = VERSION


def worked_in_full(application: Application) -> None:
    """An edit under automatic calculation while every workbook is held to full calculation: each formula of each
    workbook worked out again, with each workbook's Calculate."""
    from pyopenvba.apps.excel import _events

    for book in list(application.workbooks_.books):
        worked = book.calculator.calculate_all(full=True, background=True)
        calculated(book, True)
        _events.calculated_all(book, worked)


def dirty(target: Range) -> None:
    """Range.Dirty: the range's formulas, and what reads a cell of it, to work out again, at once under automatic
    calculation."""
    from pyopenvba.apps.excel import _events

    calculator = target.sheet.book.calculator
    for area in target.areas:
        calculator.dirtied(target.sheet.name, area)
    if calculator.automatic:
        _events.calculated_all(target.sheet.book, calculator.calculate_all(volatile=False, background=True))


def row_major(target: Range) -> None:
    """Range.CalculateRowMajorOrder under manual calculation: each formula of the range worked out once, a row at a
    time from the left, reading what the range's other cells hold then, worked out or not, where Range.Calculate
    works out first what a cell reads; as with Range.Calculate, they join the next calculation asked for."""
    from pyopenvba.apps.excel import _events

    calculator = target.sheet.book.calculator
    _events.calculated_all(target.sheet.book, [])
    name = target.sheet.name.lower()
    for area in target.areas:
        for row, column in sorted(position for position, cell in target.sheet.cells_.items()
                                  if cell.formula and area.contains(*position)):
            worked = calculator.calculate_all(target.sheet, areas=(Area(row, column, row, column),), force=True,
                                              volatile=False)
            calculator.pending_notifications.update(worked)
            target.sheet.cells_[(row, column)].stale = True
            compiled = calculator.compiled.get((name, row, column))
            if compiled is not None:
                compiled.udf_pending = compiled.calls_user
