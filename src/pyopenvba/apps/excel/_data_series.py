"""Range.DataSeries: a series filled from the first cell of each row or column, as Excel fills it.

Measured in live Excel (scripts/measure_data_series.py,
tests/fixtures/data_series.json), every value to the bit.

- Rowcol says whether each column is filled down (xlColumns) or each row
  across (xlRows); left out, a range with more rows than columns is filled
  down and any other across. Each area is filled on its own.
- Each line starts from its first cell, which must hold a number, a date
  among them; a line whose first cell is blank, text or logical is left as
  it is. The first cell stays as it is, a formula too, and the cells after
  it take its value as a constant and its format.
- Linear adds Step, 1 when left out: value k is first + k * Step, in
  Doubles, with no rounding to 15 digits. Growth multiplies: first *
  Step^k, the power worked out by repeated squaring. Chronological counts
  in Date's unit: days as linear, weekdays skipping Saturday and Sunday,
  and months and years kept on the first cell's day, the last day of a
  shorter month instead, its time kept. AutoFill fills as AutoFill does
  from the line's leading cells.
- Stop ends a line before the first value past it, in the direction the
  series runs; the cells after keep what they held. The first cell is
  kept even past Stop.
- Trend over one number is the Step series; over two or more Excel fits
  them by least squares, whose bits the model does not reproduce (see
  _autofill), so it reports itself unsupported.
- DataSeries returns True, raises one Change for the cells it filled and
  leaves the selection as it was. Another Rowcol, Type or Date, or a Step
  or Stop that is no number, is error 1004.
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, VBADate, error, to_bool, to_number

if TYPE_CHECKING:
    from collections.abc import Callable

    from pyopenvba.apps.excel._model import Range, Worksheet
    from pyopenvba.apps.excel._styles import Style

ROWS, COLUMNS = 1, 2
LINEAR, GROWTH, CHRONOLOGICAL, AUTOFILL = -4132, 2, 3, 4
DAY, WEEKDAY, MONTH, YEAR = 1, 2, 3, 4
_FAILED = "Unable to get the DataSeries property of the Range class"


def _number(value: object) -> float:
    """A number DataSeries takes; text that is none is error 1004."""
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, bool):
        raise VBAUnsupportedError("DataSeries with a logical value for a number is not implemented")
    if isinstance(value, (int, float, Decimal)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(to_number(value))
        except VBARuntimeError:
            pass
    raise error(1004, _FAILED)


def _choice(value: object, default: int, allowed: tuple[int, ...]) -> int:
    if value is MISSING:
        return default
    chosen = _number(value)
    if chosen not in allowed:
        raise error(1004, _FAILED)
    return int(chosen)


def data_series(target: Range, rowcol: object, kind: object, unit: object, step: object, stop: object,
                trend: object) -> bool:
    from pyopenvba.apps.excel import _arrays, _events
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._protection import whole

    sheet = target.sheet
    chosen_kind = _choice(kind, LINEAR, (LINEAR, GROWTH, CHRONOLOGICAL, AUTOFILL))
    chosen_unit = _choice(unit, DAY, (DAY, WEEKDAY, MONTH, YEAR))
    by = 1.0 if step is MISSING else _number(step)
    until = None if stop is MISSING else _number(stop)
    fitted = False if trend is MISSING else to_bool(trend)
    lines_by = None if rowcol is MISSING else _choice(rowcol, ROWS, (ROWS, COLUMNS))
    whole(sheet, list(target.areas), "DataSeries")
    _arrays.refuse(sheet, list(target.areas), "DataSeries over")
    written: list[Area] = []
    for area in target.areas:
        down = (area.rows > area.columns) if lines_by is None else lines_by == COLUMNS
        lines = range(area.left, area.right + 1) if down else range(area.top, area.bottom + 1)
        for line in lines:
            cells = [(row, line) for row in range(area.top, area.bottom + 1)] if down \
                else [(line, column) for column in range(area.left, area.right + 1)]
            if len(cells) < 2:
                continue
            filled = _fill_line(sheet, cells, chosen_kind, chosen_unit, by, until, fitted)
            if filled:
                (top, left), (bottom, right) = cells[1], cells[filled]
                written.append(Area(top, left, bottom, right, sheet.name))
    if written:
        _events.after_edit(Range(sheet, written))
    return True


def _fill_line(sheet: Worksheet, cells: list[tuple[int, int]], kind: int, unit: int, step: float,
               stop: float | None, fitted: bool) -> int:
    """Fill one line; the index of the last cell written, 0 for none."""
    from pyopenvba.apps.excel._model import Range

    values = [sheet.book.calculator.value_of(sheet.name, *cell) for cell in cells]
    if kind == AUTOFILL:
        lead = next((at for at, value in enumerate(values) if _blank(value)), len(values))
        if lead == 0 or lead == len(values):
            return 0
        from pyopenvba.apps.excel._autofill import autofill

        (top, left), (bottom, right) = cells[0], cells[lead - 1]
        autofill(Range(sheet, [Area(top, left, bottom, right, sheet.name)]),
                 Range(sheet, [Area(cells[0][0], cells[0][1], cells[-1][0], cells[-1][1], sheet.name)]), MISSING)
        return len(cells) - 1
    first = values[0]
    if isinstance(first, VBADate):
        first = first.serial
    if isinstance(first, bool) or not isinstance(first, (int, float)):
        return 0
    if fitted:
        numbers = 0
        for value in values:
            if isinstance(value, bool) or not isinstance(value, (int, float, VBADate)):
                break
            numbers += 1
        if numbers >= 2:
            raise VBAUnsupportedError("DataSeries with Trend over two or more numbers is not implemented: Excel fits "
                                      "them by least squares, whose bits the model does not reproduce")
    value_at = _series(float(first), kind, unit, step, sheet.book.epoch_1904)
    rising = _rising(value_at, kind, step)
    last = 0
    style = sheet.style_at(*cells[0])
    for index in range(1, len(cells)):
        value = value_at(index)
        if stop is not None and rising is not None and (value > stop if rising else value < stop):
            break
        _write(sheet, cells[index], value, style)
        last = index
    return last


def _blank(value: object) -> bool:
    from pyopenvba.formula._values import BLANK
    from pyopenvba.interpreter._values import EMPTY

    return value is EMPTY or value is BLANK or value is None


def _rising(value_at: Callable[[int], float], kind: int, step: float) -> bool | None:
    """Which way the series runs, for Stop: None where it stays where it is."""
    if kind == GROWTH:
        first, second = value_at(0), value_at(1)
        return None if first == second else second > first
    return None if step == 0 else step > 0


def _series(first: float, kind: int, unit: int, step: float, epoch_1904: bool) -> Callable[[int], float]:
    if kind == GROWTH:
        return lambda k: first * _power(step, k)
    if kind == CHRONOLOGICAL and unit != DAY:
        return lambda k: _calendar(first, unit, step, k, epoch_1904)
    return lambda k: first + k * step


def _power(base: float, exponent: int) -> float:
    """base**exponent by repeated squaring, a multiplication a step, as Excel's growth series works it out."""
    result = 1.0
    while exponent:
        if exponent & 1:
            result *= base
        base *= base
        exponent >>= 1
    return result


def _origin(epoch_1904: bool) -> _dt.date:
    return _dt.date(1904, 1, 1) if epoch_1904 else _dt.date(1899, 12, 30)


def _calendar(first: float, unit: int, step: float, k: int, epoch_1904: bool) -> float:
    """The k-th date after ``first`` counted in weekdays, months or years, its time kept."""
    if not step.is_integer():
        raise VBAUnsupportedError("DataSeries by weekdays, months or years with a Step that is no whole number is "
                                  "not implemented")
    day = int(first // 1)
    if day < 61 and not epoch_1904:
        raise VBAUnsupportedError("DataSeries by weekdays, months or years before March 1900 is not implemented")
    time = first - day
    origin = _origin(epoch_1904)
    start = origin + _dt.timedelta(days=day)
    count = int(step) * k
    if unit == WEEKDAY:
        moved = start
        direction = 1 if count >= 0 else -1
        for _ in range(abs(count)):
            moved += _dt.timedelta(days=direction)
            while moved.weekday() >= 5:
                moved += _dt.timedelta(days=direction)
    else:
        months = count if unit == MONTH else 12 * count
        year, month = divmod(start.month - 1 + months, 12)
        year += start.year
        month += 1
        length = (_dt.date(year + month // 12, month % 12 + 1, 1) - _dt.timedelta(days=1)).day
        moved = _dt.date(year, month, min(start.day, length))
    return (moved - origin).days + time


def _write(sheet: Worksheet, at: tuple[int, int], value: float, style: Style) -> None:
    from pyopenvba.apps.excel._model import Cell

    default = sheet.book.stylesheet.default
    cell = Cell(value=value, style=None if style == default else style)
    sheet.book.stylesheet.meet(cell.style)
    sheet.cells_[at] = cell
    sheet.settle(*at)
    sheet.cell_changed(*at)
