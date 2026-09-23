"""COUNTIF, SUMIF, AVERAGEIF and their several-criteria forms.

A criterion is parsed once by :func:`.common.criterion` and tested against
every cell of its range. The range to add up or average may be a different
size from the criteria range: SUMIF reads it from its top left cell at the
criteria range's size, while SUMIFS requires every range to be one size.
"""

from __future__ import annotations

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import (
    Criterion,
    area_of,
    block,
    criterion,
    matching,
    shaped,
)
from pyopenvba.formula._calc.numbers import total
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import DIV0, VALUE, Area, ExcelError, Scalar, Value
from pyopenvba.formula._calc.cells import CellError


def _pairs(context: Context, args: tuple[Value, ...]) -> list[tuple[Area, Criterion]]:
    if len(args) % 2:
        raise ExcelError(VALUE)
    pairs: list[tuple[Area, Criterion]] = []
    for index in range(0, len(args), 2):
        pairs.append((area_of(args[index]), criterion(context, context.first(args[index + 1]))))
    return pairs


def _values_at(context: Context, area: Area, positions: list[tuple[int, int]]) -> list[Scalar]:
    """The values of ``area`` at the matching offsets."""
    if not positions:
        return []
    height = max(row for row, _ in positions) + 1
    width = max(column for _, column in positions) + 1
    grid = block(context, area, min(height, area.height), min(width, area.width))
    return [grid[row][column] for row, column in positions if row < len(grid) and column < len(grid[0])]


def _numbers_at(context: Context, area: Area, positions: list[tuple[int, int]]) -> list[float]:
    found: list[float] = []
    for value in _values_at(context, area, positions):
        if isinstance(value, CellError):
            raise ExcelError(value)
        if isinstance(value, float):
            found.append(value)
    return found


def _sum(values: list[float]) -> float:
    return checked(total(values))


@function("COUNTIF", R, V)
def COUNTIF(context: Context, range_: Value, test: Scalar) -> Value:
    area = area_of(range_)
    found, beyond = matching(context, [(area, criterion(context, test))], [])
    return float(len(found) + beyond)


@function("COUNTIFS", R, V, maximum=254, repeat=2)
def COUNTIFS(context: Context, *args: Value) -> Value:
    found, beyond = matching(context, _pairs(context, args), [])
    return float(len(found) + beyond)


@function("SUMIF", R, V, R, minimum=2)
def SUMIF(context: Context, range_: Value, test: Scalar, sum_range: Value | None = None) -> Value:
    area = area_of(range_)
    target = area if sum_range is None else shaped(area_of(sum_range), area)
    found, _ = matching(context, [(area, criterion(context, test))], [target])
    return _sum(_numbers_at(context, target, found))


@function("SUMIFS", R, R, V, maximum=255, repeat=2)
def SUMIFS(context: Context, sum_range: Value, *args: Value) -> Value:
    target = area_of(sum_range)
    pairs = _pairs(context, args)
    if any(area.height != target.height or area.width != target.width for area, _ in pairs):
        return VALUE
    found, _ = matching(context, pairs, [target])
    return _sum(_numbers_at(context, target, found))


@function("AVERAGEIF", R, V, R, minimum=2)
def AVERAGEIF(context: Context, range_: Value, test: Scalar, average_range: Value | None = None) -> Value:
    area = area_of(range_)
    target = area if average_range is None else shaped(area_of(average_range), area)
    found, _ = matching(context, [(area, criterion(context, test))], [target])
    values = _numbers_at(context, target, found)
    if not values:
        return DIV0
    return checked(total(values) / len(values))


@function("AVERAGEIFS", R, R, V, maximum=255, repeat=2)
def AVERAGEIFS(context: Context, average_range: Value, *args: Value) -> Value:
    target = area_of(average_range)
    pairs = _pairs(context, args)
    if any(area.height != target.height or area.width != target.width for area, _ in pairs):
        return VALUE
    found, _ = matching(context, pairs, [target])
    values = _numbers_at(context, target, found)
    if not values:
        return DIV0
    return checked(total(values) / len(values))


@function("MAXIFS", R, R, V, maximum=255, repeat=2)
def MAXIFS(context: Context, max_range: Value, *args: Value) -> Value:
    target = area_of(max_range)
    pairs = _pairs(context, args)
    if any(area.height != target.height or area.width != target.width for area, _ in pairs):
        return VALUE
    found, _ = matching(context, pairs, [target])
    values = _numbers_at(context, target, found)
    return max(values) if values else 0.0


@function("MINIFS", R, R, V, maximum=255, repeat=2)
def MINIFS(context: Context, min_range: Value, *args: Value) -> Value:
    target = area_of(min_range)
    pairs = _pairs(context, args)
    if any(area.height != target.height or area.width != target.width for area, _ in pairs):
        return VALUE
    found, _ = matching(context, pairs, [target])
    values = _numbers_at(context, target, found)
    return min(values) if values else 0.0


__all__: list[str] = []
