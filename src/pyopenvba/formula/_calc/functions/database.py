"""Database functions: DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN,
DPRODUCT, DSTDEV, DSTDEVP, DVAR and DVARP.

A database is a range whose first row labels its columns, one record to a
row below. The criteria are another range: its first row names columns,
and every row under it is one alternative, a record counting when all the
conditions in any one row hold. A blank condition holds for everything.

A condition reads like a COUNTIF criterion, ``">10"`` or ``"<>Plum"``,
with one difference Excel shares with its Advanced Filter: text with no
operator matches text that begins with it, so ``Pear`` also finds
``Pearmain``. A formula under a label that names no column is a computed
condition: it is written for the first record, and Excel evaluates it once
for every record with its relative references moved to that record's row,
so ``=E2>80`` asks about each record's own E.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from functools import cache

from pyopenvba.formula._calc import collate as _collate
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.aggregate import spread
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import Criterion, area_of, criterion
from pyopenvba.formula._calc.nodes import Node
from pyopenvba.formula._calc.numbers import total
from pyopenvba.formula._calc.parser import parse
from pyopenvba.formula._calc.precise import divide, multiply, square_root
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import (
    DIV0,
    NUM,
    VALUE,
    Area,
    Empty,
    ExcelError,
    Scalar,
    Value,
    scalar_text,
)
from pyopenvba.formula._calc.host import translate_formula
from pyopenvba.formula._calc.cells import CellError

_OPERATORS = ("<", ">", "=")

Test = Callable[[int], bool]


@cache
def _shifted(formula: str, rows: int) -> Node:
    return parse(translate_formula(formula, rows, 0))


def _label(value: Scalar) -> str:
    return scalar_text(value) if not isinstance(value, CellError) else value.code


def _column(labels: list[str], name: str) -> int | None:
    for index, label in enumerate(labels):
        if label and _collate.equal(label, name):
            return index
    return None


def _condition(context: Context, value: Scalar) -> Criterion:
    """A condition cell as a criterion: plain text matches as a prefix."""
    if isinstance(value, str) and not value.startswith(_OPERATORS):
        parsed = criterion(context, value)
        if isinstance(parsed.value, str):
            return Criterion("=", value + "*", True)
        return parsed
    return criterion(context, value)


def _computed(context: Context, area: Area, database: Area, row: int, column: int, formula: str) -> Test:
    """A formula condition, evaluated for the record in ``record``'s row."""
    first = database.top + 1

    def holds(record: int) -> bool:
        offset = record - first
        inner = Context(context.book, area.sheet, row + offset, column, today=context.today, now=context.now)
        try:
            found = inner.first(inner.formula(_shifted(formula, offset)))
            return not isinstance(found, CellError) and inner.logical(found)
        except ExcelError:
            return False

    return holds


def _criteria(context: Context, area: Area, database: Area, labels: list[str]) -> list[list[Test]]:
    """Each alternative of a criteria range, as the tests a record's row
    must pass."""
    book = context.book
    names = [_label(book.cell(area.sheet, area.top, column)) for column in range(area.left, area.right + 1)]
    alternatives: list[list[Test]] = []
    for row in range(area.top + 1, area.bottom + 1):
        tests: list[Test] = []
        for offset, name in enumerate(names):
            column = area.left + offset
            value = book.cell(area.sheet, row, column)
            if isinstance(value, Empty) or value == "":
                continue
            field = _column(labels, name) if name else None
            formula = book.formula_text(area.sheet, row, column)
            if field is None and formula is not None:
                tests.append(_computed(context, area, database, row, column, formula))
            elif field is None:
                tests.append(lambda _record: False)
            else:
                test = _condition(context, value)
                cell_column = database.left + field

                def matches(record: int, test: Criterion = test, cell_column: int = cell_column) -> bool:
                    return test.matches(book.cell(database.sheet, record, cell_column), context)

                tests.append(matches)
        alternatives.append(tests)
    return alternatives


def _field(context: Context, field: Scalar, labels: list[str]) -> int | None:
    """The database column a field argument names, by label or by number
    from 1; ``None`` when it is left out."""
    if isinstance(field, Empty):
        return None
    if isinstance(field, CellError):
        raise ExcelError(field)
    if isinstance(field, float):
        index = math.trunc(field)
        if not 1 <= index <= len(labels):
            raise ExcelError(VALUE)
        return index - 1
    if isinstance(field, str):
        found = _column(labels, field)
        if found is not None:
            return found
    raise ExcelError(VALUE)


def records(context: Context, database: Value, field: Scalar, criteria: Value) -> tuple[list[Scalar], bool]:
    """The field's value in every record the criteria select, and whether
    the field was named at all."""
    table = area_of(database)
    wanted = area_of(criteria)
    book = context.book
    labels = [_label(book.cell(table.sheet, table.top, column)) for column in range(table.left, table.right + 1)]
    column = _field(context, field, labels)
    alternatives = _criteria(context, wanted, table, labels)
    found: list[Scalar] = []
    for row in range(table.top + 1, table.bottom + 1):
        if any(all(test(row) for test in tests) for tests in alternatives):
            found.append(book.cell(table.sheet, row, table.left + column) if column is not None else True)
    return found, column is not None


def _numbers(values: list[Scalar]) -> list[float]:
    found: list[float] = []
    for value in values:
        if isinstance(value, CellError):
            raise ExcelError(value)
        if isinstance(value, float):
            found.append(value)
    return found


def _selected(context: Context, database: Value, field: Scalar, criteria: Value) -> list[float]:
    values, named = records(context, database, field, criteria)
    if not named:
        raise ExcelError(VALUE)
    return _numbers(values)


@function("DSUM", R, V, R)
def DSUM(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return checked(total(_selected(context, database, field, criteria)))


@function("DAVERAGE", R, V, R)
def DAVERAGE(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    values = _selected(context, database, field, criteria)
    if not values:
        return DIV0
    return checked(divide(total(values), float(len(values))))


@function("DCOUNT", R, V, R)
def DCOUNT(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    values, named = records(context, database, field, criteria)
    if not named:
        return float(len(values))
    return float(sum(1 for value in values if isinstance(value, float)))


@function("DCOUNTA", R, V, R)
def DCOUNTA(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    values, named = records(context, database, field, criteria)
    if not named:
        return float(len(values))
    return float(sum(1 for value in values if not isinstance(value, Empty)))


@function("DGET", R, V, R)
def DGET(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    values, named = records(context, database, field, criteria)
    if not named or not values:
        return VALUE
    if len(values) > 1:
        return NUM
    return values[0]


@function("DMAX", R, V, R)
def DMAX(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return max(_selected(context, database, field, criteria), default=0.0)


@function("DMIN", R, V, R)
def DMIN(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return min(_selected(context, database, field, criteria), default=0.0)


@function("DPRODUCT", R, V, R)
def DPRODUCT(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    values = _selected(context, database, field, criteria)
    if not values:
        return 0.0
    result = 1.0
    for value in values:
        result = multiply(result, value)
    return checked(result)


def _spread(context: Context, database: Value, field: Scalar, criteria: Value, *, sample: bool, root: bool) -> Value:
    """The variance of the selected values, or its root, as VAR takes it."""
    found = spread(_selected(context, database, field, criteria), sample=sample)
    if isinstance(found, CellError):
        return found
    return checked(square_root(found) if root else found)


@function("DSTDEV", R, V, R)
def DSTDEV(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return _spread(context, database, field, criteria, sample=True, root=True)


@function("DSTDEVP", R, V, R)
def DSTDEVP(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return _spread(context, database, field, criteria, sample=False, root=True)


@function("DVAR", R, V, R)
def DVAR(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return _spread(context, database, field, criteria, sample=True, root=False)


@function("DVARP", R, V, R)
def DVARP(context: Context, database: Value, field: Scalar, criteria: Value) -> Value:
    return _spread(context, database, field, criteria, sample=False, root=False)


__all__ = ["records"]
