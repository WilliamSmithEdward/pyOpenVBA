"""Matrix functions: MDETERM, MINVERSE, MMULT and MUNIT.

MDETERM eliminates with partial pivoting, the largest entry of each column
brought up, and multiplies the pivots: measured, that is how
``MDETERM({1,3,8,5;1,3,6,1;1,1,1,0;7,3,10,2})`` comes to be
``87.999999999999972`` rather than 88. MINVERSE solves on the same
factorization, and a matrix with a zero pivot has no inverse: ``#NUM!``.
"""

from __future__ import annotations

from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.functions.arithmetic import checked
from pyopenvba.formula._calc.functions.common import matrix
from pyopenvba.formula._calc.registry import R, V, function
from pyopenvba.formula._calc.values import NUM, VALUE, Array, ExcelError, Scalar, Value
from pyopenvba.formula._calc.cells import CellError


def numeric(context: Context, value: Value) -> list[list[float]]:
    """A range or array of numbers as rows of floats: anything else in it,
    a blank included, is ``#VALUE!``, and an error is itself."""
    rows: list[list[float]] = []
    for row in matrix(context, value).rows:
        numbers: list[float] = []
        for item in row:
            if isinstance(item, CellError):
                raise ExcelError(item)
            if not isinstance(item, float):
                raise ExcelError(VALUE)
            numbers.append(item)
        rows.append(numbers)
    return rows


def _square(context: Context, value: Value) -> list[list[float]]:
    rows = numeric(context, value)
    if not rows or any(len(row) != len(rows) for row in rows):
        raise ExcelError(VALUE)
    return rows


def _factor(rows: list[list[float]]) -> tuple[list[list[float]], list[int], float] | None:
    """The LU factorization with partial pivoting, in place of a copy: the
    factors, the row each came from, and the sign of the permutation, or
    ``None`` when a pivot is zero."""
    a = [list(row) for row in rows]
    size = len(a)
    order = list(range(size))
    sign = 1.0
    for k in range(size):
        pivot = max(range(k, size), key=lambda i: abs(a[i][k]))
        if a[pivot][k] == 0.0:
            return None
        if pivot != k:
            a[k], a[pivot] = a[pivot], a[k]
            order[k], order[pivot] = order[pivot], order[k]
            sign = -sign
        for i in range(k + 1, size):
            factor = a[i][k] / a[k][k]
            a[i][k] = factor
            for j in range(k + 1, size):
                a[i][j] -= factor * a[k][j]
    return a, order, sign


@function("MDETERM", R)
def MDETERM(context: Context, value: Value) -> Value:
    rows = _square(context, value)
    factored = _factor(rows)
    if factored is None:
        return 0.0
    a, _, sign = factored
    result = sign
    for k in range(len(a)):
        result *= a[k][k]
    return checked(result)


@function("MINVERSE", R)
def MINVERSE(context: Context, value: Value) -> Value:
    rows = _square(context, value)
    factored = _factor(rows)
    if factored is None:
        return NUM
    a, order, _ = factored
    size = len(a)
    inverse: list[list[Scalar]] = [[0.0] * size for _ in range(size)]
    for column in range(size):
        # Solve A x = e_column: forward through L, back through U.
        x = [1.0 if order[i] == column else 0.0 for i in range(size)]
        for i in range(size):
            for j in range(i):
                x[i] -= a[i][j] * x[j]
        for i in reversed(range(size)):
            for j in range(i + 1, size):
                x[i] -= a[i][j] * x[j]
            x[i] /= a[i][i]
        for i in range(size):
            inverse[i][column] = checked(x[i])
    return Array(inverse)


@function("MMULT", R, R)
def MMULT(context: Context, first: Value, second: Value) -> Value:
    left = numeric(context, first)
    right = numeric(context, second)
    if len(left[0]) != len(right):
        return VALUE
    product: list[list[Scalar]] = []
    for row in left:
        items: list[Scalar] = []
        for column in range(len(right[0])):
            total = 0.0
            for k, value in enumerate(row):
                total += value * right[k][column]
            items.append(checked(total))
        product.append(items)
    return Array(product)


@function("MUNIT", V)
def MUNIT(context: Context, size: Scalar) -> Value:
    count = context.integer(size)
    if count < 1:
        return VALUE
    return Array([[1.0 if row == column else 0.0 for column in range(count)] for row in range(count)])


__all__ = ["numeric"]
