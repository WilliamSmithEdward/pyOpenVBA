"""Worksheet functions a macro calls: WorksheetFunction.X raises an error, Application.X hands it back.

Both go through the formula engine, so a function the engine has behaves
in VBA as it does in a cell. Measured in live Excel
(scripts/measure_worksheet_functions.py), every call made both ways:

- A Range argument is its cells. A VBA array is a row when it has one
  dimension and rows when it has two, and an array of arrays is rows as
  well. A Date is its serial number and a Currency or Decimal its number.
  Empty is an argument left out, which counts as 0 or "" as the function
  needs it; Null is error 1004 either way.
- A number comes back a Double, text a String, TRUE and FALSE Booleans,
  and a blank cell Empty. An array comes back a VBA array counted from
  1, one-dimensional when it is one row and two-dimensional otherwise,
  except that part of a range, which INDEX, CHOOSE and XLOOKUP answer
  with when given ranges, reads back two-dimensional as Range.Value does.
- An error answer is error 1004 from WorksheetFunction and an error value
  from Application. A WorksheetFunction member the type library types
  other than Variant cannot hand back an array: that is error 13.
"""

from __future__ import annotations

from decimal import Decimal
from typing import TYPE_CHECKING, Any, Final

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._values import BLANK, ExcelError, Matrix
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBAErrorValue,
    VBAInt,
    VBASingle,
    error,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application

#: WorksheetFunction members the type library types As Variant; every other one answers a single value.
_VARIANT: Final = frozenset({
    "choose", "dget", "encodeurl", "fieldvalue", "filter", "filterxml", "frequency", "growth", "hlookup",
    "iferror", "ifna", "index", "linest", "logest", "lookup", "minverse", "mmult", "mode_mult", "munit",
    "randarray", "rtd", "sequence", "single", "sort", "sortby", "stockhistory", "transpose", "trend", "unique",
    "vlookup", "webservice", "xlookup",
})
#: Functions that answer with part of a range when handed one, which reads back as Range.Value reads.
_REFERENCES: Final = frozenset({"INDEX", "CHOOSE", "XLOOKUP"})


def engine_name(name: str) -> str | None:
    """The formula engine's name for a WorksheetFunction member, where the engine has the function."""
    from pyopenvba.formula._functions import FUNCTIONS

    upper = name.upper().replace("_", ".")
    return upper if upper in FUNCTIONS else None


def call(application: Application, name: str, args: Any, named: Any, *, raising: bool) -> object:
    """``name`` called with VBA ``args``: through WorksheetFunction when ``raising``, else through Application."""
    from pyopenvba.formula._engine import Context
    from pyopenvba.formula._functions import call as run

    function = engine_name(name)
    assert function is not None
    values = _laid_out(list(args), named or {})
    ranged = False
    converted: list[object] = []
    for value in values:
        ranged = ranged or _is_range(value)
        converted.append(_argument(value))
    sheet = application.active_book.active_sheet if application.active_book is not None else None
    if sheet is None:
        raise error(1004, "There is no workbook open")
    context = Context(sheet.book.calculator, sheet.name)
    try:
        answer = run(function, [P.Literal(value=one) for one in converted], context)
    except ExcelError as failure:
        answer = failure
    if isinstance(answer, ExcelError):
        if raising:
            raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        return _scalar(answer)
    if isinstance(answer, Matrix):
        if raising and name.lower() not in _VARIANT:
            raise error(13, "Type mismatch")
        return _array(answer, whole=ranged and function in _REFERENCES)
    return _scalar(answer)


def _laid_out(args: list[object], named: dict[str, object]) -> list[object]:
    """Positional arguments, with named ones (Arg1, Arg2...) put in their places and the ones left off the end gone."""
    for key, value in named.items():
        if not key.lower().startswith("arg") or not key[3:].isdigit():
            raise error(448, f"Named argument not found: {key}")
        index = int(key[3:]) - 1
        while len(args) <= index:
            args.append(MISSING)
        args[index] = value
    while args and args[-1] is MISSING:
        args.pop()
    return args


def _is_range(value: object) -> bool:
    from pyopenvba.apps.excel._model import Range

    return isinstance(value, Range)


def _argument(value: object) -> object:
    """A VBA value as the engine takes it."""
    from pyopenvba.apps.excel._model import Range

    if value is MISSING or value is EMPTY:
        return BLANK
    if value is NULL:
        raise error(1004, "A worksheet function cannot take Null")
    if isinstance(value, Range):
        if len(value.areas) != 1:
            raise VBAUnsupportedError("a worksheet function given a range of several areas is not implemented")
        return value.sheet.book.calculator.block(value.sheet.name, value.first)
    if isinstance(value, VBAArray):
        return _matrix(value)
    return _item(value)


def _item(value: object) -> object:
    from pyopenvba.apps.excel._calc import from_vba

    if value is EMPTY:
        return BLANK
    if isinstance(value, bool):
        return value
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (VBACurrency, Decimal, VBAInt, VBASingle, int, float)):
        return float(value)
    if isinstance(value, str):
        return value
    if isinstance(value, VBAErrorValue):
        return from_vba(value)
    raise error(13, "Type mismatch")


def _matrix(array: VBAArray) -> Matrix:
    items = array.elements()
    if array.dimensions == 1:
        if items and all(isinstance(item, VBAArray) and item.dimensions == 1 for item in items):
            # An array of arrays is rows.
            return Matrix([[_item(one) for one in row.elements()] for row in items])  # type: ignore[union-attr]
        return Matrix([[_item(one) for one in items]])
    if array.dimensions != 2:
        raise error(13, "Type mismatch")
    (low, high), (left, right) = array.bounds
    return Matrix([[_item(array.get([row, column])) for column in range(left, right + 1)]
                   for row in range(low, high + 1)])


def _scalar(value: object) -> object:
    from pyopenvba.apps.excel._calc import as_vba

    if value is BLANK:
        return EMPTY
    if isinstance(value, ExcelError):
        # CVErr's numbers, 2042 for #N/A, as a cell's error reads in VBA.
        return as_vba(value)
    if isinstance(value, bool) or isinstance(value, str):
        return value
    if isinstance(value, (int, float)):
        return float(value)
    return value


def answer(value: object) -> object:
    """An engine answer as VBA reads it back: an array counted from 1, one-dimensional when it is one row."""
    return _array(value, whole=False) if isinstance(value, Matrix) else _scalar(value)


def _array(matrix: Matrix, *, whole: bool) -> VBAArray:
    """An engine array as VBA reads it: one row one-dimensional, unless it is part of a range."""
    if matrix.height == 1 and not whole:
        return VBAArray([(1, matrix.width)], items=[_scalar(one) for one in matrix.rows[0]])
    # VBA lays a two-dimensional array out column by column.
    items = [_scalar(matrix.rows[row][column]) for column in range(matrix.width) for row in range(matrix.height)]
    return VBAArray([(1, matrix.height), (1, matrix.width)], items=items)
