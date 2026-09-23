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
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.nodes import Given, Missing, Node
from pyopenvba.formula._calc.values import Array, Empty, Omitted, Reference, Scalar, Value
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
    from pyopenvba.apps.excel._model import Application, Workbook
    from pyopenvba.formula._calc.evaluator import Context

#: WorksheetFunction members the type library types As Variant; every other one answers a single value.
_VARIANT: Final = frozenset({
    "choose", "dget", "encodeurl", "fieldvalue", "filter", "filterxml", "frequency", "growth", "hlookup",
    "iferror", "ifna", "index", "linest", "logest", "lookup", "minverse", "mmult", "mode_mult", "munit",
    "randarray", "rtd", "sequence", "single", "sort", "sortby", "stockhistory", "transpose", "trend", "unique",
    "vlookup", "webservice", "xlookup",
})
def engine_name(name: str) -> str | None:
    """The formula engine's name for a WorksheetFunction member, where the engine has the function."""
    from pyopenvba.formula._calc import registry

    upper = name.upper().replace("_", ".")
    return upper if upper in registry.FUNCTIONS else None


def call(application: Application, name: str, args: Any, named: Any, *, raising: bool) -> object:
    """``name`` called with VBA ``args``: through WorksheetFunction when ``raising``, else through Application.

    The arguments are handed to the engine as they are, a Range as its cells and Empty as an argument left out,
    and worked out as an array formula works them out: nothing is cut to a row, as Round(Range("A1:A2"), 0) is
    an array."""
    from pyopenvba.formula._calc import registry
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.values import VALUE, ExcelError as EngineError

    function = engine_name(name)
    assert function is not None
    from pyopenvba.apps.excel._model import Range

    sheet = application.active_book.active_sheet if application.active_book is not None else None
    if sheet is None:
        raise error(1004, "There is no workbook open")
    values = _laid_out(list(args), named or {})
    nodes = tuple(_node(value, sheet.book) for value in values)
    single = not any(isinstance(value, (Range, VBAArray)) for value in values)
    calculator = sheet.book.calculator
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, 1, 1, array=True, today=now.date(), now=now)
    entry = registry.FUNCTIONS[function]
    try:
        found = registry.call(entry, context, nodes) if entry.minimum <= len(nodes) <= entry.maximum else VALUE
    except EngineError as failure:
        found = failure.error
    if single and isinstance(found, Array) and found.height == found.width == 1:
        # Given single values, a function whose answer is one item hands back the item: Transpose(5) is 5.
        found = found.rows[0][0]
    return _returned(context, found, name, raising=raising)


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


def _node(value: object, book: Workbook) -> Node:
    """A VBA argument as the engine takes it: Empty as an argument left out, a Range as its cells."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.formula._calc.values import Area, Reference

    if value is MISSING or value is EMPTY:
        return Missing()
    if value is NULL:
        raise error(1004, "A worksheet function cannot take Null")
    if isinstance(value, Range):
        if value.sheet.book is not book:
            raise VBAUnsupportedError("a worksheet function given a range of another workbook is not implemented")
        return Given(Reference(tuple(Area(value.sheet.name, one.top, one.left, one.bottom, one.right)
                                     for one in value.areas)))
    if isinstance(value, VBAArray):
        return Given(_given_array(value))
    return Given(_given_item(value))


def _given_item(value: object) -> Scalar:
    """One VBA value as the engine computes with it: a Date its serial, a Currency or Decimal its number."""
    from pyopenvba.apps.excel._engine_book import scalar

    if value is EMPTY or isinstance(value, (bool, str, VBAErrorValue)):
        return scalar(value)
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (VBACurrency, Decimal, VBAInt, VBASingle, int, float)):
        return float(value)
    raise error(13, "Type mismatch")


def _given_array(array: VBAArray) -> Array:
    """A VBA array as the engine's: a row when it has one dimension, rows when it has two or holds arrays."""
    items = array.elements()
    if array.dimensions == 1:
        if items and all(isinstance(item, VBAArray) and item.dimensions == 1 for item in items):
            # An array of arrays is rows.
            return Array([[_given_item(one) for one in row.elements()] for row in items])  # type: ignore[union-attr]
        return Array([[_given_item(one) for one in items]])
    if array.dimensions != 2:
        raise error(13, "Type mismatch")
    (low, high), (left, right) = array.bounds
    return Array([[_given_item(array.get([row, column])) for column in range(left, right + 1)]
                  for row in range(low, high + 1)])


def _returned(context: Context, found: Value, name: str, *, raising: bool) -> object:
    """The engine's answer as VBA reads it: an error raised by WorksheetFunction and handed back by Application,
    part of a range as Range.Value reads it, and an array counted from 1."""
    whole = False
    if isinstance(found, Reference):
        if found.area is None:
            found = CellError("#VALUE!")
        elif found.area.is_cell:
            found = context.book.cell(found.area.sheet, found.area.top, found.area.left)
        else:
            # Part of a range, as INDEX, CHOOSE and XLOOKUP answer, reads back two-dimensional as Range.Value does.
            found, whole = context.array_of(found), True
    if isinstance(found, CellError) and raising:
        raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
    if isinstance(found, Array) and raising and name.lower() not in _VARIANT:
        raise error(13, "Type mismatch")
    return answer(found, whole=whole)


def answer(value: Scalar | Array, *, whole: bool = False) -> object:
    """An engine answer as VBA reads it back: an array counted from 1, one-dimensional when it is one row unless
    ``whole`` says it is part of a range; an error its CVErr value."""
    if not isinstance(value, Array):
        return _vba_item(value)
    rows = [[_vba_item(one) for one in row] for row in value.rows]
    if len(rows) == 1 and not whole:
        return VBAArray([(1, len(rows[0]))], items=list(rows[0]))
    # VBA lays a two-dimensional array out column by column.
    width = len(rows[0]) if rows else 0
    items = [rows[row][column] for column in range(width) for row in range(len(rows))]
    return VBAArray([(1, len(rows)), (1, width)], items=items)


def _vba_item(value: Scalar) -> object:
    """One item of an answer as VBA reads it: a blank Empty, an argument left out 0, an error its CVErr value."""
    from pyopenvba.apps.excel._calc import as_vba
    from pyopenvba.apps.excel._engine_book import model_value

    if isinstance(value, Omitted):
        # Choose(1, Empty) is 0 (tests/fixtures/worksheet_functions.json).
        return 0.0
    if isinstance(value, Empty):
        return EMPTY
    return as_vba(model_value(value))
