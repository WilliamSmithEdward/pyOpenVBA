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
- A range of another open workbook is its cells as well, whichever
  workbook is active (scripts/measure_excel_everyday_3.py).
"""

from __future__ import annotations

import dataclasses
from collections.abc import Iterator
import math
from decimal import ROUND_HALF_UP, Decimal
from typing import TYPE_CHECKING, Any, Final

from pyopenvba.apps.excel._engine_book import EngineBook
from pyopenvba.apps.excel._worksheet_function_refusals import REFUSALS
from pyopenvba.formula._calc import functions as functions  # imported to register every function
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.nodes import Given, Missing, Node
from pyopenvba.formula._calc.values import Area, Array, Empty, Omitted, Reference, Scalar, Value
from pyopenvba.interpreter._objects import VBAObject
from pyopenvba.interpreter._values import (
    EMPTY,
    MISSING,
    NOTHING,
    NULL,
    VBAArray,
    VBACurrency,
    VBADate,
    VBAErrorValue,
    VBAInt,
    VBASingle,
    error,
    to_bool,
    to_number,
    to_text,
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._calc import Calculator
    from pyopenvba.apps.excel._model import Application, Worksheet
    from pyopenvba.apps.excel._pivot_reports import PivotReport
    from pyopenvba.formula._calc.evaluator import Context, TableShape

#: Declared String, yet they take a range of several cells, typed too (tests/fixtures/worksheet_function_types.json): the
#: places of the texts each joins, and TextJoin's delimiter.
_NOT_CONVERTED: Final = {"textjoin": (0, 2), "concat": (0,)}
#: Functions that refuse a range of separate areas, measured (tests/fixtures/worksheet_function_heldout.json has the areas kind).
_ONE_AREA: Final = frozenset({"ftest", "f_test", "arraytotext", "sort", "unique", "xlookup"})
#: A Null where a function needs a value: a blank cell's value.
BLANK: Final = object()
#: Functions whose third argument is one value, whatever it is given: a lookup's column or row, and Match's type.
_ONE_VALUE: Final = frozenset({"vlookup", "hlookup", "match"})
#: The number base functions, whose places take no blank (#NUM!) and no logical (#VALUE!), item by item.
_BASE_CONVERSIONS: Final = frozenset({"bin2dec", "bin2hex", "bin2oct", "dec2bin", "dec2hex", "dec2oct", "hex2bin", "hex2dec", "hex2oct", "oct2bin", "oct2dec", "oct2hex"})
_ONE_CELL_WHEN_LIFTED: Final = _BASE_CONVERSIONS | {"bitand", "bitor", "bitxor", "bitlshift", "bitrshift"}
_ARRAY_OF_ONE: Final = frozenset({"sort", "sortby", "unique", "filter"})
_SWITCHES: Final = {"sort": (3,), "unique": (1, 2)}
#: WorksheetFunction members the type library types As Variant; every other one answers a single value.
_VARIANT: Final = frozenset({
    "choose", "dget", "encodeurl", "fieldvalue", "filter", "filterxml", "frequency", "growth", "hlookup",
    "iferror", "ifna", "index", "linest", "logest", "lookup", "minverse", "mmult", "mode_mult", "munit",
    "randarray", "rtd", "sequence", "single", "sort", "sortby", "stockhistory", "transpose", "trend", "unique",
    "vlookup", "webservice", "xlookup",
})
#: WorksheetFunction members the type library types As String, which hand a number back as its text: Bin2Dec("1100")
#: is "12" there and 12 from Application (tests/fixtures/worksheet_functions_more.json).
_STRING: Final = frozenset({
    "arraytotext", "asc", "bahttext", "base", "bin2dec", "bin2hex", "bin2oct", "clean", "complex", "concat", "dbcs",
    "dec2bin", "dec2hex", "dec2oct", "dollar", "fixed", "hex2bin", "hex2dec", "hex2oct", "imabs", "imargument",
    "imconjugate", "imcos", "imcosh", "imcot", "imcsc", "imcsch", "imdiv", "imexp", "imln", "imlog10", "imlog2",
    "impower", "improduct", "imsec", "imsech", "imsin", "imsinh", "imsqrt", "imsub", "imsum", "imtan", "oct2bin",
    "oct2dec", "oct2hex", "phonetic", "proper", "replace", "replaceb", "rept", "roman", "substitute", "text",
    "textjoin", "trim", "unichar", "usdollar", "valuetotext",
})
#: Functions whose one-item answer stays an array given single values: RandArray(1, 1) is a one-item array, where
#: Transpose(5) is 5 (tests/fixtures/worksheet_functions_more.json).
_ALWAYS_ARRAY: Final = frozenset({"randarray", "munit"})
def engine_name(name: str) -> str | None:
    """The formula engine's name for a WorksheetFunction member, where the engine has the function."""
    from pyopenvba.formula._calc import registry

    upper = name.upper().replace("_", ".")
    return upper if upper in registry.FUNCTIONS else None


def call(application: Application, name: str, args: Any, named: Any, *, raising: bool, late: bool = False) -> object:
    """``name`` called with VBA ``args``: through WorksheetFunction when ``raising``, else through Application.

    The arguments are handed to the engine as they are, a Range as its cells and Empty as an argument left out,
    and worked out as an array formula works them out: nothing is cut to a row, as Round(Range("A1:A2"), 0) is
    an array."""
    from pyopenvba.formula._calc import registry
    from pyopenvba.formula._calc.evaluator import Context
    from pyopenvba.formula._calc.values import ExcelError as EngineError

    function = engine_name(name)
    assert function is not None
    from pyopenvba.apps.excel._model import Range

    sheet = application.active_book.active_sheet if application.active_book is not None else None
    if sheet is None:
        raise error(1004, "There is no workbook open")
    values = _laid_out(list(args), named or {})
    entry = registry.FUNCTIONS[function]
    given = list(values)
    lifted_year_frac = raising and name.lower() == "yearfrac" and len(values) == 3 and any(isinstance(value, VBAArray) for value in values)
    if raising and any(_is_text_cell(values[place]) for place in _SWITCHES.get(name.lower(), ()) if place < len(values)):
        # A cell of text where a switch goes fails item by item, so the answer is an array holding the error.
        return answer(Array([[CellError("#VALUE!")]]))
    if raising and name.lower() in _BASE_CONVERSIONS and any(isinstance(value, VBAArray) for value in values):
        values = _base_conversion_items(values)
    if raising and name.lower() in _ONE_CELL_WHEN_LIFTED:
        lifted = next((value for value in values if isinstance(value, VBAArray)), None)
        if lifted is not None and any(isinstance(value, Range) and (len(value.areas) > 1 or (not value.single and name.lower() in _BASE_CONVERSIONS)) for value in values):
            # No value to call with, whatever the array: separate areas, and for a number base several cells.
            if not late:
                raise error(13, "Type mismatch")
            shape = _given_array(lifted)
            return answer(Array([[CellError("#VALUE!")] * shape.width for _ in range(shape.height)]))
    if raising and not lifted_year_frac:
        values = _as_excel_reads(name, values, entry.minimum, late=late)
        # Measured at 1E+300, past what a whole number of digits can be; where it starts is not.
        if name.lower() == "round" and len(values) == 2 and isinstance(values[1], float) and abs(values[1]) >= 2 ** 31:
            return 0.0
    calculator = sheet.book.calculator
    books = _OpenBooks(calculator)
    nodes = () if lifted_year_frac else tuple(_node(value, books) for value in values)
    single = not any(isinstance(value, (Range, VBAArray)) for value in values)
    now = calculator.now()
    context = Context(books, sheet.name, 1, 1, array=True, today=now.date(), now=now)
    if lifted_year_frac:
        return _year_frac_lifted(given, entry, context, books, late=late)
    if not entry.minimum <= len(nodes) <= entry.maximum:
        # Too few or too many arguments is error 1004 through Application as well: Application.Xnpv(values, dates)
        # (tests/fixtures/worksheet_functions_more.json).
        raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
    try:
        found = registry.call(entry, context, nodes)
    except EngineError as failure:
        found = failure.error
    if raising and late and isinstance(found, CellError) and _text_cell_for_switch(name, values):
        # Excel's own conversion fails item by item, so the answer is an array (tests/fixtures/worksheet_function_types.json).
        found = Array([[found]])
    if (single and isinstance(found, Array) and found.height == found.width == 1
            and name.lower() not in _ALWAYS_ARRAY):
        # Given single values, a function whose answer is one item hands back the item: Transpose(5) is 5.
        found = found.rows[0][0]
    if raising and name.lower() == "index" and isinstance(values[0], (int, VBAInt)) and not isinstance(values[0], bool) \
            and isinstance(found, float):
        # A whole number indexed answers as a Long (tests/fixtures/worksheet_function_args.json).
        return VBAInt(int(found), "Long")
    if raising and name.lower() == "xlookup" and isinstance(found, Reference) and found.area is not None and found.area.is_cell:
        # An error in the cell it answers is its answer, not a failure.
        return answer(context.book.cell(found.area.sheet, found.area.top, found.area.left))
    if raising and name.lower() in _ARRAY_OF_ONE:
        return _array_of_one(context, found, values[0], name, books_of=application, late=late)
    return _returned(context, found, name, raising=raising, books_of=application, late=late)


def _base_conversion_items(values: list[object]) -> list[object]:
    """A number base function's arguments with an array among them: its places as what each item fails with."""
    from pyopenvba.apps.excel._model import Range

    if len(values) < 2:
        return values
    places = values[1]
    if places is EMPTY or (isinstance(places, Range) and places.single and places.vba_value() is EMPTY):
        return [values[0], VBAErrorValue(2036), *values[2:]]
    if isinstance(places, bool):
        return [values[0], VBAErrorValue(2015), *values[2:]]
    if isinstance(places, VBAArray):
        items = [VBAErrorValue(2036) if one is EMPTY else VBAErrorValue(2015) if isinstance(one, bool) else one for one in places.elements()]
        return [values[0], VBAArray(places.bounds, items=items), *values[2:]]
    return values


def _year_frac_item(values: list[Any], context: Context) -> Scalar | None:
    """What one call of YearFrac with an array among its arguments fails with before it is worked out, looking at the basis, then
    each date: a logical or several cells #VALUE!, Null among the dates #N/A, a number out of range #NUM!."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.formula._calc.values import ExcelError as EngineError

    start, stop, kind = values
    if isinstance(kind, (VBAErrorValue, CellError)):
        return _given_item(kind) if isinstance(kind, VBAErrorValue) else kind
    if isinstance(kind, bool) or (isinstance(kind, Range) and not kind.single):
        return "#VALUE!"
    if isinstance(kind, Range):
        kind = kind.vba_value()
    if isinstance(kind, str):
        # A text is read as a number or a date.
        try:
            kind = context.number(kind)
        except EngineError:
            return "#VALUE!"
    if isinstance(kind, (int, float, VBAInt)) and not 0 <= math.trunc(float(kind)) <= 4:
        return "#NUM!"
    for date in (start, stop):
        if isinstance(date, Range) and date.single:
            date = date.vba_value()
        if isinstance(date, (VBAErrorValue, CellError)):
            return _given_item(date) if isinstance(date, VBAErrorValue) else date
        if date is NULL:
            return "#N/A"
        if isinstance(date, bool) or (isinstance(date, Range) and not date.single):
            return "#VALUE!"
        if isinstance(date, str):
            try:
                date = context.number(date)
            except EngineError:
                return "#VALUE!"
        if isinstance(date, (int, float, VBAInt)) and not 0 <= float(date) < 2958466:
            return "#NUM!"
    return None


def _year_frac_lifted(values: list[object], entry: Any, context: Context, books: _OpenBooks, *, late: bool) -> object:
    """YearFrac with an array among its arguments, call by call."""
    from pyopenvba.formula._calc import registry
    from pyopenvba.formula._calc.values import ExcelError as EngineError

    if not late:
        raise error(13, "Type mismatch")
    arrays = {place: _given_array(value) for place, value in enumerate(values) if isinstance(value, VBAArray)}
    height, width = max(one.height for one in arrays.values()), max(one.width for one in arrays.values())
    rows: list[list[Scalar]] = []
    for row in range(height):
        line: list[Scalar] = []
        for column in range(width):
            items: list[Any] = [arrays[place].at(row, column) if place in arrays else value for place, value in enumerate(values)]
            failed = _year_frac_item(items, context)
            if failed is not None:
                line.append(CellError(failed) if isinstance(failed, str) else failed)
                continue
            nodes = tuple(Given(item) if place in arrays else Missing() if item is NULL else _node(item, books) for place, item in enumerate(items))
            try:
                found = registry.call(entry, context, nodes)
            except EngineError as failure:
                found = failure.error
            line.append(context.first(found))
        rows.append(line)
    return answer(Array(rows))


def _is_text_cell(value: object) -> bool:
    from pyopenvba.apps.excel._model import Range

    return isinstance(value, Range) and value.single and isinstance(value.vba_value(), str)


def _array_of_one(context: Context, found: Value, given: object, name: str, *, books_of: Application, late: bool) -> object:
    """Sort and its kin answer an array even for one item, and hand back a whole number that fits an Integer as a Long."""
    from pyopenvba.apps.excel._model import Range

    if isinstance(found, Reference) and found.area is not None and found.area.is_cell:
        found = context.book.cell(found.area.sheet, found.area.top, found.area.left)
    if isinstance(found, Array) and found.height == found.width == 1 and not isinstance(given, (Range, VBAArray)):
        found = found.rows[0][0]
    if isinstance(found, (Array, Reference, CellError)):
        return _returned(context, found, name, raising=True, books_of=books_of, late=late)
    if isinstance(given, VBAInt) and -32768 <= int(given) <= 32767:
        return VBAArray([(1, 1)], items=[VBAInt(int(given), "Long")])
    return answer(Array([[found]]))


def _first_item(value: object) -> object:
    """The first cell's value of a range, or an array's first item."""
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        first = value.first
        from pyopenvba._a1 import Area as SheetArea

        return Range(value.sheet, [SheetArea(first.top, first.left, first.top, first.left, value.sheet.name)]).vba_value()
    if isinstance(value, VBAArray):
        return next(iter(_given_array(value).rows[0]))
    return value


def _typed(value: object, declared: str) -> object:
    """An argument converted as VBA converts it for a parameter declared other than Variant, before Excel sees it."""
    from pyopenvba.apps.excel._model import Range

    if value is MISSING:
        return value
    if declared == "Range":
        if isinstance(value, Range):
            return value
        if value is NOTHING:
            raise error(5)
        raise error(13 if isinstance(value, VBAObject) else 424)
    if value is NOTHING:
        raise error(91)
    if isinstance(value, Range):
        # Its value, as VBA reads it: several cells of one area are an array, and of several areas the first area's
        # (tests/fixtures/worksheet_function_heldout.json).
        first = value.areas[0]
        if first.rows * first.columns > 1:
            raise error(13)
        value = Range(value.sheet, [first]).vba_value()
    elif isinstance(value, VBAObject):
        raise error(438)
    if isinstance(value, (VBAArray, VBAErrorValue)):
        raise error(13)
    if value is NULL:
        raise error(94)
    if declared == "Double":
        return float(to_number(value))
    return to_text(value) if declared == "String" else bool(to_bool(value))


def _kind_of(value: object) -> str:
    """The kind of value, as _worksheet_function_refusals has them."""
    from pyopenvba.apps.excel._model import Range

    if isinstance(value, Range):
        if not value.single:
            return "cells"
        return "empty cell" if value.vba_value() is EMPTY else "cell"
    if isinstance(value, VBAArray):
        return "array"
    if value is NULL:
        return "null"
    if value is EMPTY:
        return "empty"
    if isinstance(value, bool):
        return "true" if value else "false"
    return "other"


def _as_excel_reads(name: str, values: list[object], minimum: int, *, late: bool) -> list[object]:
    """What WorksheetFunction gives the engine (tests/fixtures/worksheet_function_args.json). Typed, VBA converts first
    for what the library declares. Then Nothing is 5 and an object that is no range 1004; Null counts as left out, but
    for the last argument it cannot do without; and the column of a lookup and Match's type read one value."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.interpreter._type_library import find_class

    library = find_class("excel", "worksheetfunction")
    forms = library.members.get(name.lower(), []) if library is not None else []
    declared = [one.split(":")[-1] for one in forms[0][4]] if forms else []
    found: list[object] = []
    # With an array among them Excel calls once per item, so what would be refused is that item's answer.
    lifting = any(isinstance(value, VBAArray) for value in values)
    for index, value in enumerate(values):
        kind = declared[index] if index < len(declared) else "Variant"
        if not late and kind != "Variant" and index not in _NOT_CONVERTED.get(name.lower(), ()):
            value = _typed(value, kind)
        if value is NOTHING:
            raise error(5)
        if isinstance(value, VBAObject) and not isinstance(value, Range):
            raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        if _kind_of(value) in REFUSALS.get(name.lower(), {}).get(index, ()) and not lifting:
            raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        lower = name.lower()
        if lower in _ONE_AREA and isinstance(value, Range) and len(value.areas) > 1:
            raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        if lower.startswith("im") and isinstance(value, (VBADate, VBACurrency)):
            # The complex number functions read a Date or Currency as its number, where others read its text.
            value = float(to_number(value))
        if value is NULL:
            # The last it cannot do without, with nothing after it, or a column; one it needs before that is a blank
            # cell, and one it does not is left out (tests/fixtures/worksheet_function_types.json).
            if (index == minimum - 1 and index == len(values) - 1) or (not lifting and ((lower == "index" and index in (0, 2))
                                                                                        or (lower in ("vlookup", "hlookup") and index == 2))):
                raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
            value = BLANK if index < minimum else MISSING
        if lower == "xlookup" and value is EMPTY:
            # A fallback given as Empty is one: nothing found answers Empty, where one left out is an error. A search
            # mode given as Empty is 0, which is no mode.
            value = BLANK if index == 3 else 0.0 if index == 5 else value
        if lower == "yearfrac" and isinstance(value, VBADate):
            # Its number, which before March 1900 is a day from what its text would read as.
            value = float(to_number(value))
        if lower in _ONE_VALUE and index == 2:
            if lower != "match" or isinstance(value, Range):
                value = _first_item(value)
            elif isinstance(value, VBAArray) and not late:
                raise error(13)
        found.append(value)
    if name.lower() == "index" and found and found[0] is EMPTY:
        raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
    # Left out at the end, it is no argument at all: Count(1, Null) is 1.
    while found and found[-1] is MISSING:
        found.pop()
    if name.lower() == "unique" and len(found) == 1 and values[0] is NULL:
        raise error(1004, f"Unable to get the {name} property of the WorksheetFunction class")
    return found


def _text_cell_for_switch(name: str, values: list[object]) -> bool:
    """Whether one cell is given where the library declares a Boolean."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.interpreter._type_library import find_class

    library = find_class("excel", "worksheetfunction")
    forms = library.members.get(name.lower(), []) if library is not None else []
    declared = [one.split(":")[-1] for one in forms[0][4]] if forms else []
    return any(isinstance(value, Range) and value.single and index < len(declared) and declared[index] == "Boolean"
               for index, value in enumerate(values))


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


def _node(value: object, books: _OpenBooks) -> Node:
    """A VBA argument as the engine takes it: Empty as an argument left out, a Range as its cells."""
    from pyopenvba.apps.excel._model import Range

    if value is MISSING or value is EMPTY:
        return Missing()
    if value is BLANK:
        return Given(_given_item(EMPTY))
    if value is NULL:
        raise error(1004, "A worksheet function cannot take Null")
    if isinstance(value, Range):
        sheet = books.key(value.sheet)
        return Given(Reference(tuple(Area(sheet, one.top, one.left, one.bottom, one.right) for one in value.areas)))
    if isinstance(value, VBAArray):
        return Given(_given_array(value))
    return Given(_given_item(value))


def _given_item(value: object) -> Scalar:
    """One VBA value as the engine computes with it: a Date its serial, a Currency or Decimal its number."""
    from pyopenvba.apps.excel._engine_book import scalar

    if value is EMPTY or isinstance(value, (bool, str, VBAErrorValue)):
        return scalar(value)
    if isinstance(value, VBACurrency):
        # A Currency reaches Excel as its text, in the currency format: $4.00 with a space after it, ($4.00) below nothing,
        # to the nearest cent, half away from nothing (tests/fixtures/worksheet_function_heldout.json).
        cents = abs(Decimal(value)).quantize(Decimal("0.01"), ROUND_HALF_UP)
        return f"(${cents:,.2f})" if value < 0 else f"${cents:,.2f} "
    if isinstance(value, VBADate):
        # A Date reaches Excel as its text: IsText(Date) is TRUE, AND passes it over as it passes text over, and SUM
        # reads the number back (native probes).
        return to_text(value)
    if isinstance(value, (Decimal, VBAInt, VBASingle, int, float)):
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


class _OpenBooks(EngineBook):
    """The active workbook as the engine reads it, along with the sheets of other open workbooks a macro hands
    ranges of. Such a sheet goes by a key no sheet's name can be, so two workbooks' Sheet1 stay apart, and is read
    from its own workbook, worked out by that workbook's calculator."""

    def __init__(self, calculator: Calculator) -> None:
        super().__init__(calculator)
        self._others: dict[str, tuple[EngineBook, str]] = {}
        #: The sheet behind each key, for an answer that is a range.
        self.other_sheets: dict[str, Any] = {}

    def key(self, sheet: Worksheet) -> str:
        """The name the engine reads a sheet by: its own in the active workbook, a key in another."""
        if sheet.book is self.book:
            return sheet.name
        key = f"\0{id(sheet.book)}\0{sheet.name}"
        self._others[key] = (sheet.book.calculator.engine_book, sheet.name)
        self.other_sheets[key] = sheet
        return key

    def _routed(self, sheet: str) -> tuple[EngineBook, str]:
        """The engine book a sheet is read through, and its name there."""
        return self._others.get(sheet) or (self.calculator.engine_book, sheet)

    def cell(self, sheet: str, row: int, column: int) -> Scalar:
        book, name = self._routed(sheet)
        return book.cell(name, row, column)

    def cells(self, area: Area) -> Iterator[tuple[int, int, Scalar]]:
        book, name = self._routed(area.sheet)
        return book.cells(dataclasses.replace(area, sheet=name))

    def used(self, sheet: str) -> tuple[int, int]:
        book, name = self._routed(sheet)
        return book.used(name)

    def row_hidden(self, sheet: str, row: int) -> bool:
        book, name = self._routed(sheet)
        return book.row_hidden(name, row)

    def row_filtered(self, sheet: str, row: int) -> bool:
        book, name = self._routed(sheet)
        return book.row_filtered(name, row)

    def subtotal(self, sheet: str, row: int, column: int) -> bool:
        book, name = self._routed(sheet)
        return book.subtotal(name, row, column)

    def formula_text(self, sheet: str, row: int, column: int) -> str | None:
        book, name = self._routed(sheet)
        return book.formula_text(name, row, column)

    def pivot_reports(self, sheet: str) -> list[PivotReport]:
        book, name = self._routed(sheet)
        return book.pivot_reports(name)

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None:
        book, name = self._routed(sheet)
        found = book.table_at(name, row, column)
        return None if found is None else dataclasses.replace(found, sheet=sheet)

    def spill(self, sheet: str, row: int, column: int) -> Area | None:
        book, name = self._routed(sheet)
        found = book.spill(name, row, column)
        return None if found is None else dataclasses.replace(found, sheet=sheet)


def _returned(context: Context, found: Value, name: str, *, raising: bool, books_of: Application, late: bool = False) -> object:
    """The engine's answer as VBA reads it: an error raised by WorksheetFunction and handed back by Application,
    part of a range as Range.Value reads it, and an array counted from 1."""
    whole = False
    if raising and isinstance(found, Reference) and found.area is not None and name.lower() in ("index", "choose"):
        # WorksheetFunction.Index and Choose answer the range itself: f.Index(r, 0, 1).Address (tests/fixtures/excel_everyday_15.json).
        return _range_of(books_of, found, getattr(context.book, 'other_sheets', {}))
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
    # Called through an Object nothing holds it to its signature: Match(Array(1, 2), r, 0) is an array and Bin2Dec a number, where
    # typed they are error 13 and text (tests/fixtures/excel_everyday_15.json, tests/fixtures/worksheet_functions_more.json).
    if isinstance(found, Array) and raising and not late and name.lower() not in _VARIANT:
        raise error(13, "Type mismatch")
    if raising and not late and isinstance(found, float) and name.lower() in _STRING:
        return to_text(found)
    return answer(found, whole=whole)


def _range_of(application: Application, found: Reference, other_sheets: dict[str, Any]) -> object:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Range

    area = found.area
    assert area is not None
    held = other_sheets.get(area.sheet)
    if held is not None:
        return Range(held, [Area(area.top, area.left, area.bottom, area.right, held.name)])
    book = application.active_book
    assert book is not None
    for sheet in book.sheets_:
        if sheet.name.casefold() == area.sheet.casefold():
            return Range(sheet, [Area(area.top, area.left, area.bottom, area.right, sheet.name)])
    from pyopenvba.exceptions import VBAUnsupportedError

    raise VBAUnsupportedError("WorksheetFunction.Index of a range in another workbook is not implemented")


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
