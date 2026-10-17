"""A VBA Function called from a cell's formula: a user-defined function.

The formula engine hands a name it does not know to the workbook
(``EngineBook.user_function``); a Function of that name in a standard module
of the project, Private or not and in any case, is called here, as Excel
calls one (tests/fixtures/excel_vba_udfs.json, native probes):

- a reference is passed as a Range, a number as a Double, text, a Boolean,
  an error as an Error value, an array as a two-dimensional Variant array
  from 1 (one row a one-dimensional array), and an argument left out as
  missing;
- what comes back goes into the cell: a Range as its value, an array as an
  array (its corner in a legacy formula), an Error value as that error;
- a run-time error the function does not handle makes the cell #VALUE!;
- while it runs, Application.Caller is the cell and a write to a cell is
  error 1004; it is called as any procedure is, preserving the current Err and
  leaving what it ends with in Err, and an error it does not handle leaves
  40040 there. A formula that calls one is worked out as it is written.
"""

from __future__ import annotations

from collections.abc import Generator
from contextlib import contextmanager
from decimal import Decimal
from typing import TYPE_CHECKING, Any

from pyopenvba.apps.excel._engine_book import ERROR_NUMBERS
from pyopenvba.exceptions import VBARuntimeError
from pyopenvba.formula._calc.cells import CellError
from pyopenvba.formula._calc.nodes import Call, Missing, Node, walk
from pyopenvba.formula._calc.values import EMPTY as ENGINE_EMPTY
from pyopenvba.formula._calc.values import VALUE, Array, Empty, Reference, Scalar, Value
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
)

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Application, Range, Workbook
    from pyopenvba.formula._calc.evaluator import Context


def has_user_function(book: Workbook, node: Node | None) -> bool:
    """Unknown formula names alone do not require manual-mode UDF entry evaluation."""
    from pyopenvba.apps.excel._projects import attach_project

    interpreter = attach_project(book)
    if interpreter is None or node is None:
        return False
    _initialise(book)
    for entry in walk(node):
        if isinstance(entry, Call) and '.' not in entry.name:
            procedure, runtime = interpreter.find_procedure(entry.name)
            if procedure is not None and runtime is not None and procedure.kind == 'function':
                return True
    return False


def call_user_function(book: Workbook, context: Context, name: str, nodes: tuple[Node, ...]) -> Value | None:
    """The value of ``name(nodes)`` worked out by the project's Function of that name, or None without one."""
    from pyopenvba.apps.excel._projects import attach_project

    interpreter = attach_project(book)
    if interpreter is None or "." in name:
        return None
    _initialise(book)
    try:
        with interpreter.udf_bindings(), running(book.application, _cell(book, context)):
            procedure, runtime = interpreter.find_procedure(name)
            if procedure is None or runtime is None or procedure.kind != "function":
                return None
            args = [_argument(book, context, node) for node in nodes]
            # Like an ordinary call, it preserves Err and leaves its final state visible to the macro.
            with book.calculator.user_call():
                result = interpreter.call(procedure, runtime, args, {})
            book.calculator.after_udf()
    except VBARuntimeError:
        # An error the function does not handle leaves 40040 in the macro's Err, and the cell #VALUE!.
        interpreter.err.number = 40040
        interpreter.err.description = "Application-defined or object-defined error"
        book.calculator.after_udf(error=True)
        return VALUE
    return _value(result)


def _initialise(book: Workbook) -> None:
    from pyopenvba.apps.excel._projects import attach_project
    from pyopenvba.apps.excel._vbide import VBProject

    interpreter = attach_project(book)
    assert interpreter is not None
    if isinstance(book.vbide_project, VBProject):
        book.vbide_project.compile_pending(for_udf=True)
    interpreter.initialise()


@contextmanager
def running(application: Application, caller: Range) -> Generator[None]:
    """A user-defined function at work for ``caller``: what Application.Caller answers, and no cell writes."""
    application.udf_callers.append(caller)
    try:
        yield
    finally:
        application.udf_callers.pop()


def _cell(book: Workbook, context: Context) -> Range:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Range

    sheet = book.sheet_named(context.sheet)
    return Range(sheet, [Area(context.row, context.column, context.row, context.column, sheet.name)])


def _argument(book: Workbook, context: Context, node: Node) -> object:
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Range

    if isinstance(node, Missing):
        return MISSING
    value = context.evaluate(node)
    if isinstance(value, Reference):
        first = book.sheet_named(value.areas[0].sheet)
        return Range(first, [Area(area.top, area.left, area.bottom, area.right, first.name) for area in value.areas])
    if isinstance(value, Array):
        if value.height == 1:
            # One row is a one-dimensional array from 1: {1,2,3} is (1 To 3).
            return VBAArray([(1, value.width)], items=[_vba(value.at(0, column)) for column in range(value.width)])
        items: list[object] = []
        # VBA lays a two-dimensional array out column by column.
        for column in range(value.width):
            for row in range(value.height):
                items.append(_vba(value.at(row, column)))
        return VBAArray([(1, value.height), (1, value.width)], items=items)
    return _vba(value)


def _vba(value: Scalar) -> object:
    if isinstance(value, Empty):
        return EMPTY
    if isinstance(value, CellError):
        return VBAErrorValue(ERROR_NUMBERS.get(value.code, 2015))
    return value


def _value(result: object) -> Value:
    from pyopenvba.apps.excel._model import Range

    if isinstance(result, Range):
        result = result.vba_get("Value")
    if isinstance(result, VBAArray):
        if result.dimensions == 1:
            return Array([[_item(item) for item in result.elements()]])
        if result.dimensions == 2:
            (low, high), (left, right) = result.bounds
            return Array([[_item(result.get([row, column])) for column in range(left, right + 1)]
                          for row in range(low, high + 1)])
        return VALUE
    return _item(result)


def _item(value: Any) -> Scalar:
    from pyopenvba.apps.excel._engine_book import scalar

    if value is EMPTY or value is NULL or value is MISSING:
        return ENGINE_EMPTY
    if isinstance(value, VBADate):
        return value.serial
    if isinstance(value, (VBACurrency, Decimal, VBAInt, VBASingle)):
        return float(value)
    return scalar(value)
