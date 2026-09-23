"""The number format a formula brings to the cell it is written to.

Measured in live Excel (tests/fixtures/excel_model/probes.txt, the
section on the number format a formula takes). A formula written to a
General cell gives it a format once, as it is written, and never again:
a source formatted later changes nothing. The format is worked out from
the formula as written at the top left of the range and given to every
cell of the write, so Range("C1:C2").Formula = "=A1+1" dates both cells
when A1 is a date, whatever A2 is.

What a formula brings follows its shape:

* a reference brings its top-left cell's format, blank or not;
* ``+`` and ``-`` bring the first side's format, or the second's when
  the first has none, except that two dates bring none: =A1-A1 is a
  number of days. A unary sign or brackets keep what is inside;
* ``*``, ``/``, ``^``, ``&`` and comparisons bring nothing, even from a
  percentage;
* DATE and TODAY bring m/d/yyyy, NOW m/d/yyyy h:mm and TIME h:mm AM/PM;
* SUM, MAX, MIN, INT, ROUND, ROUNDDOWN, ROUNDUP, TRUNC and MOD hand on
  their first argument's format that there is one: SUM(5,A1) is a date
  when A1 is. Every other function measured brings nothing -- AVERAGE,
  ABS, MEDIAN, SUMIF, IF, INDEX and VLOOKUP among them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Final

from pyopenvba.apps.excel._typing import is_date_format
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.evaluator import Context
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import (
    AreaReference,
    AxisReference,
    Binary,
    Call,
    CellReference,
    NameReference,
    Node,
    Paren,
    Unary,
)
from pyopenvba.formula._calc.values import ExcelError, Reference

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: Functions whose answer brings a format of its own.
_BRINGS: Final = {"DATE": "m/d/yyyy", "TODAY": "m/d/yyyy", "NOW": "m/d/yyyy h:mm", "TIME": "h:mm AM/PM"}
#: Functions that hand on the format of their first argument that has one.
_HANDS_ON: Final = frozenset({"SUM", "MAX", "MIN", "INT", "ROUND", "ROUNDDOWN", "ROUNDUP", "TRUNC", "MOD"})
#: What names cells as it is written: a reference or a defined name.
_CELLS: Final = (CellReference, AreaReference, AxisReference, NameReference)


def brought_format(sheet: Worksheet, formula: str, row: int, column: int) -> str:
    """The format a formula written at ``row``, ``column`` of ``sheet`` brings, or "" for none."""
    calculator = sheet.book.calculator
    try:
        node = calculator.engine_book.read("=" + formula.removeprefix("="))
    except FormulaSyntaxError:
        return ""
    now = calculator.now()
    context = Context(calculator.engine_book, sheet.name, row, column, today=now.date(), now=now)
    return _format_of(node, context, sheet)


def _format_of(node: Node, context: Context, sheet: Worksheet) -> str:
    if isinstance(node, _CELLS) or (isinstance(node, Binary) and node.op in (":", " ", ",")):
        return _cells_format(node, context, sheet)
    if isinstance(node, Paren):
        return _format_of(node.inner, context, sheet)
    if isinstance(node, Unary):
        return _format_of(node.operand, context, sheet) if node.op in ("+", "-") else ""
    if isinstance(node, Binary):
        if node.op not in ("+", "-"):
            return ""
        left = _format_of(node.left, context, sheet)
        right = _format_of(node.right, context, sheet)
        if left and right and is_date_format(left) and is_date_format(right):
            return ""
        return left or right
    if isinstance(node, Call):
        name = node.function
        if name in _BRINGS:
            return _BRINGS[name]
        if name in _HANDS_ON:
            return next((code for argument in node.args if (code := _format_of(argument, context, sheet))), "")
    return ""


def _cells_format(node: Node, context: Context, sheet: Worksheet) -> str:
    """The format of the top-left cell a reference names, "" for General or for no cells.

    A name the engine cannot work out, one with a relative reference, brings nothing, and the formula is still
    written: reading the cell reports it."""
    try:
        value = context.evaluate(node)
    except (ExcelError, VBAUnsupportedError):
        return ""
    if not isinstance(value, Reference):
        return ""
    first = value.areas[0]
    code = sheet.book.sheet_named(first.sheet).style_at(first.top, first.left).number_format
    return "" if code in ("General", "") else code
