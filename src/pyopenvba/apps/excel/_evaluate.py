"""Evaluate and [...]: an expression worked out as Excel's Evaluate works it out.

Measured from VBA (scripts/measure_evaluate.py, tests/fixtures/evaluate.json).
Application.Evaluate and the bracket form work on the active sheet,
Worksheet.Evaluate on its own. An expression that comes to cells -- a
reference, a name for cells, INDEX, OFFSET, INDIRECT, CHOOSE or IF that
lands on cells, an intersection or a union -- is a Range. Anything else
is a value, an array counted from 1, or an error value; Evaluate hands an
error back rather than raising it. Blocks are worked out whole, as an
array formula works them: A1:A3*2 is three numbers. A last sum that all
but cancels is 0, as in a cell. An expression Excel cannot read, an
empty one, or one longer than 255 characters is Error 2015.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._engine import Context, evaluate_formula
from pyopenvba.formula._values import ExcelError
from pyopenvba.interpreter._values import VBAErrorValue

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: The longest expression Evaluate reads.
LONGEST = 255
#: What Evaluate answers for an expression it cannot read: #VALUE!, by its CVErr number.
UNREADABLE = 2015


def evaluated(sheet: Worksheet, text: str) -> object:
    """What Evaluate answers for ``text``, with ``sheet`` the one its references are on."""
    from pyopenvba.apps.excel._calc import as_vba
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._worksheet_functions import answer

    if len(text) > LONGEST:
        return VBAErrorValue(UNREADABLE)
    if "[" in text:
        raise VBAUnsupportedError(f"Evaluate of {text!r}, which names another workbook, is not implemented")
    try:
        node = P.parse(text.strip().removeprefix("="))
    except P.FormulaError:
        return VBAErrorValue(UNREADABLE)
    if isinstance(node, P.Call) and node.name.upper() == "XLOOKUP":
        raise VBAUnsupportedError("Evaluate of XLOOKUP, which comes to a Range, is not implemented")
    context = Context(sheet.book.calculator, sheet.name, array=True)
    try:
        areas = context.areas_of(node)
    except ExcelError as failure:
        return as_vba(failure)
    if areas is not None:
        owner = sheet.book.sheet_named(areas[0].sheet)
        assert owner is not None
        return Range(owner, areas)
    try:
        value = evaluate_formula(node, context)
    except ExcelError as failure:
        value = failure
    return answer(value)
