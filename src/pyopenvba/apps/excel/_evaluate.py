"""Evaluate and [...]: an expression worked out as Excel's Evaluate works it out.

Measured from VBA (scripts/measure_evaluate.py, tests/fixtures/evaluate.json).
Application.Evaluate and the bracket form work on the active sheet,
Worksheet.Evaluate on its own. An expression that comes to cells -- a
reference, a name for cells, INDEX, OFFSET, INDIRECT, CHOOSE, IF, IFS,
SWITCH or XLOOKUP that lands on cells, an intersection or a union -- is
a Range. Anything else is a value, an array counted from 1, or an error
value; Evaluate hands an error back rather than raising it. Blocks are
worked out whole, as an array formula works them: A1:A3*2 is three
numbers. A last sum that all but cancels is 0, as in a cell. An
expression Excel cannot read, an empty one, or one longer than 255
characters is Error 2015.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._engine_book import EngineBook
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula import _parse as P
from pyopenvba.formula._calc.evaluator import TableShape
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.values import ExcelError, Reference
from pyopenvba.interpreter._values import VBAErrorValue

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet

#: The longest expression Evaluate reads.
LONGEST = 255
#: What Evaluate answers for an expression it cannot read: #VALUE!, by its CVErr number.
UNREADABLE = 2015


def evaluated(sheet: Worksheet, text: str) -> object:
    """What Evaluate answers for ``text``, with ``sheet`` the one its references are on."""
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._worksheet_functions import answer
    from pyopenvba.formula._calc.evaluator import Context

    if len(text) > LONGEST:
        return VBAErrorValue(UNREADABLE)
    body = text.strip().removeprefix("=")
    try:
        tokens = P.tokenize(body)
    except P.FormulaError:
        return VBAErrorValue(UNREADABLE)
    if any(token.kind == "structured" and following.kind in ("ref", "name", "structured")
           and following.at == token.at + len(token.text) for token, following in zip(tokens, tokens[1:])):
        raise VBAUnsupportedError(f"Evaluate of {text!r}, which names another workbook, is not implemented")
    calculator = sheet.book.calculator
    try:
        node = calculator.engine_book.read("=" + body)
    except FormulaSyntaxError:
        return VBAErrorValue(UNREADABLE)
    now = calculator.now()
    context = Context(_FromNoCell(calculator), sheet.name, 1, 1, array=True, today=now.date(), now=now)
    try:
        value = context.formula(node)
    except ExcelError as failure:
        value = failure.error
    if isinstance(value, Reference):
        return _range(sheet, value) if value.areas else Range(sheet, [])
    return answer(value)


class _FromNoCell(EngineBook):
    """The workbook as Evaluate reads it: from no cell, so a column named without its table has no table to be in,
    and Evaluate("[Qty]") is an error (tests/fixtures/structured_references/)."""

    def table_at(self, sheet: str, row: int, column: int) -> TableShape | None:
        return None


def _range(sheet: Worksheet, reference: Reference) -> object:
    """The cells an expression came to, as a Range on the sheet they are on."""
    from pyopenvba.apps.excel._model import Range

    owner = sheet.book.sheet_named(reference.areas[0].sheet)
    if any(area.sheet != owner.name for area in reference.areas):
        raise VBAUnsupportedError("Evaluate of cells on more than one sheet is not implemented")
    return Range(owner, [Area(area.top, area.left, area.bottom, area.right, owner.name) for area in reference.areas])
