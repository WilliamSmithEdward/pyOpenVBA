"""An Excel formula engine.

Parses a formula, evaluates it against anything that can answer for a
grid, and reports an error the way a cell does.

    >>> from pyopenvba.formula import parse, evaluate
    >>> from pyopenvba.formula._engine import Context
    >>> node = parse("=1 + 2 * 3")
    >>> evaluate(node, Context(grid=None, sheet=""))  # doctest: +SKIP
    7.0

The workbook's own wiring -- which cells are stale, what order to
recalculate them in, and where a cell's value comes from -- lives in
:mod:`pyopenvba.apps.excel._calc`.
"""

from pyopenvba.formula._engine import Context, Grid, evaluate
from pyopenvba.formula._parse import FormulaError, Node, is_volatile, parse, references
from pyopenvba.formula._values import (
    BLANK,
    DIV0,
    ERRORS,
    NA,
    NAME,
    NUM,
    REF,
    VALUE,
    ExcelError,
    Matrix,
    as_bool,
    as_number,
    as_text,
    number_text,
)

__all__ = [
    "BLANK",
    "DIV0",
    "ERRORS",
    "NA",
    "NAME",
    "NUM",
    "REF",
    "VALUE",
    "Context",
    "ExcelError",
    "FormulaError",
    "Grid",
    "Matrix",
    "Node",
    "as_bool",
    "as_number",
    "as_text",
    "evaluate",
    "is_volatile",
    "number_text",
    "parse",
    "references",
]
