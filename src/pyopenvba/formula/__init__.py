"""Excel formulas: reading one, and the values and errors a cell holds.

    >>> from pyopenvba.formula import parse
    >>> parse("=1 + 2 * 3")  # doctest: +SKIP

A workbook's formulas are worked out by :mod:`pyopenvba.formula._calc`,
and the wiring around it -- which cells are stale, and where a cell's
value comes from -- lives in :mod:`pyopenvba.apps.excel._calc`.
"""

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
    "ExcelError",
    "FormulaError",
    "Matrix",
    "Node",
    "as_bool",
    "as_number",
    "as_text",
    "is_volatile",
    "number_text",
    "parse",
    "references",
]
