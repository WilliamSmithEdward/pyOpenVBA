"""Where the engine meets the rest of pyOpenVBA: number formats, sheet names in quotes, and a formula moved.

pyOfficeEditor's engine takes these three from its own library. Here TEXT,
DOLLAR and FIXED format through pyOpenVBA's display engine, the one
Range.Text shows a cell with, so a formula and a cell cannot disagree.
"""

from __future__ import annotations

import re

from pyopenvba.formula._calc.cells import UnsupportedFormulaError
from pyopenvba.formula._calc.reference import CellRef

_BARE_SHEET_NAME = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.\\]*$")


def format_value(value: object, code: str, *, epoch_1904: bool = False) -> str:
    """The text a number format shows a value as, as TEXT gives it; #VALUE! where the format cannot show it."""
    # The tree's module reads quote_sheet_name from here, so the values, which read the tree, come in late.
    from pyopenvba.formula._calc.values import VALUE, ExcelError
    from pyopenvba.formula._display import UndisplayableError, format_value as displayed

    if epoch_1904:
        raise UnsupportedFormulaError("dates in the 1904 date system are not implemented")
    try:
        return displayed(value, code)
    except UndisplayableError:
        raise ExcelError(VALUE) from None


def quote_sheet_name(name: str) -> str:
    """A sheet name as a formula must spell it.

    Excel quotes a name that is not a bare identifier, and doubles any
    apostrophe inside it. A name that looks like a cell reference has to be
    quoted too, or ``=A1!B2`` would read as a reference to column A.
    """
    if _BARE_SHEET_NAME.match(name) and not _looks_like_a_reference(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def _looks_like_a_reference(name: str) -> bool:
    try:
        CellRef.parse(name)
    except ValueError:
        return False
    return True


def translate_formula(formula: str, rows: int, columns: int) -> str:
    """``formula`` as it reads ``rows`` down and ``columns`` across: relative references move, absolute ones stay."""
    from pyopenvba.formula._parse import shift_text

    return shift_text(formula, rows, columns)


__all__ = ["format_value", "quote_sheet_name", "translate_formula"]
