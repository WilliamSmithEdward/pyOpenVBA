"""Which of Excel's worksheet functions the formula engine does not implement, for the docs to list.

Excel's functions are the engine's catalog,
:data:`pyopenvba.formula._calc.catalog.EXCEL_FUNCTIONS`: every function a
sheet has, spelled as a formula spells it. A cell calling one of the
rest says it is not implemented; a name the catalog has not got is
``#NAME?``, as in Excel.
"""

from __future__ import annotations


def unimplemented() -> list[str]:
    """Every Excel function this does not implement, for the docs to list."""
    from pyopenvba.formula._calc import functions as functions  # imported to register every function
    from pyopenvba.formula._calc.catalog import EXCEL_FUNCTIONS
    from pyopenvba.formula._calc.registry import FUNCTIONS

    return sorted(EXCEL_FUNCTIONS - frozenset(FUNCTIONS))
