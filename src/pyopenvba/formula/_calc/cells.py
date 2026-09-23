"""The engine's error value, and the error it raises for a formula it cannot work out.

pyOfficeEditor's engine carries an error as a :class:`CellError`, which a
cell of type ``e`` holds; pyOpenVBA's cells hold
:class:`~pyopenvba.formula._values.ExcelError`, and the calculator turns
one into the other where the two meet.
"""

from __future__ import annotations

from dataclasses import dataclass

from pyopenvba.exceptions import VBAUnsupportedError

#: Excel's own error strings.  A cell of type ``e`` holds one of these.
ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"}
)


@dataclass(frozen=True)
class CellError:
    """An Excel error value, such as ``#DIV/0!``.

    A distinct type, so a formula that failed is never mistaken for a cell
    whose text happens to look like an error.
    """

    code: str

    def __str__(self) -> str:
        return self.code

    @property
    def is_known(self) -> bool:
        """Whether this is one of the errors Excel documents."""
        return self.code in ERROR_CODES


class UnsupportedFormulaError(VBAUnsupportedError):
    """A formula the engine cannot work out as Excel would: a function it does not have, a reference to another
    workbook, a data table. Like every unsupported error, On Error cannot trap it."""


__all__ = ["ERROR_CODES", "CellError", "UnsupportedFormulaError"]
