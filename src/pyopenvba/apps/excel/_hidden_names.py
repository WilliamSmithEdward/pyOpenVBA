"""The hidden names a workbook's formulas bring into Names, which no file lists: _xlfn.SINGLE where a formula shows an @ in
Formula2, _xlfn.NAME for a function a file prefixes or one of a few older ones, _xlpm.name for what a LET or a LAMBDA binds.
Each reads =#NAME?. Once there it stays while the workbook is open, whatever becomes of the formula, and opening a file brings
those its formulas call for (tests/fixtures/xlfn_names.json, tests/fixtures/xlfn_functions.json)."""
from __future__ import annotations

import re
from functools import lru_cache
from typing import TYPE_CHECKING, Final

from pyopenvba.formula._prefixes import in_file

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Workbook

#: Older functions a file writes bare, which bring a name of their own as those it prefixes do.
_OLDER: Final = frozenset({
    "AVERAGEIF", "AVERAGEIFS", "BAHTTEXT", "COUNTIFS", "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER",
    "CUBESET", "CUBESETCOUNT", "CUBEVALUE", "ECMA.CEILING", "IFERROR", "ISO.CEILING", "NETWORKDAYS.INTL", "RTD", "SUMIFS",
    "WORKDAY.INTL"})
_TEXT: Final = re.compile(r'"[^"]*"')
_PREFIXED: Final = re.compile(r"(?<![\w.])_xl(?:fn|pm)\.[\w.]+")
_CALLED: Final = re.compile(r"(?<![\w.])([^\W\d][\w.]*)\(")
SINGLE: Final = "_xlfn.SINGLE"


@lru_cache(maxsize=4096)
def names_in(formula: str, *, single: bool = True) -> frozenset[str]:
    """The hidden names a formula brings, from its Formula2 text; ``single`` False for a defined name's, which brings none for an @."""
    try:
        spelled = in_file(formula)
    except ValueError:
        return frozenset()
    bare = _TEXT.sub('""', spelled)
    found = set(_PREFIXED.findall(bare))
    found.update(f"_xlfn.{name.upper()}" for name in _CALLED.findall(bare) if name.upper() in _OLDER)
    if not single:
        found.discard(SINGLE)
    return frozenset(found)


def noted(book: Workbook, formula: str, *, single: bool = True) -> None:
    book.hidden_names.update(names_in(formula, single=single))


def edited(target: Range) -> None:
    """An edit: what the formula in each area's first cell brings, which any other cell of the area repeats."""
    from pyopenvba.apps.excel._model import Range

    sheet = target.sheet
    for area in target.areas:
        cell = sheet.cell(area.top, area.left)
        if cell is not None and cell.formula:
            first = Range(sheet, [type(area)(area.top, area.left, area.top, area.left, sheet.name)])
            noted(sheet.book, str(first.Formula2()))


def scanned(book: Workbook) -> None:
    """What the file's formulas call for, once, as Names is first asked: reading them costs more than opening needs."""
    from pyopenvba._a1 import Area
    from pyopenvba.apps.excel._model import Range

    if not book.hidden_names_due:
        return
    book.hidden_names_due = False
    for sheet in book.sheets_:
        for (row, column), cell in list(sheet.cells_.items()):
            if cell.formula:
                noted(book, str(Range(sheet, [Area(row, column, row, column, sheet.name)]).Formula2()))
    for entry in book.names_.entries:
        noted(book, entry.refers_to, single=False)
