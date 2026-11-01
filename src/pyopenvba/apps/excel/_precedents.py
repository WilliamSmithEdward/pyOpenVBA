"""Range.Precedents, DirectPrecedents, Dependents and DirectDependents: the cells on a sheet a range's formulas read,
and the formulas there that read the range, as Excel finds and orders them.

Measured in live Excel (scripts/measure_precedents.py, tests/fixtures/precedents.json):

- A formula's precedents are its references on its own sheet in the order
  the formula gives them, a name's cells among them. Another sheet's are
  left out, and so is what INDIRECT names, while OFFSET's reference counts.
  A reference a legacy formula cuts to one value, or an @ cuts, counts as
  the cell the formula's row or column meets, where it meets one, and whole
  where it does not.
- DirectPrecedents of a range are its formula cells' precedents, taken
  along the range. Precedents adds theirs, each formula taken once, level
  by level.
- DirectDependents are the formulas on the sheet reading the range, a
  dynamic-array formula by the whole of what it spills, in the order of
  their rows and then their columns. Dependents adds those reading them,
  level by level, leaving out a formula found before.
- An answer joins its areas as they come: one that makes a rectangle with
  an area already there, holds it or is held by it, takes that area out,
  and the two go last as one. That is why Excel's areas come in the order
  they do: a chain of four answers "$B$1,$A$1:$A$4".
- Nothing found is error 1004, "No cells were found." A file answers as the
  workbook it was saved from did.
"""

from __future__ import annotations

from collections.abc import Iterable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._engine_book import block
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._calc.lexer import FormulaSyntaxError
from pyopenvba.formula._calc.nodes import (
    AreaReference,
    AxisReference,
    Binary,
    Call,
    CellReference,
    Invoke,
    NameReference,
    Node,
    Paren,
    Postfix,
    Prefix,
    Unary,
    function_key,
)
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

NOTHING = "No cells were found."


def _joined(one: Area, other: Area) -> Area | None:
    """The one rectangle two areas make together, where they make one."""
    if one.top <= other.top and other.bottom <= one.bottom and one.left <= other.left and other.right <= one.right:
        return one
    if other.top <= one.top and one.bottom <= other.bottom and other.left <= one.left and one.right <= other.right:
        return other
    if one.left == other.left and one.right == other.right and other.top <= one.bottom + 1 \
            and one.top <= other.bottom + 1:
        return Area(min(one.top, other.top), one.left, max(one.bottom, other.bottom), one.right, one.sheet)
    if one.top == other.top and one.bottom == other.bottom and other.left <= one.right + 1 \
            and one.left <= other.right + 1:
        return Area(one.top, min(one.left, other.left), one.bottom, max(one.right, other.right), one.sheet)
    return None


def joined(areas: list[Area], new: Area) -> None:
    """Add ``new`` to ``areas`` as Excel's answers do: an area it joins into one rectangle goes, and the joined area
    goes last (see the module's notes)."""
    while True:
        for index, one in enumerate(areas):
            together = _joined(one, new)
            if together is not None:
                del areas[index]
                new = together
                break
        else:
            areas.append(new)
            return


def _cut(area: Area, row: int, column: int) -> Area:
    """What one value a reference is cut to reads, from the formula at ``row`` and ``column``: the cell its row meets
    in a column of cells, its column in a row of them; the whole where it meets none."""
    if area.left == area.right and area.top <= row <= area.bottom:
        return Area(row, area.left, row, area.left, area.sheet)
    if area.top == area.bottom and area.left <= column <= area.right:
        return Area(area.top, column, area.top, column, area.sheet)
    return area


def _here(prefix: Prefix | None, sheet: Worksheet) -> bool:
    """Whether a reference's prefix names ``sheet``, or there is none."""
    from pyopenvba.apps.excel._names import same_sheet_name

    if prefix is None:
        return True
    return prefix.book is None and prefix.sheet is not None and prefix.last_sheet is None \
        and same_sheet_name(prefix.sheet, sheet.name)


def reads(sheet: Worksheet, row: int, column: int) -> list[Area]:
    """The cells on its own sheet a formula reads, in the order the formula gives them."""
    from pyopenvba.apps.excel._model import Range

    book = sheet.book.calculator.engine_book
    text = Range(sheet, [Area(row, column, row, column, sheet.name)]).vba_get("Formula2")
    try:
        node = book.read(str(text))
    except FormulaSyntaxError:
        raise VBAUnsupportedError("the precedents of a formula the model cannot read are not implemented") from None
    found: list[Area] = []
    names: set[str] = set()
    stack: list[tuple[Node, bool]] = [(node, False)]
    while stack:
        current, cutting = stack.pop()
        if isinstance(current, Unary) and current.op == "@":
            stack.append((current.operand, True))
        elif isinstance(current, Call) and function_key(current.function) == "SINGLE":
            # The engine reads a formula as a file spells it, an @ as a call of SINGLE.
            stack.extend((one, True) for one in reversed(current.args))
        elif isinstance(current, (CellReference, AreaReference, AxisReference)):
            if _here(current.prefix, sheet):
                area = block(current, sheet.name)
                found.append(_cut(area, row, column) if cutting else area)
        elif isinstance(current, NameReference):
            key = current.name.casefold()
            formula = None if current.prefix is not None or key in names else \
                book.defined_name(current.name, sheet.name) or book.defined_name(current.name, None)
            if formula is not None:
                names.add(key)
                stack.append((formula, cutting))
        elif isinstance(current, (Unary, Postfix)):
            stack.append((current.operand, cutting))
        elif isinstance(current, Binary):
            stack.extend(((current.right, cutting), (current.left, cutting)))
        elif isinstance(current, Call):
            stack.extend((one, cutting) for one in reversed(current.args))
        elif isinstance(current, Invoke):
            stack.extend((one, cutting) for one in reversed(current.args))
            stack.append((current.target, cutting))
        elif isinstance(current, Paren):
            stack.append((current.inner, cutting))
    return found


def _formula_cells(sheet: Worksheet, areas: Iterable[Area]) -> list[tuple[int, int]]:
    """The formula cells of ``areas``, area by area, each row by row, each cell once."""
    seen: set[tuple[int, int]] = set()
    found: list[tuple[int, int]] = []
    formulas = sorted(position for position, cell in sheet.cells_.items() if cell.formula)
    for area in areas:
        for position in formulas:
            if position not in seen and area.contains(*position):
                seen.add(position)
                found.append(position)
    return found


def _precedents_of(sheet: Worksheet, areas: Iterable[Area], taken: set[tuple[int, int]]) -> list[Area]:
    """What the formulas of ``areas`` not yet taken read, joined as Excel joins it; each one taken now."""
    found: list[Area] = []
    for position in _formula_cells(sheet, areas):
        if position in taken:
            continue
        taken.add(position)
        for area in reads(sheet, *position):
            joined(found, area)
    return found


def _readers(sheet: Worksheet) -> list[tuple[tuple[int, int], Area, list[Area]]]:
    """Each formula on the sheet, where it shows its answer, a dynamic-array formula's whole spill, and what it
    reads."""
    found: list[tuple[tuple[int, int], Area, list[Area]]] = []
    for position, cell in sorted(sheet.cells_.items()):
        if not cell.formula:
            continue
        spill = sheet.spills.get(position)
        shown = spill.area if spill is not None else Area(*position, *position, sheet.name)
        found.append((position, Area(shown.top, shown.left, shown.bottom, shown.right, sheet.name),
                      reads(sheet, *position)))
    return found


def _dependents_of(readers: list[tuple[tuple[int, int], Area, list[Area]]], areas: list[Area],
                   found_before: set[tuple[int, int]]) -> list[Area]:
    """The formulas reading ``areas`` not found before, in the order of their rows and columns, joined as Excel
    joins them; each one found now."""
    reading = [(shown, position) for position, shown, reads in readers if position not in found_before
               and any(_meets(read, area) for read in reads for area in areas)]
    found: list[Area] = []
    for shown, position in sorted(reading, key=lambda one: (one[0].top, one[0].left)):
        found_before.add(position)
        joined(found, shown)
    return found


def _meets(one: Area, other: Area) -> bool:
    return one.top <= other.bottom and other.top <= one.bottom and one.left <= other.right \
        and other.left <= one.right


def _folded(levels: list[list[Area]]) -> list[Area]:
    """Each level joined with all those after it, the deepest first, as Excel's answer is made."""
    after: list[Area] = []
    for level in reversed(levels):
        together = list(level)
        for area in after:
            joined(together, area)
        after = together
    return after


def answer(target: Range, *, dependents: bool, direct: bool) -> Range:
    """Range.Precedents, DirectPrecedents, Dependents or DirectDependents of ``target``."""
    from pyopenvba.apps.excel._model import Range

    sheet = target.sheet
    areas = [Area(one.top, one.left, one.bottom, one.right, sheet.name) for one in target.areas]
    levels: list[list[Area]] = []
    if dependents:
        readers = _readers(sheet)
        found_before: set[tuple[int, int]] = set()
        level = _dependents_of(readers, areas, found_before)
        while level:
            levels.append(level)
            if direct:
                break
            level = _dependents_of(readers, level, found_before)
    else:
        taken: set[tuple[int, int]] = set()
        level = _precedents_of(sheet, areas, taken)
        while level:
            levels.append(level)
            if direct:
                break
            level = _precedents_of(sheet, level, taken)
    found = _folded(levels)
    if not found:
        raise error(1004, NOTHING)
    return Range(sheet, found)
