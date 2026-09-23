"""Array formulas: Range.FormulaArray, and the cells one formula fills.

Measured in live Excel (tests/fixtures/excel_model/probes.txt, the array
formula section; tests/fixtures/formula_storage/arrays.xlsx). An array
formula is worked out once, as an array formula is, over its block, and
each cell of the block shows its item: one row or column repeats across
the block, and past the end of the answer a cell shows #N/A. The block's
first cell holds the formula; every cell of the block reports it as its
Formula and HasFormula, and a file keeps it in the first cell with the
block beside it.

Excel guards the block as one thing:

* FormulaArray written to one of its cells rewrites the whole block's
  formula; written over part of it, several cells, is error 1004; over
  all of it and more, a new array replaces it.
* Value, Value2, Formula and FormulaR1C1 written over part of it change
  nothing at all, not even the cells written outside it; written over
  all of it, plain cells replace it.
* ClearContents and Clear over part of it are error 1004.
* Inserting or deleting rows or columns through it is error 1004; the
  block moves with an insert before it and goes with a delete of all of
  it.
* Copying the whole block copies the array.

HasArray is True when a range meets one array and its first cell is in
it, False when it meets none, and Null otherwise. FormulaArray follows
it, and for a range with no array gives the formula its cells share, or
Null. CurrentArray is the array of the range's first cell.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, NULL, error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

#: What Excel says when an edit would take part of an array formula.
PART_OF_AN_ARRAY = "You can't change part of an array."
#: The longest formula FormulaArray takes.
LONGEST = 255


def array_at(sheet: Worksheet, row: int, column: int) -> tuple[tuple[int, int], Area] | None:
    """The first cell and the block of the array formula a cell is part of, if any."""
    for anchor, area in sheet.array_formulas.items():
        if area.contains(row, column):
            return anchor, area
    return None


def _meets(first: Area, second: Area) -> bool:
    return not (first.bottom < second.top or second.bottom < first.top
                or first.right < second.left or second.right < first.left)


def _covers(outer: Area, inner: Area) -> bool:
    return outer.top <= inner.top and inner.bottom <= outer.bottom and outer.left <= inner.left \
        and inner.right <= outer.right


def met(sheet: Worksheet, areas: list[Area]) -> tuple[list[tuple[int, int]], list[tuple[int, int]]]:
    """The arrays ``areas`` cover whole, and those they meet in part, each by its first cell."""
    whole: list[tuple[int, int]] = []
    part: list[tuple[int, int]] = []
    for anchor, block in sheet.array_formulas.items():
        if any(_meets(area, block) for area in areas):
            (whole if any(_covers(area, block) for area in areas) else part).append(anchor)
    return whole, part


def write_admitted(target: Range) -> bool:
    """Whether a Value or Formula write goes ahead: not when it meets part of an array, which it leaves as it was.

    One that covers arrays whole takes them away, for plain cells to
    replace them.
    """
    whole, part = met(target.sheet, target.areas)
    if part:
        return False
    for anchor in whole:
        del target.sheet.array_formulas[anchor]
    return True


def clear_admitted(target: Range) -> None:
    """Refuse a clear over part of an array; take away the arrays it clears whole."""
    whole, part = met(target.sheet, target.areas)
    if part:
        raise error(1004, PART_OF_AN_ARRAY)
    for anchor in whole:
        del target.sheet.array_formulas[anchor]


def refuse(sheet: Worksheet, areas: list[Area], doing: str) -> None:
    """Report an operation the model does not carry through array formulas."""
    whole, part = met(sheet, areas)
    if whole or part:
        raise VBAUnsupportedError(f"{doing} cells of an array formula is not implemented")


def put(target: Range, formula: str) -> None:
    """Range.FormulaArray written: the formula, already spelled, over the range as one array."""
    sheet = target.sheet
    if len(target.areas) != 1:
        raise VBAUnsupportedError("FormulaArray over several areas is not implemented")
    area = target.first
    if len(formula) > LONGEST:
        raise error(1004, "Unable to set the FormulaArray property of the Range class")
    found = array_at(sheet, area.top, area.left) if area.rows == area.columns == 1 else None
    if found is not None:
        # One cell of an array rewrites the whole array's formula.
        anchor, block = found
    else:
        whole, part = met(sheet, [area])
        if part:
            raise error(1004, PART_OF_AN_ARRAY)
        for one in whole:
            del sheet.array_formulas[one]
        block = Area(area.top, area.left, area.bottom, area.right)
        anchor = (block.top, block.left)
        for position in [position for position in sheet.cells_ if block.contains(*position)]:
            cell = sheet.cells_[position]
            cell.value, cell.formula, cell.stale, cell.shared = EMPTY, "", False, None
            sheet.settle(*position)
        sheet.array_formulas[anchor] = block
    cell = sheet.cell(*anchor, create=True)
    assert cell is not None
    cell.formula, cell.stale, cell.value, cell.shared = formula, True, EMPTY, None
    for row in range(block.top, block.bottom + 1):
        for column in range(block.left, block.right + 1):
            sheet.cell_changed(row, column)


def has_array(target: Range) -> object:
    """HasArray: True, False or Null, as the module docstring has it."""
    whole, part = met(target.sheet, target.areas)
    arrays = whole + part
    if not arrays:
        return False
    first = target.first
    if len(arrays) == 1 and array_at(target.sheet, first.top, first.left) is not None:
        return True
    return NULL


def formula_array(target: Range, formula_of: Callable[[int, int], str]) -> object:
    """FormulaArray read: the array's formula, a formula every cell shares, or Null."""
    first = target.first
    held = has_array(target)
    if held is True:
        found = array_at(target.sheet, first.top, first.left)
        assert found is not None
        return formula_of(*found[0])
    if held is NULL:
        return NULL
    texts: set[str] = set()
    for area in target.areas:
        if area.rows * area.columns > 1048576:
            raise VBAUnsupportedError("FormulaArray read over more than 1,048,576 cells is not implemented")
        for row in range(area.top, area.bottom + 1):
            for column in range(area.left, area.right + 1):
                texts.add(formula_of(row, column))
                if len(texts) > 1:
                    return NULL
    return texts.pop()


def current_array(target: Range) -> Area:
    """The block of the array the range's first cell is part of; error 1004 when it is part of none."""
    first = target.first
    found = array_at(target.sheet, first.top, first.left)
    if found is None:
        raise error(1004, "No array formula at this cell")
    return found[1]


def edit_admitted(sheet: Worksheet, *, rows: bool, start: int, count: int, delete: bool) -> None:
    """Refuse inserting or deleting rows (or columns) through an array formula, before anything moves."""
    for block in sheet.array_formulas.values():
        low, high = (block.top, block.bottom) if rows else (block.left, block.right)
        if delete:
            end = start + count - 1
            if start <= high and low <= end and not (start <= low and high <= end):
                raise error(1004, PART_OF_AN_ARRAY)
        elif low < start <= high:
            raise error(1004, PART_OF_AN_ARRAY)


def moved(sheet: Worksheet, rewrite: Callable[[str], str]) -> None:
    """Carry each array's block through an insert or a delete, as ``rewrite`` carries a reference.

    A block the edit deletes whole is gone, and so is the array.
    """
    carried: dict[tuple[int, int], Area] = {}
    for area in sheet.array_formulas.values():
        text = rewrite("=" + area.address(absolute=False))
        try:
            block = parse_area(text.removeprefix("="), sheet="")
        except ValueError:
            continue
        carried[(block.top, block.left)] = block
    sheet.array_formulas = carried


def elements(sheet: Worksheet, escape: Callable[[str], str],
             always: Callable[[str], bool]) -> dict[tuple[int, int], str]:
    """The ``<f>`` element of each array's first cell, as a save writes it.

    An array whose formula ``always`` works out whenever anything changes is ca="1", and each of its other cells
    carries <f ca="1"/> beside its value; it is aca="1" too, unless it covers several cells and calls RAND
    (tests/fixtures/formula_prefixes/)."""
    from pyopenvba.formula._prefixes import calls

    out: dict[tuple[int, int], str] = {}
    for (row, column), block in sheet.array_formulas.items():
        cell = sheet.cells_.get((row, column))
        if cell is None or not cell.formula:
            continue
        volatile = always(cell.formula)
        whole = volatile and (block.rows * block.columns == 1 or not calls(cell.formula, "RAND"))
        flags = (' aca="1"' if whole else "", ' ca="1"' if volatile else "")
        out[(row, column)] = (f'<f t="array"{flags[0]} ref="{block.address(absolute=False)}"{flags[1]}>'
                              f'{escape(cell.formula[1:])}</f>')
        if volatile:
            for member_row in range(block.top, block.bottom + 1):
                for member_column in range(block.left, block.right + 1):
                    out.setdefault((member_row, member_column), '<f ca="1"/>')
    return out
