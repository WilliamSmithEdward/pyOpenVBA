"""Shared formulas: one formula written to a block, which Excel saves once for the whole block.

Measured in live Excel (scripts/measure_formula_storage.py,
tests/fixtures/formula_storage/). A formula written to several cells at
once -- through Formula, FormulaR1C1 or Value -- is one shared formula
over the block, unless it names no cell: =5 written to two cells is two
formulas. The file keeps the formula once, in the group's first cell with
the group's block, and every other cell only points at the group.

The block outlives its cells. Rewriting or clearing a cell of the group
takes that cell out and leaves the block as it was; when the first cell
goes, the next one carries the formula for the same block. Inserting and
deleting rows and columns stretch and shrink the block as they do a
reference. A save numbers the groups by where their first cell is, row by
row.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_area
from pyopenvba.formula._parse import FormulaError, shift_text, tokenize

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Cell, Worksheet


def shares(formula: str) -> bool:
    """Whether a formula written to a block becomes a shared formula: it names a cell."""
    try:
        return any(token.kind == "ref" for token in tokenize(formula.removeprefix("=")))
    except FormulaError:
        return False


def group(sheet: Worksheet, area: Area, cells: list[Cell]) -> None:
    """Make the cells a formula was just written to over ``area`` one shared formula."""
    key = max(sheet.shared_groups, default=-1) + 1
    sheet.shared_groups[key] = area
    for cell in cells:
        cell.shared = key


def formula_elements(sheet: Worksheet, escape: Callable[[str], str]) -> dict[tuple[int, int], str]:
    """The ``<f>`` element of each cell a save writes as part of a shared formula, by position.

    A group's first cell with a formula, in row order, carries the formula
    and the block; a later cell is written as the group's only while its
    formula is still the first one's moved to where it stands, and it lies
    inside the block. Anything else is written as a formula of its own.
    """
    members: dict[int, list[tuple[int, int, Cell]]] = {}
    for (row, column), cell in sorted(sheet.cells_.items()):
        key = cell.shared
        if key is None or not cell.formula:
            continue
        area = sheet.shared_groups.get(key)
        if area is not None and area.contains(row, column):
            members.setdefault(key, []).append((row, column, cell))
    out: dict[tuple[int, int], str] = {}
    for index, (key, cells) in enumerate(sorted(members.items(), key=lambda item: item[1][0][:2])):
        top, left, first = cells[0]
        block = sheet.shared_groups[key].address(absolute=False)
        out[(top, left)] = f'<f t="shared" ref="{block}" si="{index}">{escape(first.formula[1:])}</f>'
        for row, column, cell in cells[1:]:
            if shift_text(first.formula, row - top, column - left) == cell.formula:
                out[(row, column)] = f'<f t="shared" si="{index}"/>'
    return out


def moved(sheet: Worksheet, rewrite: Callable[[str], str]) -> None:
    """Carry each group's block through an insert or a delete, as ``rewrite`` carries a reference.

    A block the edit deletes whole is gone, and so is the group.
    """
    for key, area in list(sheet.shared_groups.items()):
        text = rewrite("=" + area.address(absolute=False))
        try:
            sheet.shared_groups[key] = parse_area(text.removeprefix("="), sheet=area.sheet)
        except ValueError:
            del sheet.shared_groups[key]
