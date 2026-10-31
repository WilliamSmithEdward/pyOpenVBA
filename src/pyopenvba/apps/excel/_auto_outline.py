"""Range.AutoOutline: rows and columns grouped from the formulas that sum the cells beside them
(tests/fixtures/auto_outline.json, scripts/measure_auto_outline.py).

A formula reading two cells or more, all in its own column and in one block above or below it, groups the rows
from that block to the row before or after it, a blank row between them too; one reading its own row groups
columns the same way. The direction it finds sets Outline.SummaryRow or SummaryColumn. A single cell outlines the
whole sheet, a larger range the formulas in it; the outline it makes takes the place of any there was, and with
nothing to group it is error 1004 "Cannot create an outline.".

Excel nests groups whose summaries read other summaries; the model makes one level, and reports a formula reading
cells of its column apart from one another, or groups that overlap, as unsupported.
"""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import parse_reference, split_sheet
from pyopenvba.apps.excel import _outline
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import error

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

_FAILED = "Cannot create an outline."


def _cells_read(sheet: Worksheet, formula: str) -> list[tuple[int, int]] | None:
    """The cells of this sheet a formula reads by reference, or None where it reads another sheet's."""
    from pyopenvba.formula._parse import FormulaError, tokenize

    try:
        tokens = tokenize(formula.lstrip("="))
    except FormulaError:
        return None
    cells: list[tuple[int, int]] = []
    for token in tokens:
        if token.kind != "ref":
            continue
        name, bare = split_sheet(token.text)
        if name and name.lower() != sheet.name.lower():
            return None
        try:
            areas = parse_reference(bare.replace("$", ""))
        except ValueError:
            return None
        for area in areas:
            if area.rows * area.columns > 4096:
                return None
            cells.extend((row, column) for row in range(area.top, area.bottom + 1)
                         for column in range(area.left, area.right + 1))
    return cells


def _group(own: int, lines: list[int]) -> tuple[int, int, bool] | None:
    """The lines a summary at ``own`` groups from the lines it reads, and whether it sits after them."""
    if len(lines) < 2:
        return None
    first, last = min(lines), max(lines)
    if sorted(set(lines)) != list(range(first, last + 1)):
        raise VBAUnsupportedError("AutoOutline of a formula reading cells apart from one another is not implemented")
    if last < own:
        return first, own - 1, True
    if first > own:
        return own + 1, last, False
    return None


def auto_outline(target: Range) -> None:
    sheet = target.sheet
    whole = len(target.areas) == 1 and target.first.rows == 1 and target.first.columns == 1
    row_groups: list[tuple[int, int, bool]] = []
    column_groups: list[tuple[int, int, bool]] = []
    for (row, column), cell in sorted(sheet.cells_.items()):
        if not cell.formula or not (whole or any(area.contains(row, column) for area in target.areas)):
            continue
        cells = _cells_read(sheet, cell.formula)
        if not cells:
            continue
        if all(c == column for _, c in cells):
            found = _group(row, [r for r, _ in cells])
            if found is not None:
                row_groups.append(found)
        elif all(r == row for r, _ in cells):
            found = _group(column, [c for _, c in cells])
            if found is not None:
                column_groups.append(found)
    if not row_groups and not column_groups:
        raise error(1004, _FAILED)
    rows, below = _laid(row_groups)
    columns, right = _laid(column_groups)
    depths = {number: 0 for number in _outline.outlined(sheet, True)} | dict.fromkeys(rows, 1)
    widths = {number: 0 for number in _outline.outlined(sheet, False)} | dict.fromkeys(columns, 1)
    _outline.relevel(sheet, depths, widths, set(), set())
    if below is not None:
        _outline.set_summary_below(sheet, below)
    if right is not None:
        _outline.set_summary_right(sheet, right)


def _laid(groups: list[tuple[int, int, bool]]) -> tuple[set[int], bool | None]:
    """The lines the groups cover and the side their summaries sit on; groups that overlap report themselves."""
    distinct = sorted(set(groups))
    for (_, last, _), (other, _, _) in zip(distinct, distinct[1:]):
        if other <= last:
            raise VBAUnsupportedError("AutoOutline of groups that overlap, which Excel nests, is not implemented")
    sides = {after for _, _, after in distinct}
    if len(sides) > 1:
        raise VBAUnsupportedError("AutoOutline with summaries on both sides of their detail is not implemented")
    lines = {number for first, last, _ in distinct for number in range(first, last + 1)}
    return lines, (next(iter(sides)) if sides else None)
