"""DisplayFormat views over conditional styles, without changing cell formats."""
from __future__ import annotations

from collections.abc import Callable
from typing import TYPE_CHECKING, Any, cast

from pyopenvba._a1 import MAX_COLUMNS, MAX_ROWS, Area
from pyopenvba.apps.excel import _formats as formats
from pyopenvba.apps.excel import _styles as styles
from pyopenvba.apps.excel._model import ExcelObject, Range
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import member
from pyopenvba.interpreter._values import MISSING, NULL, error, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Worksheet


def _intersection(one: Area, other: Area) -> Area | None:
    top, left = max(one.top, other.top), max(one.left, other.left)
    bottom, right = min(one.bottom, other.bottom), min(one.right, other.right)
    return Area(top, left, bottom, right) if top <= bottom and left <= right else None


def _without(area: Area, cut: Area) -> list[Area]:
    overlap = _intersection(area, cut)
    if overlap is None:
        return [area]
    found: list[Area] = []
    if area.top < overlap.top:
        found.append(Area(area.top, area.left, overlap.top - 1, area.right))
    if overlap.bottom < area.bottom:
        found.append(Area(overlap.bottom + 1, area.left, area.bottom, area.right))
    if area.left < overlap.left:
        found.append(Area(overlap.top, area.left, overlap.bottom, overlap.left - 1))
    if overlap.right < area.right:
        found.append(Area(overlap.top, overlap.right + 1, overlap.bottom, area.right))
    return found


class _DisplaySheet:
    """A local read adapter; never replaces or patches the real worksheet."""

    def __init__(self, sheet: Worksheet) -> None:
        self.original = sheet

    def __getattr__(self, name: str) -> Any:
        return getattr(self.original, name)

    def style_at(self, row: int, column: int) -> styles.Style:
        from pyopenvba.apps.excel._conditional_eval import display_style

        return display_style(self.original, row, column)


def _computed_range(target: Range, *, split: bool = True) -> Range:
    if not split:
        # A neighbour's conditional opposite side can draw this cell's edge.
        return Range(cast('Worksheet', _DisplaySheet(target.sheet)), list(target.areas))
    covered: list[Area] = []
    remaining = list(target.areas)
    for rule in target.sheet.conditional_formats:
        for area in rule.areas:
            for one in remaining:
                overlap = _intersection(one, area)
                if overlap is not None:
                    covered.append(overlap)
            remaining = [piece for one in remaining for piece in _without(one, area)]
    if not covered:
        return target
    # Preserve sparse whole-row/column reads outside the rules; walk every
    # covered cell because formulas can distinguish otherwise identical styles.
    if sum(area.rows * area.columns for area in covered) > 1048576:
        raise VBAUnsupportedError('Reading conditional display styles over more than 1,048,576 covered cells is not implemented')
    parts = list(remaining)
    for area in covered:
        for left in range(area.left, area.right + 1, 4096):
            right = min(area.right, left + 4095)
            height = max(1, 4096 // (right - left + 1))
            parts.extend(Area(top, left, min(area.bottom, top + height - 1), right)
                         for top in range(area.top, area.bottom + 1, height))
    return Range(cast('Worksheet', _DisplaySheet(target.sheet)), parts)


class _DisplayFont(formats.Font):
    vba_type_name = 'Font'

    def __init__(self, target: Range) -> None:
        self.original = target

    @property
    def target(self) -> Range:
        return _computed_range(self.original)

    def guard_set(self, member: str) -> None:
        formats.Font(self.original).guard_set(member)

    def _write(self, change: Callable[[styles.Font], styles.Font]) -> None:
        # Native DisplayFormat.Font setters change the underlying Font while
        # matching conditional rules continue to control what is displayed.
        formats.restyle(self.original, lambda style: styles.applying(style, 'font', font=change(style.font)))

    @member
    def Parent(self) -> object:
        return self.original


class _DisplayInterior(formats.Interior):
    vba_type_name = 'Interior'

    def __init__(self, target: Range) -> None:
        self.original = target

    @property
    def target(self) -> Range:
        return _computed_range(self.original)

    def guard_set(self, member: str) -> None:
        raise error(1004)

    def _write(self, change: Callable[[styles.Fill], styles.Fill]) -> None:
        raise error(1004)

    @member
    def Parent(self) -> object:
        return self.original


class _DisplayBorder(formats.Border):
    vba_type_name = 'Border'

    def __init__(self, target: Range, index: int) -> None:
        self.original, self.index = target, index

    @property
    def target(self) -> Range:
        return _computed_range(self.original, split=False)

    def guard_set(self, member: str) -> None:
        raise error(1004)

    def apply(self, change: Callable[[styles.Side], styles.Side]) -> None:
        raise error(1004)

    def sides(self) -> list[styles.Side]:
        area = self.original.first
        if self.index in (5, 6):
            # Diagonals belong to each cell, not its right-hand edge. Distinct
            # styles suffice for the collection's uniform/mixed aggregation.
            first = Range(self.original.sheet, [area])
            return [style.border.diagonal if (style.border.diagonal_down if self.index == 5
                                              else style.border.diagonal_up) else styles.NO_SIDE
                    for style in formats.styles_of(_computed_range(first))]
        around = Area(max(1, area.top - 1), max(1, area.left - 1),
                      min(MAX_ROWS, area.bottom + 1), min(MAX_COLUMNS, area.right + 1))
        if not any(not rule.deleted and _intersection(around, covered) is not None
                   for rule in self.original.sheet.conditional_formats for covered in rule.areas):
            # Base readers sample sparse whole-row/column formats, rather than
            # enumerating a million edges that no conditional rule affects.
            return formats.Border(self.original, self.index).sides()
        sheet = self.target.sheet
        if self.index in formats.EDGES:
            name = formats.EDGES[self.index]
            if name in ('top', 'bottom'):
                row = area.top if name == 'top' else area.bottom
                return [formats.effective_side(sheet, row, column, name) for column in range(area.left, area.right + 1)]
            column = area.left if name == 'left' else area.right
            return [formats.effective_side(sheet, row, column, name) for row in range(area.top, area.bottom + 1)]
        if area.rows * area.columns > 1048576:
            raise VBAUnsupportedError('Reading every conditional inside border over this range is not implemented')
        if self.index == 12:
            return [formats.effective_side(sheet, row, column, 'bottom') for row in range(area.top, area.bottom)
                    for column in range(area.left, area.right + 1)]
        return [formats.effective_side(sheet, row, column, 'right') for row in range(area.top, area.bottom + 1)
                for column in range(area.left, area.right)]

    @member
    def Parent(self) -> object:
        return self.original


class _DisplayBorders(formats.Borders):
    vba_type_name = 'Borders'

    def guard_set(self, member: str) -> None:
        raise error(1004)

    def vba_items(self) -> list[object]:
        return [_DisplayBorder(self.target, index) for index in (7, 8, 9, 10, 5, 6)]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        number = int(to_integer(index, 'Long'))
        aliases = {1: 7, 2: 10, 3: 8, 4: 9, **{i: i for i in range(5, 13)}}
        if number not in aliases:
            raise error(1004)
        return _DisplayBorder(self.target, aliases[number])

    def _all(self, what: str) -> object:
        colors = self.target.sheet.book.stylesheet.colors
        found = formats.uniform(formats.side_answer(colors, side, what) for index in self._indexes()
                          for side in _DisplayBorder(self.target, index).sides())
        return 0.0 if found is NULL and what == 'Color' else found

    def _set_all(self, what: str, value: object) -> None:
        raise error(1004)


class DisplayFormat(ExcelObject):
    """A live view of the ordinary conditional formats affecting a range."""

    vba_type_name = 'DisplayFormat'

    def __init__(self, target: Range) -> None:
        self.target, self.sheet = target, target.sheet

    def guard_set(self, member: str) -> None:
        raise error(424)

    @member
    def Parent(self) -> object:
        return self.target

    @member
    def Application(self) -> object:
        return self.sheet.book.application

    @member
    def Font(self) -> object:
        return _DisplayFont(self.target)

    @member
    def Interior(self) -> object:
        return _DisplayInterior(self.target)

    @member
    def Borders(self, Index: object = MISSING) -> object:  # noqa: N803 -- public Excel argument
        borders = _DisplayBorders(self.target)
        return borders if Index is MISSING else borders.Item(Index)

    @member
    def NumberFormat(self) -> object:
        # Native Excel reports the base format here even when its renderer uses
        # a conditional number format for Range.Text.
        return self.target.NumberFormat()

    @member
    def NumberFormatLocal(self) -> object:
        return self.target.NumberFormatLocal()

    def _alignment(self, name: str) -> object:
        return formats.read_alignment(_computed_range(self.target), name)

    @member
    def HorizontalAlignment(self) -> object:
        return self._alignment('HorizontalAlignment')

    @member
    def VerticalAlignment(self) -> object:
        return self._alignment('VerticalAlignment')

    @member
    def WrapText(self) -> object:
        return self._alignment('WrapText')

    @member
    def ShrinkToFit(self) -> object:
        return self._alignment('ShrinkToFit')

    @member
    def AddIndent(self) -> object:
        return self._alignment('AddIndent')

    @member
    def IndentLevel(self) -> object:
        return self._alignment('IndentLevel')

    @member
    def ReadingOrder(self) -> object:
        return self._alignment('ReadingOrder')

    @member
    def Orientation(self) -> object:
        return self._alignment('Orientation')

    @member
    def Locked(self) -> object:
        return self._alignment('Locked')

    @member
    def FormulaHidden(self) -> object:
        return self._alignment('FormulaHidden')
