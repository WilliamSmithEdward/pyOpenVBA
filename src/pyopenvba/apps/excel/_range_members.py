"""Range.NavigateArrow and Range.Justify, measured in live Excel (scripts/measure_range_members.py,
tests/fixtures/range_members.json).

The model draws no tracer arrows, but it keeps which cells ShowPrecedents and ShowDependents have traced, how many levels
deep, since NavigateArrow follows them. A traced cell's first arrow leads to another sheet, with a link for each reference
there in the order the formula gives them. The rest lead to the references on its own sheet, last in the formula first,
each once. An arrow number past the last leaves the cell where it is, and a link number past the last raises 1004.

Justify runs the text of a range's first column together and breaks it again to the width of the range's columns, a word
to a line at least. It writes from the top down, past the range if it must, and empties what the range has left over.
Where it breaks depends on the screen: on a 144 DPI display, the only one measured, lines reached up to about a fifth
wider than the column at 96 DPI measures. So a line within that margin is refused, and only what fits plainly or plainly
does not is done."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, parse_reference
from pyopenvba.apps.excel._creator import Created
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.formula._parse import split_sheet, tokenize
from pyopenvba.interpreter._objects import member, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, error, to_bool, to_integer

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet


class SoundNote(Created):
    """A cell's sound note: an object every cell answers, whose members are not implemented."""

    vba_type_name = "SoundNote"

    def __init__(self, target: Range) -> None:
        self.target = target


class CellControl(Created):
    """A cell's control: an object every cell answers, whose members are not implemented."""

    vba_type_name = "CellControl"

    def __init__(self, target: Range) -> None:
        self.target = target


class XPath(Created):
    """A cell's XML mapping, of which the model has none: an empty path, and asking for its map fails with Excel's
    unspecified error (tests/fixtures/range_missing.json)."""

    vba_type_name = "XPath"

    def __init__(self, target: Range) -> None:
        self.target = target

    @member
    def Value(self) -> object:
        return ""

    @member
    def Map(self) -> object:
        raise error(-2147467259, "Automation error")


class EmptyPhonetic(Created):
    """The furigana of a cell with none."""

    vba_type_name = 'Phonetic'

    @member
    def Text(self) -> object:
        return ''

    @member
    def Visible(self) -> object:
        return False

    @setter('Visible')
    def _visible(self, value: object) -> None:
        """Excel takes it, and reads False still."""


def trace(target: Range, *, dependents: bool, remove: object) -> None:
    """Another level of arrows at the range's first cell, or one fewer."""
    first = target.first
    key = (first.top, first.left, dependents)
    depth = target.sheet.traced.get(key, 0)
    depth += -1 if remove is not MISSING and to_bool(remove) else 1
    if depth > 0:
        target.sheet.traced[key] = depth
    else:
        target.sheet.traced.pop(key, None)


def _traced(sheet: Worksheet, row: int, column: int) -> bool:
    """Whether the cell's precedents have arrows: it was traced, or a level of tracing reached it."""
    from pyopenvba.apps.excel._precedents import reads

    for (top, left, dependents), depth in sheet.traced.items():
        if dependents:
            continue
        level = {(top, left)}
        for _ in range(depth):
            if (row, column) in level:
                return True
            level = {(cell_row, cell_column) for one_row, one_column in level
                     for area in reads(sheet, one_row, one_column)
                     for cell_row in range(area.top, area.bottom + 1)
                     for cell_column in range(area.left, area.right + 1)}
    return False


def _elsewhere(sheet: Worksheet, formula: str) -> list[Area]:
    """The references of a formula to other sheets, in its order."""
    found: list[Area] = []
    for token in tokenize(formula):
        if token.kind != 'ref':
            continue
        name, _ = split_sheet(token.text)
        if name and name.casefold() != sheet.name.casefold():
            if name.startswith('['):
                raise VBAUnsupportedError('NavigateArrow to another workbook is not implemented')
            found += parse_reference(token.text)
    return found


def _precedent_arrows(sheet: Worksheet, row: int, column: int) -> list[list[Area]]:
    from pyopenvba.apps.excel._precedents import reads

    cell = sheet.cells_.get((row, column))
    arrows: list[list[Area]] = []
    if cell is None or not cell.formula:
        return arrows
    away = _elsewhere(sheet, cell.formula)
    if away:
        arrows.append(away)
    local: list[Area] = []
    for area in reads(sheet, row, column):
        if not any((one.top, one.left, one.bottom, one.right) == (area.top, area.left, area.bottom, area.right)
                   for one in local):
            local.append(area)
    arrows += [[area] for area in reversed(local)]
    return arrows


def _dependent_arrows(sheet: Worksheet, row: int, column: int) -> list[list[Area]]:
    """Those to other sheets, then the one on its own; of several there, Excel's order is not understood."""
    from pyopenvba.apps.excel._precedents import _readers  # pyright: ignore[reportPrivateUsage]

    away: list[Area] = []
    for owner in sheet.book.sheets_:
        if owner is sheet:
            continue
        for (cell_row, cell_column), cell in owner.cells_.items():
            if cell.formula and any(
                    area.sheet and area.sheet.casefold() == sheet.name.casefold()
                    and area.top <= row <= area.bottom and area.left <= column <= area.right
                    for area in _elsewhere(owner, cell.formula)):
                away.append(Area(cell_row, cell_column, cell_row, cell_column, owner.name))
    local = [cell for cell, _area, areas in _readers(sheet)
             if any(area.top <= row <= area.bottom and area.left <= column <= area.right for area in areas)]
    if len(local) > 1:
        raise VBAUnsupportedError('NavigateArrow among several dependents on a sheet is not implemented')
    arrows = [away] if away else []
    return arrows + [[Area(cell[0], cell[1], cell[0], cell[1])] for cell in local]


def navigate_arrow(target: Range, toward: object, arrow: object, link: object) -> object:
    """The cells an arrow leads to, selected."""
    from pyopenvba.apps.excel._model import Range

    target.tracing()
    sheet = target.sheet
    first = target.first
    row, column = first.top, first.left
    # It goes to the cell first, so what fails leaves it there.
    answer = Range(sheet, [Area(row, column, row, column)])
    sheet.Activate()
    answer.Select()
    if toward is MISSING or arrow is MISSING:
        raise error(1004, 'NavigateArrow method of Range class failed')
    precedents = to_bool(toward)
    if precedents:
        arrows = _precedent_arrows(sheet, row, column) if _traced(sheet, row, column) else []
    else:
        arrows = _dependent_arrows(sheet, row, column) if (row, column, True) in sheet.traced else []
    number = int(to_integer(arrow, 'Long'))
    place = 1 if link is MISSING else int(to_integer(link, 'Long'))
    if 1 <= number <= len(arrows):
        links = arrows[number - 1]
        if not 1 <= place <= len(links):
            raise error(1004, 'NavigateArrow method of Range class failed')
        area = links[place - 1]
        owner = sheet.book.sheet_named(area.sheet) if area.sheet else sheet
        answer = Range(owner, [Area(area.top, area.left, area.bottom, area.right)])
        owner.Activate()
        answer.Select()
    return answer


def _broken(words: list[str], fits: int, width: dict[str, int]) -> list[str]:
    lines: list[str] = []
    line = ''
    for word in words:
        trial = f'{line} {word}' if line else word
        wide = sum(width[char] for char in trial)
        if line and fits < wide <= fits * 6 // 5:
            raise VBAUnsupportedError('Justify on a line about as wide as its column is not implemented')
        if line and wide > fits:
            lines.append(line)
            line = word
        else:
            line = trial
    return [*lines, line] if line else lines


def justify(target: Range) -> object:
    from pyopenvba.apps.excel._column_fit import _table  # pyright: ignore[reportPrivateUsage]
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._typing import Typed

    sheet = target.sheet
    first = target.first
    if len(target.areas) != 1:
        raise VBAUnsupportedError('Justify on several areas is not implemented')
    pieces: list[str] = []
    for row in range(first.top, first.bottom + 1):
        cell = sheet.cells_.get((row, first.left))
        if cell is None or cell.value is EMPTY or cell.value == '':
            continue
        if cell.formula or not isinstance(cell.value, str):
            raise error(1004, 'Justify method of Range class failed')
        pieces.append(cell.value)
    text = ' '.join(pieces)
    if '  ' in text or text != text.strip():
        raise VBAUnsupportedError('Justify on text with runs of spaces is not implemented')
    table = _table(sheet.style_at(first.top, first.left).font)
    if table is None or any(char not in table for char in text):
        raise VBAUnsupportedError('Justify in a font or with characters whose widths are not known is not implemented')
    fits = sum(sheet.dims.column_pixels(column) for column in range(first.left, first.right + 1))
    lines = _broken(text.split(' '), fits, table)
    for offset in range(max(len(lines), first.bottom - first.top + 1)):
        # What it writes is text, a word that looks like a number included.
        if offset < len(lines):
            target.store_typed(first.top + offset, first.left, Typed(lines[offset]))
        else:
            Range(sheet, [Area(first.top + offset, first.left, first.top + offset, first.left)]).vba_set('Value', EMPTY)
    return True
