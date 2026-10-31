"""Range.Consolidate, and the ConsolidationFunction, ConsolidationOptions and ConsolidationSources a sheet keeps,
as Excel consolidates (tests/fixtures/consolidate.json, scripts/measure_consolidate.py).

- Sources are R1C1 references with their sheet, Sheet1!R2C2:R4C3, or defined names; an A1 reference is 1004
  "Reference isn't valid.", the same reference twice 1004 "Cannot add duplicate source reference.". A sheet that
  is not there gives nothing. The sheet keeps them in A1, Sheet1!$B$2:$C$4, with the function and the TopRow,
  LeftColumn and CreateLinks options, and a later Consolidate takes whatever it leaves out from there.
- By position, cell (i, j) of the result is the function over cell (i, j) of every source, the result as tall
  and wide as the largest source. With TopRow the first row of each source names its columns, and columns of
  the same name, ignoring case, are one, in the order the names first appear, spelt as they first appear; a
  column with no name is left out. LeftColumn does the same for rows. The other direction stays by position.
- Every cell of the result is written: the function's value with the number format of the first source cell
  it took in, or cleared where no source cell counts. With both labels the corner is left as it was. A
  destination of more than one cell cuts the result to its size.
- The functions take numbers, a date among them, and an error stops them with it; xlCount counts text and
  logical values too. StDev and Var of one number are #DIV/0!.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area, column_letter
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError
from pyopenvba.interpreter._values import EMPTY, MISSING, VBAArray, VBAErrorValue, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Worksheet

SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNT_NUMS = -4157, -4112, -4106, -4136, -4139, -4149, -4113
STDEV, STDEVP, VAR, VARP = -4155, -4156, -4164, -4165
FUNCTIONS = frozenset({SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNT_NUMS, STDEV, STDEVP, VAR, VARP})
#: CVErr(xlErrDiv0).
DIV0 = 2007
_FAILED = "Consolidate method of Range class failed"
_REFERENCE = re.compile(r"R([0-9]+)C([0-9]+)(?::R([0-9]+)C([0-9]+))?", re.IGNORECASE)
_COLUMNS = re.compile(r"C[0-9]+(?::C[0-9]+)?|R[0-9]+(?::R[0-9]+)?", re.IGNORECASE)


@dataclass
class Settings:
    """What a sheet keeps of its last consolidation."""

    function: int = SUM
    top_row: bool = False
    left_column: bool = False
    create_links: bool = False
    sources: tuple[str, ...] = ()


def settings(sheet: Worksheet) -> Settings:
    kept = sheet.consolidation
    return kept if kept is not None else Settings()


def options(sheet: Worksheet) -> VBAArray:
    kept = settings(sheet)
    return VBAArray([(1, 3)], items=[kept.top_row, kept.left_column, kept.create_links])


def sources(sheet: Worksheet) -> object:
    kept = settings(sheet).sources
    return VBAArray([(1, len(kept))], items=list(kept)) if kept else EMPTY


@dataclass
class _Source:
    sheet: Worksheet | None
    area: Area | None


def _resolved(book: object, text: str, *, kept_before: bool = False) -> tuple[str, _Source]:
    """A source as the sheet keeps it, and the cells it names; ``kept_before`` for one the sheet kept, in A1."""
    from pyopenvba._a1 import column_number, split_sheet
    from pyopenvba.apps.excel import _names
    from pyopenvba.apps.excel._model import DefinedName, Workbook

    assert isinstance(book, Workbook)
    if "[" in text:
        raise VBAUnsupportedError("Consolidate from another workbook is not implemented")
    if "," in text:
        raise error(1004, "Reference isn't valid.")
    sheet_name, bare = split_sheet(text)
    if kept_before and sheet_name:
        bare = re.sub(r"\$([A-Z]+)\$([0-9]+)", lambda found: f"R{found.group(2)}C{column_number(found.group(1))}", bare)
    if not sheet_name:
        entry = next((one for one in book.names_.entries if _names.same_entry(one.name, text)), None)
        if entry is None:
            raise error(1004, "Reference isn't valid.")
        found = DefinedName(entry).refers_to_range()
        if len(found.areas) != 1:
            raise VBAUnsupportedError("Consolidate from a name of more than one area is not implemented")
        return text, _Source(found.sheet, found.first)
    matched = _REFERENCE.fullmatch(bare)
    if matched is None:
        if _COLUMNS.fullmatch(bare):
            raise VBAUnsupportedError("Consolidate from whole rows or columns is not implemented")
        raise error(1004, "Reference isn't valid.")
    top, left = int(matched.group(1)), int(matched.group(2))
    bottom, right = int(matched.group(3) or top), int(matched.group(4) or left)
    top, bottom, left, right = min(top, bottom), max(top, bottom), min(left, right), max(left, right)
    first, last = f"${column_letter(left)}${top}", f"${column_letter(right)}${bottom}"
    kept = sheet_name_text(text) + (first if first == last else f"{first}:{last}")
    try:
        sheet = book.sheet_named(sheet_name)
    except VBARuntimeError:  # A sheet that is not there gives nothing, as in Excel.
        return kept, _Source(None, None)
    return kept, _Source(sheet, Area(top, left, bottom, right))


def sheet_name_text(text: str) -> str:
    """The sheet part of a source as it was written, its ! included."""
    return text[:text.rindex("!") + 1]


def _texts(value: object) -> list[str]:
    if isinstance(value, VBAArray):
        items = list(value.items)
        if not items:
            raise error(13)
        return [to_text(item) for item in items]
    return [to_text(value)]


def consolidate(target: Range, sources_given: object, function: object, top_row: object, left_column: object,
                create_links: object) -> None:
    sheet = target.sheet
    kept = settings(sheet)
    texts = _texts(sources_given) if sources_given is not MISSING else None
    chosen = kept.function if function is MISSING else int(to_integer(function))
    if chosen not in FUNCTIONS:
        raise error(1004, _FAILED)
    top = kept.top_row if top_row is MISSING else to_bool(top_row)
    left = kept.left_column if left_column is MISSING else to_bool(left_column)
    links = kept.create_links if create_links is MISSING else to_bool(create_links)
    if links:
        raise VBAUnsupportedError("Consolidate with CreateLinks, which builds an outline of links, is not implemented")
    if texts is None:
        if not kept.sources:
            raise error(1004, "No source references specified for consolidation.")
        resolved = [_resolved(sheet.book, text, kept_before=True) for text in kept.sources]
    else:
        resolved = [_resolved(sheet.book, text) for text in texts]
    names = [name.lower() for name, _ in resolved]
    if len(set(names)) != len(names):
        raise error(1004, "Cannot add duplicate source reference.")
    sheet.consolidation = Settings(chosen, top, left, links, tuple(name for name, _ in resolved))
    found = [source for _, source in resolved if source.sheet is not None and source.area is not None]
    _write(target, chosen, top, left, found)


def _value(sheet: Worksheet, row: int, column: int) -> object:
    from pyopenvba.apps.excel._model import Range

    return Range(sheet, [Area(row, column, row, column)]).vba_get("Value2")


def _blank(value: object) -> bool:
    return value is EMPTY or value is None or value == ""


def _label(value: object) -> tuple[str, str] | None:
    """How two labels are told apart: text ignoring case, and a number by its value."""
    if _blank(value):
        return None
    if isinstance(value, str):
        return ("text", value.lower())
    return ("value", repr(value))


class _Grid:
    """The result: labels in the order they first appear, and the source cells that fall on each cell."""

    def __init__(self) -> None:
        self.rows: list[object] = []
        self.columns: list[object] = []
        self.row_keys: dict[tuple[str, str], int] = {}
        self.column_keys: dict[tuple[str, str], int] = {}
        self.cells: dict[tuple[int, int], list[tuple[Worksheet, int, int]]] = {}
        self.height = 0
        self.width = 0

    def line(self, keys: dict[tuple[str, str], int], labels: list[object], value: object) -> int | None:
        key = _label(value)
        if key is None:
            return None
        if key not in keys:
            keys[key] = len(labels)
            labels.append(value)
        return keys[key]


def _gathered(found: list[_Source], top: bool, left: bool) -> _Grid:
    grid = _Grid()
    for source in found:
        assert source.sheet is not None and source.area is not None
        sheet, area = source.sheet, source.area
        first_row, first_column = area.top + (1 if top else 0), area.left + (1 if left else 0)
        column_at: dict[int, int] = {}
        for column in range(first_column, area.right + 1):
            if top:
                at = grid.line(grid.column_keys, grid.columns, _value(sheet, area.top, column))
                if at is not None:
                    column_at[column] = at
            else:
                column_at[column] = column - first_column
        row_at: dict[int, int] = {}
        for row in range(first_row, area.bottom + 1):
            if left:
                at = grid.line(grid.row_keys, grid.rows, _value(sheet, row, area.left))
                if at is not None:
                    row_at[row] = at
            else:
                row_at[row] = row - first_row
        if not top:
            grid.width = max(grid.width, area.right - first_column + 1)
        if not left:
            grid.height = max(grid.height, area.bottom - first_row + 1)
        for row, i in row_at.items():
            for column, j in column_at.items():
                grid.cells.setdefault((i, j), []).append((sheet, row, column))
    if top:
        grid.width = len(grid.columns)
    if left:
        grid.height = len(grid.rows)
    return grid


def _number(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _worked(function: int, values: list[object]) -> tuple[object, int | None]:
    """The function over one cell's values, and which value it took first; None where none counts."""
    taken = [at for at, value in enumerate(values) if not _blank(value)
             and (function == COUNT or isinstance(value, VBAErrorValue) or _number(value) is not None)]
    if not taken:
        return EMPTY, None
    if function == COUNT:
        if any(isinstance(values[at], VBAErrorValue) for at in taken):
            raise VBAUnsupportedError("Consolidate counting an error value is not implemented")
        return float(len(taken)), taken[0]
    if function == COUNT_NUMS:
        numbers = [at for at in taken if _number(values[at]) is not None]
        if len(numbers) != len(taken):
            raise VBAUnsupportedError("Consolidate counting numbers among error values is not implemented")
        return float(len(numbers)), taken[0]
    failed = next((values[at] for at in taken if isinstance(values[at], VBAErrorValue)), None)
    if failed is not None:
        return failed, taken[0]
    numbers = [_number(values[at]) or 0.0 for at in taken]
    count = len(numbers)
    if function == SUM:
        return _sum(numbers), taken[0]
    if function == AVERAGE:
        return _sum(numbers) / count, taken[0]
    if function == MAX:
        return max(numbers), taken[0]
    if function == MIN:
        return min(numbers), taken[0]
    if function == PRODUCT:
        product = 1.0
        for number in numbers:
            product *= number
        return product, taken[0]
    sample = function in (STDEV, VAR)
    if sample and count < 2:
        return VBAErrorValue(DIV0), taken[0]
    mean = _sum(numbers) / count
    spread = _sum([(number - mean) ** 2 for number in numbers]) / (count - 1 if sample else count)
    return (math.sqrt(spread) if function in (STDEV, STDEVP) else spread), taken[0]


def _sum(numbers: list[float]) -> float:
    total = 0.0
    for number in numbers:
        total += number
    return total


def _write(target: Range, function: int, top: bool, left: bool, found: list[_Source]) -> None:
    from pyopenvba.apps.excel._model import Range

    grid = _gathered(found, top, left)
    if not grid.height or not grid.width:
        return
    sheet, start = target.sheet, target.first
    row0, column0 = start.top + (1 if top else 0), start.left + (1 if left else 0)
    limit_rows = start.rows if start.rows > 1 or start.columns > 1 else None
    limit_columns = start.columns if limit_rows is not None else None
    bottom, right = row0 + grid.height - 1, column0 + grid.width - 1
    if limit_rows is not None and limit_columns is not None:
        bottom, right = min(bottom, start.top + limit_rows - 1), min(right, start.left + limit_columns - 1)
    for source in found:
        area = source.area
        assert area is not None
        if source.sheet is sheet and not (area.bottom < start.top or area.top > bottom or area.right < start.left
                                          or area.left > right):
            raise VBAUnsupportedError("Consolidate into the cells of a source is not implemented")

    def put(row: int, column: int, value: object, number_format: str | None = None) -> None:
        if limit_rows is not None and limit_columns is not None and (
                row - start.top >= limit_rows or column - start.left >= limit_columns):
            return
        cell = Range(sheet, [Area(row, column, row, column)])
        cell.vba_set("Value", value)
        if number_format is not None:
            cell.vba_set("NumberFormat", number_format)

    if top:
        for j, label in enumerate(grid.columns):
            put(start.top, column0 + j, label)
    if left:
        for i, label in enumerate(grid.rows):
            put(row0 + i, start.left, label)
    for i in range(grid.height):
        for j in range(grid.width):
            places = grid.cells.get((i, j), [])
            values = [_value(owner, row, column) for owner, row, column in places]
            result, first = _worked(function, values)
            if first is None:
                put(row0 + i, column0 + j, EMPTY)
                continue
            owner, row, column = places[first]
            put(row0 + i, column0 + j, result, owner.style_at(row, column).number_format)
