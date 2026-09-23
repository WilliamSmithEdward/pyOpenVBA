"""Range.Find, FindNext, FindPrevious and Replace, and the search settings they share.

Find looks in one area. Among formulas it reads a formula's text and a
constant's text as the formula bar shows it; among values, what the cell
shows. Replace always edits that formula-bar text (scripts/measure_replace.py):

- ``?`` is any one character, ``*`` any run of them -- as short as will
  match, except at the end of what is sought, where it runs to the end --
  and ``~`` takes the next character as it is. A ``~`` with nothing after
  it is dropped, and a search for nothing replaces nothing.
- Every match in a cell is replaced, left to right, and what is left is
  typed again as Replace types it (see ``_typing.replaced``). A formula
  Excel cannot read stops the Replace at that cell, the cells before it
  changed and none after.
- By rows the cells are taken area by area, row by row, from the first;
  by columns, column by column, starting after the first cell of the
  first area and coming back to it last. A cell two areas share is
  replaced twice.
- Replace always answers True. It leaves its LookAt, SearchOrder,
  MatchCase and what it sought for the next Find and FindNext, and takes
  LookAt and SearchOrder from the last one; MatchCase is False unless it
  is given.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
from pyopenvba.apps.excel._visible import visible
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._values import MISSING, NOTHING, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range


@dataclass
class FindState:
    look_in: int = -4123
    look_at: int = 2
    order: int = 1
    what: object = MISSING
    match_case: bool = False


def _parts(text: str) -> list[str]:
    """What is sought, as pieces of a regular expression; a star at the end runs to the end of the text."""
    parts: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == "~":
            if index + 1 == len(text):
                break
            index += 1
            parts.append(re.escape(text[index]))
        else:
            parts.append(".*?" if char == "*" else "." if char == "?" else re.escape(char))
        index += 1
    if parts and parts[-1] == ".*?":
        parts[-1] = ".*"
    return parts


def _pattern(text: str, whole: bool, match_case: bool) -> re.Pattern[str]:
    body = "".join(_parts(text))
    if whole:
        body = r"\A(?:" + body + r")\Z"
    return re.compile(body, re.DOTALL | (0 if match_case else re.IGNORECASE))


def _sought(value: object) -> str:
    return ("TRUE" if value else "FALSE") if isinstance(value, bool) else to_text(value)


def find(target: Range, what: object, after: object = MISSING, look_in: object = MISSING,
         look_at: object = MISSING, order: object = MISSING, direction: object = MISSING,
         match_case: object = MISSING, match_byte: object = MISSING,
         search_format: object = MISSING, *, again: bool = False) -> object:
    from pyopenvba.apps.excel._model import Range
    from pyopenvba.apps.excel._typing import edit_text

    if len(target.areas) != 1:
        raise VBAUnsupportedError("Range.Find on multiple areas is not implemented")
    if match_byte is not MISSING and to_bool(match_byte):
        raise VBAUnsupportedError("Range.Find MatchByte=True is not implemented")
    if search_format is not MISSING and to_bool(search_format):
        raise VBAUnsupportedError("Range.Find SearchFormat=True is not implemented")
    state = target.sheet.book.application.find_state
    source = state.look_in if look_in is MISSING else int(to_integer(look_in, "Long"))
    whole = state.look_at if look_at is MISSING else int(to_integer(look_at, "Long"))
    traversal = state.order if order is MISSING else int(to_integer(order, "Long"))
    if source in {-4144, -4184}:
        raise VBAUnsupportedError("Range.Find in comments is not implemented")
    if source not in {-4123, -4163} or whole not in {1, 2} or traversal not in {1, 2}:
        raise error(9)
    area = target.first
    anchor = (area.top, area.left)
    if after is not MISSING:
        if not isinstance(after, Range) or not after.single or after.sheet is not target.sheet:
            raise error(13)
        anchor = (after.first.top, after.first.left)
        if not area.contains(*anchor):
            raise error(13)
    if again:
        what, case = state.what, state.match_case
        if what is MISSING:
            return NOTHING
    else:
        case = False if match_case is MISSING else to_bool(match_case)
    text = _sought(what)
    pattern = _pattern(text, whole == 1, case)
    state.look_in, state.look_at, state.order = source, whole, traversal
    state.what, state.match_case = what, case
    if text == "~":
        # Excel's empty escaped pattern returns After itself, without moving.
        row, column = anchor
        return Range(target.sheet, [Area(row, column, row, column, target.sheet.name)])
    backwards = direction is not MISSING and int(to_integer(direction, "Long")) == 2
    step = -1 if backwards else 1
    total = area.rows * area.columns

    def ordinal(position: tuple[int, int]) -> int:
        row, column = position
        if traversal == 1:
            return (row - area.top) * area.columns + column - area.left
        return (column - area.left) * area.rows + row - area.top

    def position(index: int) -> tuple[int, int]:
        if traversal == 1:
            row, column = divmod(index, area.columns)
        else:
            column, row = divmod(index, area.rows)
        return area.top + row, area.left + column

    start = ordinal(anchor)
    # Nonempty searches examine only stored cells, even on whole-sheet ranges.
    # Blank searches walk lazily from After and stop at the first empty cell.
    if text == "":
        positions = (position((start + step * distance) % total) for distance in range(1, total + 1))
    else:
        candidates = [pos for pos in target.sheet.cells_ if area.contains(*pos)]
        candidates.sort(key=lambda pos: (step * (ordinal(pos) - start) - 1) % total)
        positions = iter(candidates)
    for row, column in positions:
        cell = target.sheet.cell(row, column)
        found = Range(target.sheet, [Area(row, column, row, column, target.sheet.name)])
        if cell is not None and source == -4123:
            # Among formulas a constant reads as the formula bar shows it: 1234 in #,##0 is 1234, not 1,234.
            value = cell.formula or edit_text(cell.value, cell.number_format)
        else:
            value = to_text(found.Text())
            if cell is not None and isinstance(cell.value, bool):
                value = "TRUE" if cell.value else "FALSE"
        matches = value == "" if text == "" else bool(value and pattern.search(value))
        if matches:
            return found
    return NOTHING


def replace(target: Range, what: object, replacement: object, look_at: object = MISSING,
            order: object = MISSING, match_case: object = MISSING, match_byte: object = MISSING,
            search_format: object = MISSING, replace_format: object = MISSING,
            formula_version: object = MISSING) -> object:
    from pyopenvba.apps.excel._typing import edit_text

    if what is MISSING or replacement is MISSING:
        raise error(449)
    if match_byte is not MISSING and to_bool(match_byte):
        raise VBAUnsupportedError("Range.Replace MatchByte=True is not implemented")
    for name, flag in (("SearchFormat", search_format), ("ReplaceFormat", replace_format)):
        if flag is not MISSING and to_bool(flag):
            raise VBAUnsupportedError(f"Range.Replace {name}=True is not implemented")
    if formula_version is not MISSING and int(to_integer(formula_version, "Long")) != 0:
        raise VBAUnsupportedError("Range.Replace with Formula2 semantics is not implemented")
    state = target.sheet.book.application.find_state
    whole = state.look_at if look_at is MISSING else int(to_integer(look_at, "Long"))
    traversal = state.order if order is MISSING else int(to_integer(order, "Long"))
    if whole not in {1, 2} or traversal not in {1, 2}:
        raise error(9)
    case = False if match_case is MISSING else to_bool(match_case)
    state.look_in, state.look_at, state.order, state.what, state.match_case = -4123, whole, traversal, what, case
    text, new = _sought(what), _sought(replacement)
    if not _parts(text):
        return True
    pattern = _pattern(text, whole == 1, case)
    # On a filtered sheet only the visible cells change, as _visible has it.
    target = visible(target)
    sheet = target.sheet
    cells: list[tuple[int, int]] = []
    for area in target.areas:
        inside = [position for position in sheet.cells_ if area.contains(*position)]
        inside.sort(key=(lambda position: position) if traversal == 1 else (lambda position: position[::-1]))
        cells += inside
    first = (target.first.top, target.first.left)
    if traversal == 2 and cells and cells[0] == first:
        # By columns Excel starts after the first cell and comes back to it last.
        cells.append(cells.pop(0))
    for row, column in cells:
        cell = sheet.cells_.get((row, column))
        if cell is None:
            continue
        before = cell.formula or edit_text(cell.value, cell.number_format)
        after = _substituted(pattern, before, new) if before else None
        if after is not None and not target.replaced_in(row, column, after):
            break
    return True


def _substituted(pattern: re.Pattern[str], text: str, replacement: str) -> str | None:
    """``text`` with every match of ``pattern`` replaced, left to right, or None where nothing matches."""
    pieces: list[str] = []
    at = 0
    while at < len(text):
        found = pattern.search(text, at)
        if found is None or found.end() == found.start():
            break
        pieces += [text[at:found.start()], replacement]
        at = found.end()
    return "".join(pieces) + text[at:] if pieces else None
