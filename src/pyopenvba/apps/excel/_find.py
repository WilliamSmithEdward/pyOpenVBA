"""Measured single-area Range.Find matching and application search state."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from pyopenvba._a1 import Area
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


def _pattern(text: str, whole: bool, match_case: bool) -> re.Pattern[str]:
    parts: list[str] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char == "~" and index + 1 == len(text):
            break
        if char == "~" and index + 1 < len(text):
            index += 1
            parts.append(re.escape(text[index]))
        else:
            parts.append(".*" if char == "*" else "." if char == "?" else re.escape(char))
        index += 1
    body = "".join(parts)
    if whole:
        body = r"\A(?:" + body + r")\Z"
    return re.compile(body, re.DOTALL | (0 if match_case else re.IGNORECASE))


def find(target: Range, what: object, after: object = MISSING, look_in: object = MISSING,
         look_at: object = MISSING, order: object = MISSING, direction: object = MISSING,
         match_case: object = MISSING, match_byte: object = MISSING,
         search_format: object = MISSING, *, again: bool = False) -> object:
    from pyopenvba.apps.excel._model import Range

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
    text = ("TRUE" if what else "FALSE") if isinstance(what, bool) else to_text(what)
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
        if cell is not None and source == -4123 and cell.formula:
            value = cell.formula
        else:
            value = to_text(found.Text())
            if cell is not None and isinstance(cell.value, bool):
                value = "TRUE" if cell.value else "FALSE"
        matches = value == "" if text == "" else bool(value and pattern.search(value))
        if matches:
            return found
    return NOTHING
