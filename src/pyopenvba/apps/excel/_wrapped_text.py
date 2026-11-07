"""How many lines wrapped text takes in a column, on a 96-DPI display, as held out against live Excel (scripts/measure_wrapped_heldout.py, scripts/measure_wrapped_curves.py,
scripts/measure_wrapped_round3.py, scripts/measure_wrapped_round4.py): 16,276 of 16,380 predictions across fonts and sizes, each written before Excel was asked. The last round ran on this
code, and its 55 misses, 8,465 of 8,520 right, all end in a hyphen and a line feed, which the model refuses.

The 7 and the steps of 28 are measured, and not explained. They held across 20 fonts and sizes with nothing per font but the width of the digit 0, in `font_pad`. The steps are the scale `_column_fit.column_pixels`
widens a column by, 29/28, and a line fits one pixel sooner than that column at a whole multiple of 28; why Excel has that scale is not known.
"""
from __future__ import annotations

from collections.abc import Callable

Extent = Callable[[str], int]


def room_for(wide: int, pad: int) -> int:
    """The narrowest column a line this many pixels wide stays on one line in, in a font that adds ``pad``."""
    return wide + 7 + pad + max(0, (wide - 57) // 28)


def _fits(text: str, column: int, extent: Extent, pad: int) -> bool:
    """Whether a line fits. Spaces at its end and a hyphen there take no room."""
    return column >= room_for(extent(text.rstrip(' ').removesuffix('-')), pad)


def pieces(text: str) -> list[str]:
    """Where a line may break: after spaces, and after a hyphen with something after it."""
    found: list[str] = []
    current = ''
    for place, character in enumerate(text):
        current += character
        following = text[place + 1] if place + 1 < len(text) else ''
        if (character == ' ' and following != ' ') or (character == '-' and following not in ('', ' ')):
            found.append(current)
            current = ''
    return [*found, current] if current else found


def _paragraph_lines(text: str, column: int, extent: Extent, pad: int) -> int:
    lines, line = 0, ''
    for place, piece in enumerate(pieces(text)):
        if line and not _fits(line + piece, column, extent, pad):
            lines += 1
            line = ''
        # A word too long for a line breaks between characters, where a hyphen takes its own room. The text's first letter keeps a hyphen after it.
        if not _fits(piece, column, extent, pad) and (place or len(piece.rstrip(' ').removesuffix('-')) > 1):
            while len(piece.rstrip(' ')) > 1 and not _fits(piece, column, extent, pad):
                part = piece[0]
                for character in piece[1:]:
                    if column < room_for(extent(part + character), pad):
                        break
                    part += character
                lines += 1
                piece = piece[len(part):]
        line += piece
    return lines + 1


def lines_of(text: str, column: int, extent: Extent, pad: int) -> int:
    """How many lines a text takes in ``column`` pixels, less what an indent or italic takes. Only a line feed starts a line of its own."""
    found = sum(_paragraph_lines(one, column, extent, pad) for one in text.split('\n'))
    # A line feed at the very end takes a line of its own when the letter before it does not fit either.
    if text.endswith('\n') and text[-2:-1] not in ('', '\n') and column < room_for(extent(text[-2]), pad):
        found += 1
    return found


def row_pixels_for(lines: int, line: int, descent: int) -> int:
    """How tall lines of a font make a row: the font alone on a row for each, less what a descent short of 5 pixels leaves, to Excel's tallest row
    (tests/fixtures/font_round_arial.json, tests/fixtures/font_round_times_new_roman.json)."""
    return min(lines * line + max(0, 5 - descent), 546)