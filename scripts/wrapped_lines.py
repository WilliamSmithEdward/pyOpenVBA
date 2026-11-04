"""The word wrap under test, for scripts that predict wrapped rows: how many lines a text takes in a column, given how wide a piece of it is. See tests/fixtures/wrapped_rows.json."""
from __future__ import annotations

from collections.abc import Callable


def room_for(wide: int, digit: int) -> int:
    """The column a line this wide needs: its width, the digit-width pad, and a pixel for each 28 past 57."""
    return wide + 7 + 2 * max(0, (digit - 5) // 4) + max(0, (wide - 57) // 28)


def fits(text: str, column: int, extent: Callable[[str], int], digit: int) -> bool:
    """Whether a line fits. Spaces at its end and a hyphen there do not count."""
    return column >= room_for(extent(text.rstrip(' ').removesuffix('-')), digit)


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


def lines_of(text: str, column: int, extent: Callable[[str], int], digit: int) -> int:
    lines, line = 0, ''
    for piece in pieces(text):
        if line and not fits(line + piece, column, extent, digit):
            lines += 1
            line = ''
        # A word too long for a line breaks between characters, where a hyphen takes its own room.
        if not fits(piece, column, extent, digit) and len(piece.rstrip(' ').removesuffix('-')) > 1:
            while len(piece.rstrip(' ')) > 1 and column < room_for(extent(piece.rstrip(' ')), digit):
                part = piece[0]
                for character in piece[1:]:
                    if column < room_for(extent(part + character), digit):
                        break
                    part += character
                lines += 1
                piece = piece[len(part):]
        line += piece
    return lines + 1


def height_of(lines: int, line: float) -> float:
    """A row's height from its lines, each as tall as the font alone makes a row. A font shorter than the standard row adds a pixel once there are two, and a row stops at 409.5."""
    if lines == 1:
        return max(line, 15.0)
    return min(lines * line + (0.75 if line < 15 else 0.0), 409.5)
