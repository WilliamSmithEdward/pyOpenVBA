"""Where text in a shape breaks into lines (TextRange2.Lines): greedily, at the font's design widths, against the box less its margins.

Unlike a cell, which measures whole GDI pixels (`_wrapped_text`), nothing here rounds. Fitted on tests/fixtures/wrapped_events.json and held out by
scripts/measure_shape_lines.py.
"""
from __future__ import annotations

import re

from pyopenvba.apps.excel._wrapped_text import pieces

_PARAGRAPH = re.compile(r"[^\n\x0b]*[\n\x0b]|[^\n\x0b]+$")


def line_spans(text: str, room: float, units: dict[str, int], em: int, size: float) -> list[tuple[int, int]]:
    """Each line's start, from 0, and length, spaces and its line feed included. Spaces at a line's end take no room."""

    def fits(line: str) -> bool:
        return sum(units[character] for character in line.rstrip(" \n\x0b")) * size / em <= room

    spans: list[tuple[int, int]] = []
    for paragraph in _PARAGRAPH.finditer(text):
        begin, line = paragraph.start(), ""
        for piece in pieces(paragraph.group()):
            if line and not fits(line + piece):
                spans.append((begin, len(line)))
                begin, line = begin + len(line), ""
            while not fits(piece):
                taken = 1
                while taken < len(piece) and fits(piece[:taken + 1]):
                    taken += 1
                if not piece[taken:].strip(" \n\x0b"):
                    break
                spans.append((begin, taken))
                begin, piece = begin + taken, piece[taken:]
            line += piece
        spans.append((begin, len(line)))
    return spans
