"""Where text in a shape breaks into lines (TextRange2.Lines): greedily, at the font's design widths, against the box less its margins.

Unlike a cell, which measures whole GDI pixels (`_wrapped_text`), each glyph here is its design width rounded to 1/200 of a point. Fitted on tests/fixtures/wrapped_events.json and
tests/fixtures/shape_lines.json, and held out by scripts/measure_shape_lines.py.
"""
from __future__ import annotations

import math
import re

from pyopenvba.apps.excel._wrapped_text import pieces

_PARAGRAPH = re.compile(r"[^\n\x0b]*[\n\x0b]|[^\n\x0b]+$")
#: How finely a glyph's width is held, per point.
_STEPS = 200


def line_spans(text: str, room: float | None, units: dict[str, int], em: int, size: float) -> list[tuple[int, int]]:
    """Each line's start, from 0, and length, spaces and its line feed included. Spaces at a line's end take no room. No room is no wrapping."""
    steps = {character: math.floor(units[character] * size / em * _STEPS + 0.5) for character in set(text) - set("\n\x0b")}
    limit = None if room is None else math.floor(room * _STEPS + 1e-6)

    def fits(line: str) -> bool:
        return limit is None or sum(steps[character] for character in line.rstrip(" \n\x0b")) <= limit

    spans: list[tuple[int, int]] = []
    for paragraph in _PARAGRAPH.finditer(text):
        begin, line = paragraph.start(), ""
        whole = paragraph.group()
        body = whole.rstrip("\n\x0b")
        found = pieces(body) or [""]
        found[-1] += whole[len(body):]
        for piece in found:
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
