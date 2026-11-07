"""Score slant rules for italic wrapped text on every font round's answers, into tests/fixtures/wrapped_italic_scores.json. Needs Windows for GDI's glyph overhangs, and the font rounds' fixtures
(tests/fixtures/font_round_*.json, measured at 96 DPI by scripts/font_round.py).

The rule held out, a fifth of the font's ascent taken from the column, is scored with its rivals. Each rival changes only the room a line needs; everything else is the model's own lines_of.
"""
from __future__ import annotations

import ctypes
import importlib.util
import json
from collections import Counter
from collections.abc import Callable
from functools import cache
from pathlib import Path

from pyopenvba.apps.excel import _font_rows
from pyopenvba.apps.excel import _wrapped_text as wrapped
from pyopenvba.apps.excel._column_fit import font_pad
from pyopenvba.apps.excel._dimensions import font_pixels
from pyopenvba.apps.excel._font_widths import ascent, widths

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / 'tests' / 'fixtures'
ROUNDS = ('arial', 'times_new_roman', 'arial2', 'arial3', 'times_new_roman3', 'aptos_narrow3', 'calibri3')


class _ABC(ctypes.Structure):
    _fields_ = [('a', ctypes.c_int), ('b', ctypes.c_uint), ('c', ctypes.c_int)]


def _gdi() -> object:
    spec = importlib.util.spec_from_file_location('bake', ROOT / 'scripts' / 'bake_font_widths.py')
    assert spec is not None and spec.loader is not None
    bake = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(bake)
    found = bake._GDI()  # noqa: SLF001
    found.gdi32.GetCharABCWidthsW.argtypes = [found.wintypes.HDC, ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(_ABC)]
    return found


GDI = _gdi()


@cache
def overhang(face: str, bold: bool, pixels: int, character: str) -> int:
    """How far an italic glyph reaches past its advance, by GDI."""
    logical = GDI.LOGFONTW(lfHeight=-pixels, lfWeight=700 if bold else 400, lfItalic=1, lfCharSet=1, lfFaceName=face)  # type: ignore[attr-defined]
    font = GDI.gdi32.CreateFontIndirectW(ctypes.byref(logical))  # type: ignore[attr-defined]
    old = GDI.gdi32.SelectObject(GDI.dc, font)  # type: ignore[attr-defined]
    abc = _ABC()
    GDI.gdi32.GetCharABCWidthsW(GDI.dc, ord(character), ord(character), ctypes.byref(abc))  # type: ignore[attr-defined]
    GDI.gdi32.SelectObject(GDI.dc, old)  # type: ignore[attr-defined]
    GDI.gdi32.DeleteObject(font)  # type: ignore[attr-defined]
    return max(0, -abc.c)


def answers() -> list[tuple[str, str, float, bool, bool, int, str, int, dict[str, int], int, str]]:
    """Every wrapped answer below Excel's tallest row: text, font, size, bold, italic, indent, which reading, the column's pixels, the widths, Excel's lines, the round."""
    rows = []
    for stem in ROUNDS:
        cases = json.loads((FIXTURES / f'font_round_{stem}_predictions.json').read_text('utf-8'))['cases']
        found = json.loads((FIXTURES / f'font_round_{stem}.json').read_text('utf-8'))['answers']
        for case, answer in zip(cases, found, strict=True):
            call = case['call']
            if call[0] != 'Wrapped' or answer.startswith('ERROR'):
                continue
            _, text, font, size, bold, italic, indent, asked = call
            table = widths(font, bold, italic, size)
            assert table is not None
            line, _ = _font_rows.row_of(_font_rows.table_of(font, bold, italic), font_pixels(size), alone=True)
            for width, group in zip(asked.split(','), answer.split('|')[1:], strict=True):
                heights = group.split(',')
                for reading, column_width, height in (('after the value', float(width), heights[0]), ('after AutoFit', float(width) + 2, heights[2])):
                    pixels = round(float(height) * 4 / 3)
                    if pixels < 546:
                        rows.append((text.replace('~', '\n'), font, size, bold, italic, indent, reading, int(column_width * 7 + 5 + .5), table, max(1, pixels // line), stem))
    return rows


Rise = Callable[[str, float, bool], int]
Extra = Callable[[str, float, bool, str], int]


def _fifth(font: str, size: float, bold: bool) -> int:
    found = ascent(font, bold, True, size)
    assert found is not None
    return found // 5


RIVALS: dict[str, tuple[Rise, Extra]] = {
    'a fifth of the ascent, held out': (_fifth, lambda f, s, b, shown: 0),
    'a fifth of the ascent, rounded up': (lambda f, s, b: -(-(ascent(f, b, True, s) or 0) // 5), lambda f, s, b, shown: 0),
    'the last glyph overhang alone': (lambda f, s, b: 0, lambda f, s, b, shown: overhang(f, b, font_pixels(s), shown[-1])),
    'a fifth plus the last glyph overhang': (_fifth, lambda f, s, b, shown: overhang(f, b, font_pixels(s), shown[-1])),
    'the more of a fifth and the overhang': (lambda f, s, b: 0, lambda f, s, b, shown: max(_fifth(f, s, b), overhang(f, b, font_pixels(s), shown[-1]))),
}


def main() -> None:
    rows = answers()
    original = wrapped._fits  # noqa: SLF001
    scores: dict[str, object] = {}
    misses: list[dict[str, object]] = []
    for name, (rise, extra) in RIVALS.items():
        wrong: Counter[str] = Counter()
        total = 0
        for text, font, size, bold, italic, indent, reading, columns, table, count, stem in rows:
            if not italic:
                continue
            total += 1

            def fits(piece: str, room: int, extent: wrapped.Extent, pad: int, font: str = font, size: float = size, bold: bool = bold) -> bool:
                shown = piece.rstrip(' ').removesuffix('-')
                return room >= wrapped.room_for(extent(shown), pad) + (extra(font, size, bold, shown) if shown else 0)

            wrapped._fits = fits  # noqa: SLF001
            got = wrapped.lines_of(text, columns - 9 * indent - rise(font, size, bold), lambda s, table=table: sum(table[c] for c in s), font_pad(table))
            wrapped._fits = original  # noqa: SLF001
            if got != count:
                wrong[font] += 1
                if name.endswith('held out'):
                    misses.append({'font': font, 'size': size, 'bold': bold, 'indent': indent, 'reading': reading, 'column_pixels': columns, 'model_lines': got, 'excel_lines': count, 'round': stem})
        scores[name] = {'wrong': sum(wrong.values()), 'of': total, 'by_font': dict(sorted(wrong.items()))}
    upright = [row for row in rows if not row[4]]
    upright_wrong = sum(wrapped.lines_of(text, columns - 9 * indent, lambda s, table=table: sum(table[c] for c in s), font_pad(table)) != count
                        for text, font, size, bold, italic, indent, reading, columns, table, count, stem in upright)
    record = {'rounds': [f'tests/fixtures/font_round_{stem}.json' for stem in ROUNDS], 'italic': scores, 'upright': {'wrong': upright_wrong, 'of': len(upright)}, 'held_out_misses': misses}
    (FIXTURES / 'wrapped_italic_scores.json').write_text(json.dumps(record, indent=1), encoding='utf-8', newline='\n')
    print(json.dumps({key: value for key, value in record.items() if key != 'held_out_misses'}, indent=1))


if __name__ == '__main__':
    main()
