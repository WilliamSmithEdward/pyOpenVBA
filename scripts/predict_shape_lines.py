"""Write down where text breaks in a shape, before Excel is asked: tests/fixtures/shape_lines_predictions.json, for scripts/measure_shape_lines.py.

For fresh texts, fonts and sizes, at the widths either side of each place the count changes and a few others, in whole pixels.
"""
from __future__ import annotations

import json
import random
from pathlib import Path

from pyopenvba.apps.excel._font_advances import advances
from pyopenvba.apps.excel._text_lines import line_spans

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'shape_lines_predictions.json'
TEXTS = ('hello world', 'Quarterly sales report for the northern region', 'antidisestablishmentarianism', 'AVAVAVAV WAVE To Ty', 'a  b   c', ' leading space here', 'well-known state-of-the-art',
         'trailing-', 'tail-\n', 'one\ntwo three\n', '\nstart', 'a\n\nb', 'x\x0by z', 'iiiii lllll', 'WWWWW MMMMM', '1234567890 1,234.56', 'cost: $100 (net) / 50%', 'end space   ', 'caf\xe9 na\xefve r\xe9sum\xe9', 'short lines and long ones\nwrap on their own')
FONTS = (('Aptos Narrow', 11.0, False, False), ('Aptos Narrow', 9.0, False, False), ('Aptos Narrow', 11.0, True, False), ('Aptos Narrow', 14.0, False, True), ('Aptos Narrow', 20.0, True, True), ('Aptos Narrow', 10.5, False, False),
         ('Calibri', 8.0, False, False), ('Calibri', 11.0, False, False), ('Calibri', 12.0, False, False), ('Calibri', 14.0, True, False), ('Calibri', 18.0, False, True), ('Calibri', 24.0, False, False))
MARGINS = 14.4
PIXELS = range(20, 800)


def main() -> None:
    chooser = random.Random(5)
    cases = []
    for text in TEXTS:
        for font, size, bold, italic in FONTS:
            table = advances(font, bold, italic)
            assert table is not None
            em, units = table

            def spans(pixel: int) -> list[tuple[int, int]]:
                return line_spans(text, pixel * 0.75 - MARGINS, units, em, size)

            changes = [pixel for pixel in PIXELS if pixel > PIXELS[0] and spans(pixel) != spans(pixel - 1)]
            chosen = {near for pixel in changes for near in (pixel - 1, pixel)} | set(chooser.sample(PIXELS, 6))
            kept = sorted(pixel for pixel in chosen if pixel in PIXELS)
            if len(kept) > 40:
                kept = sorted(chooser.sample(kept, 40))
            cases.append({'text': text, 'font': font, 'size': size, 'bold': bold, 'italic': italic, 'pixels': kept, 'lines': [spans(pixel) for pixel in kept]})
    FIXTURE.write_text(json.dumps({'margins': MARGINS, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    print(len(cases), 'cases,', sum(len(case['pixels']) for case in cases), 'widths')


if __name__ == '__main__':
    main()
