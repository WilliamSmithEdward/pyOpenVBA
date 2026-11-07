"""The second round for scripts/measure_shape_lines.py: tests/fixtures/shape_lines2_predictions.json, written before Excel is asked.

Fresh texts at the widths where the 1/200-point rule and the exact design widths disagree, so the rounding is what is tested, then other margins, a text box and no wrapping.
"""
from __future__ import annotations

import json
import random
from pathlib import Path

from pyopenvba.apps.excel._font_advances import advances
from pyopenvba.apps.excel._text_lines import line_spans

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'shape_lines2_predictions.json'
WORDS = ('budget', 'forecast', 'margin', 'Revenue', 'Total', 'invoice', 'quantity', 'supplier', 'Monday', 'pending', 'approved', 'zebra', 'jazz', 'Fixed', 'gamma', 'kilo', 'river', 'Stone', 'cloud', 'mix-up', 'x-ray', '42', '3.14', '(draft)', 'R&D', 'naive')
FONTS = ('Aptos Narrow', 'Calibri')
SIZES = (8.0, 8.5, 9.0, 9.5, 10.0, 10.5, 11.0, 11.5, 12.0, 13.0, 14.0, 15.5, 16.0, 18.0, 20.0, 22.0, 26.0, 28.0)
PIXELS = range(24, 700)
MARGINS = ((7.2, 7.2), (0.0, 0.0), (3.3, 10.05), (14.4, 1.5), (5.0, 5.0))


def _exact(text: str, room: float, units: dict[str, int], em: int, size: float) -> list[tuple[int, int]]:
    """The rival: the same breaks at design widths with nothing rounded, at a very fine step."""
    return line_spans(text, room * 20000 / 200, units, em, size * 100)


def main() -> None:
    chooser = random.Random(11)
    cases: list[dict[str, object]] = []
    split = 0
    while split < 400:
        font, size, bold, italic = chooser.choice(FONTS), chooser.choice(SIZES), chooser.random() < .3, chooser.random() < .3
        text = ' '.join(chooser.choices(WORDS, k=chooser.randint(2, 7)))
        table = advances(font, bold, italic)
        assert table is not None
        em, units = table
        pixels = [pixel for pixel in PIXELS if line_spans(text, pixel * .75 - 14.4, units, em, size) != _exact(text, pixel * .75 - 14.4, units, em, size)]
        if not pixels:
            continue
        split += len(pixels)
        cases.append({'text': text, 'font': font, 'size': size, 'bold': bold, 'italic': italic, 'kind': 'rectangle', 'margins': [7.2, 7.2], 'wrap': True, 'pixels': pixels,
                      'lines': [line_spans(text, pixel * .75 - 14.4, units, em, size) for pixel in pixels], 'rival': [_exact(text, pixel * .75 - 14.4, units, em, size) for pixel in pixels]})
    texts = ('The figures for the year were checked twice', 'state-of-the-art tooling and more\nsecond paragraph here', 'Supercalifragilistic word')
    for margins in MARGINS:
        for kind in ('rectangle', 'text box'):
            for text in texts:
                for font, size in (('Aptos Narrow', 11.0), ('Calibri', 13.0)):
                    table = advances(font, False, False)
                    assert table is not None
                    em, units = table
                    pixels = sorted(chooser.sample(range(max(int(sum(margins) / .75) + 4, 24), 500), 14))
                    cases.append({'text': text, 'font': font, 'size': size, 'bold': False, 'italic': False, 'kind': kind, 'margins': list(margins), 'wrap': True, 'pixels': pixels,
                                  'lines': [line_spans(text, pixel * .75 - sum(margins), units, em, size) for pixel in pixels]})
    for text in texts:
        table = advances('Calibri', False, False)
        assert table is not None
        pixels = [40, 80, 160, 320]
        cases.append({'text': text, 'font': 'Calibri', 'size': 11.0, 'bold': False, 'italic': False, 'kind': 'rectangle', 'margins': [7.2, 7.2], 'wrap': False, 'pixels': pixels, 'lines': [line_spans(text, None, table[1], table[0], 11.0) for _ in pixels]})
    FIXTURE.write_text(json.dumps({'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    print(len(cases), 'cases,', sum(len(case['pixels']) for case in cases), 'widths,', split, 'where the rounding decides')  # type: ignore[arg-type]


if __name__ == '__main__':
    main()
