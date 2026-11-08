"""The third round for scripts/measure_shape_lines.py: tests/fixtures/shape_lines3_predictions.json, written before Excel is asked.

Arial and Times New Roman, which no earlier round asked: fresh texts in every style at the widths where the 1/200-point rule and the exact design widths
disagree, then fresh texts at random widths, other margins, a text box and no wrapping.
"""
from __future__ import annotations

import json
import random
from pathlib import Path

from pyopenvba.apps.excel._font_advances import advances
from pyopenvba.apps.excel._text_lines import line_spans

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'shape_lines3_predictions.json'
WORDS = ('ledger', 'Harbour', 'quarterly', 'Wexford', 'volume', 'tariff', 'WAREHOUSE', 'jetty', 'Oxbow', 'rate-card', 'pre-tax', 'illicit', 'fifty', 'Ampere', 'lily',
         'MMXX', '9%', '1,250.75', '(net)', 'A&B', 'zig-zag', 'Quince', 'caf\xe9', 'minimum', 'Wimble')
FONTS = ('Arial', 'Times New Roman')
SIZES = (8.0, 9.0, 10.0, 10.5, 11.0, 12.0, 12.5, 14.0, 16.0, 18.0, 20.0, 24.0, 28.0, 36.0)
PIXELS = range(24, 700)
MARGINS = ((7.2, 7.2), (0.0, 0.0), (3.3, 10.05), (14.4, 1.5))


def _exact(text: str, room: float, units: dict[str, int], em: int, size: float) -> list[tuple[int, int]]:
    """The rival: the same breaks at design widths with nothing rounded, at a very fine step."""
    return line_spans(text, room * 20000 / 200, units, em, size * 100)


def main() -> None:
    chooser = random.Random(31)
    cases: list[dict[str, object]] = []
    split = 0
    while split < 300:
        font, size, bold, italic = chooser.choice(FONTS), chooser.choice(SIZES), chooser.random() < .4, chooser.random() < .4
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
    for margins in MARGINS:
        for kind in ('rectangle', 'text box'):
            for font in FONTS:
                for bold, italic in ((False, False), (True, False), (False, True), (True, True)):
                    size = chooser.choice(SIZES)
                    text = ' '.join(chooser.choices(WORDS, k=chooser.randint(3, 9)))
                    if chooser.random() < .3:
                        text = text.replace(' ', '\n', 1)
                    table = advances(font, bold, italic)
                    assert table is not None
                    em, units = table
                    pixels = sorted(chooser.sample(range(max(int(sum(margins) / .75) + 4, 24), 500), 12))
                    cases.append({'text': text, 'font': font, 'size': size, 'bold': bold, 'italic': italic, 'kind': kind, 'margins': list(margins), 'wrap': True, 'pixels': pixels,
                                  'lines': [line_spans(text, pixel * .75 - sum(margins), units, em, size) for pixel in pixels]})
    for font in FONTS:
        text = 'Wexford quarterly ledger\nzig-zag rate-card'
        table = advances(font, False, False)
        assert table is not None
        pixels = [40, 80, 160, 320]
        cases.append({'text': text, 'font': font, 'size': 11.0, 'bold': False, 'italic': False, 'kind': 'rectangle', 'margins': [7.2, 7.2], 'wrap': False, 'pixels': pixels,
                      'lines': [line_spans(text, None, table[1], table[0], 11.0) for _ in pixels]})
    FIXTURE.write_text(json.dumps({'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    print(len(cases), 'cases,', sum(len(case['pixels']) for case in cases), 'widths,', split, 'where the rounding decides')  # type: ignore[arg-type]


if __name__ == '__main__':
    main()
