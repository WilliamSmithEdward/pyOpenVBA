"""Predictions for measure_wrapped_curves.py, written and committed before Excel is asked: how tall fresh wrapped text makes a row across column widths, in fonts the wrap was not fitted on, by scripts/wrapped_lines.py and the model's tables.

    python scripts/predict_wrapped_curves.py

Writes tests/fixtures/wrapped_curves_predictions.json.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from wrapped_lines import height_of, lines_of  # noqa: E402
from pyopenvba.apps.excel._font_rows import row_of, table_of  # noqa: E402
from pyopenvba.apps.excel._font_widths import pixel_height, widths  # noqa: E402

OUT = ROOT / 'tests/fixtures/wrapped_curves_predictions.json'
FONTS = (('Calibri', 8, False), ('Calibri', 10, False), ('Calibri', 12, False), ('Calibri', 14, False), ('Calibri', 20, False), ('Aptos Narrow', 9, False), ('Aptos Narrow', 14, False), ('Aptos Narrow', 11, True), ('Calibri', 11, True))
TEXTS = ('well-known fact', 'x-y-z', 'a--b', '-ab', 'ab-', 'ab- cd', 'ab -cd', 'pre-war post-war', 'three   spaced   out', 'Sentence one. Sentence two! And three?', 'under_score/slash.dot,comma', 'zigzag zigzag zigzag zigzag', 'MixedCaseWordWithoutAnyBreaks',
         '12345 67890 12345', 'Tab less text with several little words in it', '(brackets) [square] {curly}', 'tilde~and^caret', 'semi;colon:colon')
COLUMNS = (8, 12, 16, 20, 24, 28, 32, 40, 48, 56, 64, 72, 80, 96, 112, 128, 150, 180, 220, 280, 340)


def main() -> None:
    cases = []
    for font, size, bold in FONTS:
        table = widths(font, bold, False, float(size))
        assert table is not None, (font, size)
        pixels = pixel_height(float(size))
        assert pixels is not None
        family = table_of(font, bold, False)
        assert family is not None
        line = row_of(family, pixels, alone=True)[0] * 0.75

        def extent(text: str) -> int:
            return sum(table[character] for character in text)

        for text in TEXTS:
            cases.append({'font': font, 'size': size, 'bold': bold, 'text': text, 'line': line, 'predicted': {str(column): height_of(lines_of(text, column, extent, table['0']), line) for column in COLUMNS}})
    OUT.write_text(json.dumps({'columns': COLUMNS, 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'curves', sorted({(one['font'], one['size'], one['bold'], one['line']) for one in cases}))


if __name__ == '__main__':
    main()
