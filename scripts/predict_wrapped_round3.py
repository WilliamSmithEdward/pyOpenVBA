"""The third held-out round for wrapped rows, written and committed before Excel is asked: hyphens in columns narrower than a letter, line feeds, italic and indents.

Italic is predicted to take a pixel for each of its steps, a fifth of the font's ascent, and an indent level 9 pixels, as they do from a cell's room for its text in `_column_fit.narrowed`. Those two are carried over from that rule, not fitted here.

    python scripts/predict_wrapped_round3.py

Writes tests/fixtures/wrapped_round3_predictions.json.
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
from pyopenvba.apps.excel._font_widths import ascent, pixel_height, widths  # noqa: E402

OUT = ROOT / 'tests/fixtures/wrapped_round3_predictions.json'
#: font, size, bold, italic, indent
SETTINGS = (('Calibri', 11, False, False, 0), ('Aptos Narrow', 11, False, False, 0), ('Calibri', 18, False, False, 0), ('Calibri', 11, False, True, 0), ('Aptos Narrow', 14, False, True, 0), ('Calibri', 11, True, True, 0), ('Calibri', 11, False, False, 1), ('Calibri', 11, False, False, 3), ('Aptos Narrow', 12, True, False, 2))
TEXTS = ('1-2-3-4', 'q-r-s', 'a-b-c d-e', 'ab-c-de', 'p-', '-q-', 'r--s--t', 'first\nsecond line here', 'one\n\nthree', 'tail\n', '\nhead', 'wrap me around please\nand again please now', 'plain words for the italic and indent rows', 'unbrokenwordforcharacterbreaks')
COLUMNS = tuple(range(2, 40, 3)) + tuple(range(40, 300, 13))


def main() -> None:
    cases = []
    for font, size, bold, italic, indent in SETTINGS:
        table = widths(font, bold, italic, float(size))
        rise = ascent(font, bold, italic, float(size))
        pixels = pixel_height(float(size))
        family = table_of(font, bold, italic)
        assert table is not None and rise is not None and pixels is not None and family is not None
        line = row_of(family, pixels, alone=True)[0] * 0.75
        taken = (rise // 5 if italic else 0) + 9 * indent

        def extent(text: str) -> int:
            return sum(table[character] for character in text)

        for text in TEXTS:
            cases.append({'font': font, 'size': size, 'bold': bold, 'italic': italic, 'indent': indent, 'text': text, 'line': line, 'taken': taken, 'digit': table['0'],
                          'predicted': {str(column): height_of(lines_of(text, column, extent, table['0'], taken), line) for column in range(1, 320)}})
    OUT.write_text(json.dumps({'columns': COLUMNS, 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'curves')


if __name__ == '__main__':
    main()
