"""Score scripts/wrapped_lines.py against tests/fixtures/wrapped_curves.json at the column widths Excel gave, which are not always those asked for.

    python scripts/score_wrapped_curves.py
"""
from __future__ import annotations

import collections
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from wrapped_lines import height_of, lines_of  # noqa: E402
from pyopenvba.apps.excel._font_widths import widths  # noqa: E402

FIXTURES = ROOT / 'tests/fixtures'


def main() -> None:
    predictions = json.loads((FIXTURES / 'wrapped_curves_predictions.json').read_text(encoding='utf-8'))['cases']
    curves = json.loads((FIXTURES / 'wrapped_curves.json').read_text(encoding='utf-8'))['curves']
    good = bad = 0
    missed: collections.Counter[str] = collections.Counter()
    samples: dict[str, tuple[int, float, float]] = {}
    by_font: collections.Counter[tuple[str, int, bool, bool]] = collections.Counter()
    for case, curve in zip(predictions, curves, strict=True):
        table = widths(case['font'], case['bold'], False, float(case['size']))
        assert table is not None

        def extent(text: str) -> int:
            return sum(table[character] for character in text)

        for column, tall in curve.items():
            got = height_of(lines_of(case['text'], int(column), extent, table['0']), case['line'])
            right = abs(got - tall) < 1e-6
            good += right
            bad += not right
            by_font[(case['font'], case['size'], case['bold'], right)] += 1
            if not right:
                key = f"{case['font']} {case['size']} {'bold ' if case['bold'] else ''}{case['text']}"
                missed[key] += 1
                samples.setdefault(key, (int(column), tall, got))
    print('good', good, 'off', bad)
    for key in sorted(by_font):
        print(key, by_font[key])
    for key, count in missed.most_common(40):
        print(count, key, samples[key])


if __name__ == '__main__':
    main()
