"""Bake src/pyopenvba/apps/excel/_name_characters.py from tests/fixtures/name_characters_bmp.json, the UTF-16 units from U+0080 up
that live Excel refuses in a defined name, first and after the first (scripts/measure_name_characters_bmp.py).

    python scripts/bake_name_characters.py
"""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'tests' / 'fixtures' / 'name_characters_bmp.json'
TARGET = ROOT / 'src' / 'pyopenvba' / 'apps' / 'excel' / '_name_characters.py'
HEADER = '''"""The UTF-16 units from U+0080 up that Excel refuses in a defined name, as ranges of code units, first and after the first.

Baked by scripts/bake_name_characters.py from tests/fixtures/name_characters_bmp.json, where live Excel was asked every unit
in both places; do not edit by hand. Excel reads a name unit by unit, so a character beyond the Basic Multilingual Plane is taken
wherever its high and low surrogates are (tests/fixtures/name_surrogate_pairs.json).
"""
'''


def spans(points: list[int]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for point in sorted(set(points)):
        if out and out[-1][1] == point - 1:
            out[-1] = (out[-1][0], point)
        else:
            out.append((point, point))
    return out


def _table(name: str, ranges: list[tuple[int, int]]) -> str:
    cells = [f'(0x{lo:04X}, 0x{hi:04X})' for lo, hi in ranges]
    rows = [', '.join(cells[at:at + 6]) for at in range(0, len(cells), 6)]
    return f'{name}: tuple[tuple[int, int], ...] = (\n' + ''.join(f'    {row},\n' for row in rows) + ')\n'


def main() -> None:
    record = json.loads(SOURCE.read_text(encoding='utf-8'))
    text = HEADER + '\n' + _table('REFUSED_FIRST', spans(record['first'])) + '\n' + _table('REFUSED_LATER', spans(record['later']))
    TARGET.write_text(text, encoding='utf-8', newline='\n')
    print('baked', TARGET, len(spans(record['first'])), len(spans(record['later'])))


if __name__ == '__main__':
    main()
