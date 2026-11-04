"""Which palette slot ColorIndex reads for a colour outside the palette, replayed from live Excel: every grey and 3,000 random colours
(scripts/measure_color_index_nearest.py, tests/fixtures/color_index_nearest.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel._styles import Colors

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/color_index_nearest.json').read_text(encoding='utf-8'))


def test_each_colour_reads_the_slot_excel_read() -> None:
    colors = Colors()
    assert [(colour, colors.nearest(colour)) for colour, _ in RECORD['pairs']] == [tuple(pair) for pair in RECORD['pairs']]
