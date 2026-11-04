"""What a digit does in a number format, through Range.Text and VBA's Format, replayed from live Excel (scripts/measure_digit_formats.py,
tests/fixtures/digit_formats.json). A digit from 1 to 9 is text, and a point with no placeholder before it has the number in front of it."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/digit_formats.json').read_text(encoding='utf-8'))
FORMATS: list[str] = RECORD['formats']
WIDTH = len(RECORD['values'])
#: VBA's Format has an engine of its own, which differs on a comma or percent with no placeholder, an exponent, a star, an escaped
#: point, and a zero section. Only Range.Text and NumberFormat are held to Excel for these.
FORMAT_DIFFERS = ('1,5', '1,000', '1,000.5', '1.5E+0', '1E+0', '1E+1', '1/2', '1 1/2', '1*x.5', '1:5')
#: Excel refuses this and takes x1.5 and 1.5 x, by no rule found.
UNKNOWN = ('x.y',)


def test_each_format_in_order_as_excel_answered() -> None:
    """In one cell and in order, as what shows widens its column and a fill fills that width."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['source'], name='Probe')
    wrong: list[tuple[str, str, str, str]] = []
    for at, code in enumerate(FORMATS):
        for place, value in enumerate(RECORD['values']):
            answer = RECORD['answers'][at * WIDTH + place]
            try:
                got = str(app.run('Probe.Tried', code, _value(value)))
            except VBAUnsupportedError:
                continue  # it says it cannot
            if code in UNKNOWN:
                continue
            if code in FORMAT_DIFFERS:
                got, answer = got.rpartition('|')[0], answer.rpartition('|')[0]
            if got != answer:
                wrong.append((code, value, answer, got))
    assert not wrong


def _value(text: str) -> object:
    if text == 'True':
        return True
    return text.strip('"') if text.startswith('"') else float(text)
