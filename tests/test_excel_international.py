"""Application.International and the separator settings, measured by scripts/measure_international.py.

The model answers for the en-US locale the measuring machine runs: every
index from 1 to 46, the culture's name at 46, error 1004 outside them, and
the decimal and thousands separators with UseSystemSeparators on. Setting
a separator to another one is refused as unsupported: it would change every
number shown, which the model does not follow.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

PROBE: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/international.json').read_text(encoding='utf-8'))['probes'][0]


def test_every_international_index_and_the_separators() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(PROBE['macro'], name='Harness')
    assert app.run('Harness.Probe') == PROBE['result']


@pytest.mark.parametrize(('member', 'value', 'refused'), [
    ('DecimalSeparator', '"."', False), ('DecimalSeparator', '","', True),
    ('ThousandsSeparator', '","', False), ('ThousandsSeparator', '" "', True),
    ('UseSystemSeparators', 'False', False),
])
def test_setting_a_separator(member: str, value: str, refused: bool) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Sub Put()\n    Application.{member} = {value}\nEnd Sub\n', name='Harness')
    if refused:
        with pytest.raises(VBAUnsupportedError):
            app.run('Harness.Put')
    else:
        app.run('Harness.Put')
