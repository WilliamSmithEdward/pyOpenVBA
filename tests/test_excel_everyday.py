"""Everyday Excel macros, measured by scripts/measure_excel_everyday.py.

Seventy-seven of what macros commonly do -- find the last row, copy and
paste values, sort, filter, look up, format, add sheets and names, read and
write arrays, links, notes, shapes, charts, tables, validation, protection --
each replayed through the model, which has to give what Excel gave.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/excel_everyday.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_an_everyday_macro(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
