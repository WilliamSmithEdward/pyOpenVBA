"""Which sheet names a formula writes in apostrophes, measured by scripts/measure_sheet_quoting.py.

A name is bare when it is an ASCII letter or underscore, then letters,
digits, underscores and full stops, and past ASCII any character but the
few pyOfficeEditor measured Excel quoting for (a no-break space or a
copyright sign anywhere, a euro sign first): x.y, an accented name, a CJK
name, x-y with an en dash. Excel still quotes such a name when a formula
would read it as something else: a cell (Q4, A01, XFD1048576, though no
name starting LOG10, a function's name), R, C or RC with a number up to
1048576 and anything after it but a full stop (R2, C3, RC1, R1x, C1A; R1.x
is bare), or a logical.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/sheet_quoting.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_a_sheet_named_in_a_formula(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']
