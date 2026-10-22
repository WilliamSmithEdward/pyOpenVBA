"""Which sheet names a formula writes in apostrophes, measured by scripts/measure_sheet_quoting.py.

A name is bare when it is a letter or underscore, then letters, digits,
underscores and full stops, in any script: x.y, Donnees with its accent, a
CJK name. Excel still quotes such a name when a formula would read it as
something else: a cell (Q4, A01, XFD1048576, though not LOG10, a function's
name), a row or column in R1C1 (R, C, RC, R2, C3 and any name starting with
one and going on with more than a full stop, R1x or C1A), or a logical.
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
