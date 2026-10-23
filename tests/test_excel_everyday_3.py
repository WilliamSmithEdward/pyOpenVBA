"""A third sweep of everyday macros, replayed against what live Excel answered (scripts/measure_excel_everyday_3.py).

Each probe is a module of its own on a sheet it adds and deletes; the
model's answer, or the run-time error it stops with, is Excel's. A probe
in GAPS needs what the model does not have yet, and reports it.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/excel_everyday_3.json').read_text(encoding='utf-8'))['probes']
#: What each probe the model cannot answer yet needs; none now.
GAPS: dict[str, str] = {}


def _answer(macro: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(macro, name='Harness')
    try:
        return app.run('Harness.Probe')
    except VBARuntimeError as exc:
        return {'error': exc.number}


@pytest.mark.parametrize('probe', [
    pytest.param(probe, marks=pytest.mark.xfail(raises=VBAUnsupportedError, strict=True, reason=GAPS[probe['name']]))
    if probe['name'] in GAPS else probe for probe in PROBES], ids=lambda probe: probe['name'])
def test_an_everyday_macro(probe: dict[str, Any]) -> None:
    expected = {'error': probe['error']} if 'error' in probe else probe['result']
    assert _answer(probe['macro']) == expected


def test_each_gap_is_a_probe() -> None:
    assert set(GAPS) <= {probe['name'] for probe in PROBES}
