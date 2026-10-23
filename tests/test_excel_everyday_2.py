"""A second sweep of everyday macros, replayed against what live Excel answered (scripts/measure_excel_everyday_2.py).

Each probe is a module of its own on a sheet it adds and deletes; the
model's answer, or the run-time error it stops with, is Excel's.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, Any]] = json.loads(
    (Path(__file__).parent / 'fixtures/excel_everyday_2.json').read_text(encoding='utf-8'))['probes']


def _answer(macro: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(macro, name='Harness')
    try:
        return app.run('Harness.Probe')
    except VBARuntimeError as exc:
        return {'error': exc.number}


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_an_everyday_macro(probe: dict[str, Any]) -> None:
    expected = {'error': probe['error']} if 'error' in probe else probe['result']
    assert _answer(probe['macro']) == expected
