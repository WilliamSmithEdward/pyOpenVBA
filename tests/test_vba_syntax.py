"""Everyday VBA syntax, replayed against what live Excel answered (scripts/measure_vba_syntax.py): its answer, the run-time error it
stops with, or that it does not compile."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError, VBAUnsupportedError

PROBES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/vba_syntax.json').read_text(encoding='utf-8'))['probes']
#: Each hands out a memory address, which the engine has none of.
REFUSED = {'addressof', 'varptr'}


def _answer(macro: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(macro, name='Harness')
        return app.run('Harness.Probe')
    except VBACompileError:
        return 'does not compile'
    except VBARuntimeError as failure:
        return {'error': failure.number}


@pytest.mark.parametrize('probe', [probe for probe in PROBES if probe['name'] not in REFUSED], ids=lambda probe: probe['name'])
def test_as_excel_answered(probe: dict[str, Any]) -> None:
    expected = 'does not compile' if 'compile' in probe else {'error': probe['error']} if 'error' in probe else probe['result']
    assert _answer(probe['macro']) == expected


@pytest.mark.parametrize('probe', [probe for probe in PROBES if probe['name'] in REFUSED], ids=lambda probe: probe['name'])
def test_what_hands_out_an_address_says_it_is_not_implemented(probe: dict[str, Any]) -> None:
    with pytest.raises(VBAUnsupportedError):
        _answer(probe['macro'])
