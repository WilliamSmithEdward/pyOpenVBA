"""Class modules in everyday VBA, replayed against what live Excel answered (scripts/measure_vba_classes.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError, VBARuntimeError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_classes.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, Any]] = RECORD['probes']


def _answer(macro: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['logger'], name='Logger')
    for name, source in RECORD['classes'].items():
        app.add_module(source, name=name, kind='class')
    try:
        app.add_module(macro, name='Harness')
        return app.run('Harness.Probe')
    except VBACompileError:
        return 'does not compile'
    except VBARuntimeError as failure:
        return {'error': failure.number}


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_as_excel_answered(probe: dict[str, Any]) -> None:
    expected = 'does not compile' if 'compile' in probe else {'error': probe['error']} if 'error' in probe else probe['result']
    assert _answer(probe['macro']) == expected
