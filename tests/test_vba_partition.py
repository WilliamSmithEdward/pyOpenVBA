"""QBColor and Partition, replayed against what live Excel answered (scripts/measure_vba_partition.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, Any]] = json.loads((Path(__file__).parent / 'fixtures/vba_partition.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=lambda probe: probe['name'])
def test_as_excel_answered(probe: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    try:
        got: object = app.run('Harness.Probe')
    except VBARuntimeError as failure:
        got = {'error': failure.number}
    assert got == ({'error': probe['error']} if 'error' in probe else probe['result'])
