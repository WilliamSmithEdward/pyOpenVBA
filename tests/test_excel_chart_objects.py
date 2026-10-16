"""Embedded chart APIs measured in native Excel."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/excel_chart_objects.json').read_text(encoding='utf-8'))['probes']
GAPS = {'negative_position', 'negative_position_window'}

@pytest.mark.parametrize('probe', [pytest.param(probe, id=probe['name'], marks=[pytest.mark.xfail(
    reason='Window-dependent automatic chart placement remains unmodelled', strict=True)] if probe['name'] in GAPS else []) for probe in PROBES])
def test_native_chart_objects(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']
