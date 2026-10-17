"""Native cache and event behavior when attaching WithEvents listeners."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/sink_calculation.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_listener_attachment_retains_calculation_state(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(RECORD['helpers'], name='EventLog')
    app.add_module(RECORD['sink'], name='ExcelSink', kind='class')
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == case['result']
