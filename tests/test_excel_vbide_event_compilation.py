"""Native event edits and already compiled source replacement."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, str]] = []
for filename in ('vbide_event_edits.json', 'vbide_compiled_cache.json'):
    PROBES.extend(json.loads((Path(__file__).parent / 'fixtures' / filename).read_text(encoding='utf-8'))['probes'])

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_edited_source_execution(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']
