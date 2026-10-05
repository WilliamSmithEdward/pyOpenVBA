"""ByRef storage aliases measured in native Excel VBA."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
PROBES: list[dict[str, str | int]] = []
for filename in ('instance_array_byref.json', 'record_byref.json'):
    PROBES.extend(json.loads((Path(__file__).parent / 'fixtures' / filename).read_text(encoding='utf-8'))['probes'])

@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_byref_storage(probe: dict[str, str | int]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(probe['macro']), name='Harness')
    assert app.run('Harness.Probe') == probe['result']
