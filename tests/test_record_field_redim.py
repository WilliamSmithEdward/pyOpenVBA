"""ReDim of a record's dynamic array field, as native VBA resizes it."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication

PROBES: list[dict[str, str]] = json.loads((Path(__file__).parent / 'fixtures/record_field_redim.json').read_text(encoding='utf-8'))['probes']

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_field_redim(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    assert app.run('Harness.Probe') == probe['result']
