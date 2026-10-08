"""Native module-qualified VBA record declarations and argument checking."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

PROBES: list[dict[str, str]] = json.loads(
    (Path(__file__).parent / 'fixtures/record_qualification.json').read_text(encoding='utf-8')
)['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_qualification(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'compile_error' in probe:
        with pytest.raises(VBACompileError, match='Variable required'):
            app.run('Harness.Probe')
    else:
        assert app.run('Harness.Probe') == probe['result']
