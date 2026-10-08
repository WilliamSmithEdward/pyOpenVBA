"""Native typed-record array arguments and standard-module Variant barriers."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBACompileError

PROBES: list[dict[str, str]] = []
for fixture in ('record_array_arguments', 'record_variant_calls'):
    PROBES.extend(json.loads((Path(__file__).parent / f'fixtures/{fixture}.json').read_text(encoding='utf-8'))['probes'])

@pytest.mark.parametrize('probe', PROBES, ids=[probe['name'] for probe in PROBES])
def test_native_record_array_coercion(probe: dict[str, str]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'compile_error' in probe:
        with pytest.raises(VBACompileError, match=probe['compile_error'].split('\n')[-1]):
            app.run('Harness.Probe')
    else:
        assert app.run('Harness.Probe') == probe['result']
