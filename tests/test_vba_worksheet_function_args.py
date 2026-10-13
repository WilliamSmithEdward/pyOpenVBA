"""What worksheet functions called from VBA make of each kind of argument, as native
Excel does (scripts/measure_vba_worksheet_function_args.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = json.loads((Path(__file__).parent / 'fixtures/vba_worksheet_function_args.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, object]] = FIXTURE['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_worksheet_function_args(probe: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
