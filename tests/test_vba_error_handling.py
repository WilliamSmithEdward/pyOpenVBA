"""Error handling as native Excel runs it: errors passing up, Resume, handlers in loops,
Err.Raise and Err's properties (scripts/measure_vba_error_handling.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = json.loads((Path(__file__).parent / 'fixtures/vba_error_handling.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, object]] = FIXTURE['probes']
CLASSES: dict[str, str] = FIXTURE['classes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_error_handling(probe: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    for name, source in CLASSES.items():
        app.add_module(source, name=name, kind='class')
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
