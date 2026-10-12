"""The date functions over a matrix as native Excel runs them: DateAdd, DateDiff, DatePart,
week numbers, Weekday, DateSerial, TimeSerial and the parts (scripts/measure_vba_date_functions.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, object]] = json.loads((Path(__file__).parent / 'fixtures/vba_date_functions.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_date_function(probe: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
