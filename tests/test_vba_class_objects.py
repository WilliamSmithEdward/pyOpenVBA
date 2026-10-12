"""Class modules and objects as native VBA has them: properties, events, CallByName,
object arrays, Err across a member call, Let into an object variable and Is
(scripts/measure_vba_class_objects.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

FIXTURE = json.loads((Path(__file__).parent / 'fixtures/vba_class_objects.json').read_text(encoding='utf-8'))
PROBES: list[dict[str, object]] = FIXTURE['probes']
CLASSES: dict[str, str] = FIXTURE['classes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_class_object(probe: dict[str, object]) -> None:
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
