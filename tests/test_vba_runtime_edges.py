"""VBA run-time edges as native Excel runs them: typed-array return types, Collection
keys and positions, Round, negative zero, 0 / 0, Val of hex text, AscW, IsDate,
dotted time text and the Format functions' optional arguments
(scripts/measure_vba_runtime_edges.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, object]] = json.loads((Path(__file__).parent / 'fixtures/vba_runtime_edges.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_runtime_edge(probe: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
