"""Everyday array handling as native Excel runs it: unallocated arrays and (Not a),
ReDim and Preserve edges, Erase, array parameters, nested arrays, ParamArray, Split,
Join and Filter, and what a typed dynamic array takes (scripts/measure_vba_arrays.py)."""
from __future__ import annotations
import json
from pathlib import Path
import pytest
from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

PROBES: list[dict[str, object]] = json.loads((Path(__file__).parent / 'fixtures/vba_arrays.json').read_text(encoding='utf-8'))['probes']


@pytest.mark.parametrize('probe', PROBES, ids=[str(probe['name']) for probe in PROBES])
def test_native_arrays(probe: dict[str, object]) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(str(probe['macro']), name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
