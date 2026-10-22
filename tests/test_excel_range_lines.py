"""Range.Rows and Range.Columns, measured by scripts/measure_range_lines.py.

The range Rows or Columns answers is the same cells taken a row or a column
at a time: its Count is its first area's rows or columns, an index is a row
or a column counted from its first, For Each walks every area's, and Offset,
Resize and its own rows or columns stay that way, while Cells, Areas, Union,
Intersect and the selection are cells again. Every other member acts on its
cells, AutoFit and Hidden on its rows or columns.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/range_lines.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
#: Probes the model does not answer as Excel does, and why.
GAPS: dict[str, str] = {}


@pytest.mark.parametrize('name', [
    pytest.param(name, marks=[pytest.mark.xfail(reason=GAPS[name], strict=True)] if name in GAPS else [])
    for name in PROBES])
def test_rows_and_columns_of_a_range(name: str) -> None:
    probe = PROBES[name]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(probe['macro'], name='Harness')
    if 'error' in probe:
        with pytest.raises(VBARuntimeError) as raised:
            app.run('Harness.Probe')
        assert raised.value.number == probe['error']
    else:
        assert app.run('Harness.Probe') == probe['result']
