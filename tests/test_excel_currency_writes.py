"""Native Currency assignment precision through Range.Value and Value2."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/currency_writes.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_currency_write(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(case['source'], name='Harness')
    assert app.run('Probe') == case['result']
    app.add_module('Function ReadSaved() As String\nReadSaved=CStr(Range("A1").Value2) & "|" & Range("A1").NumberFormat\nEnd Function',
                   name='Inspection')
    path = app.save(tmp_path / 'currency.xlsm')
    assert ExcelApplication.open(path).run('ReadSaved') == case['result']
