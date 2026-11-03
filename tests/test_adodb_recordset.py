"""A disconnected ADODB.Recordset and Range.CopyFromRecordset, replayed against what live Excel answered
(scripts/measure_adodb_recordset.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/adodb_recordset.json').read_text(encoding='utf-8'))
#: adDBTime reads as today's date, and a recordset copied over a table takes the table away: neither is modelled.
GAPS = ('"F", 134,', 'ws.ListObjects.Add')
KEPT = [index for index, (case, _) in enumerate(RECORD['cases']) if not any(text in case for text in GAPS)]
LEFT = [index for index in range(len(RECORD['cases'])) if index not in KEPT]


def _made() -> ExcelApplication:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module(RECORD['source'], name='RecordsetProbe')
    return made


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    return _made()


@pytest.mark.parametrize('index', KEPT)
def test_as_excel_answered(app: ExcelApplication, index: int) -> None:
    assert str(app.run('RecordsetProbe.Probe', index)) == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', LEFT)
def test_what_is_left_says_so(index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _made().run('RecordsetProbe.Probe', index)


def test_a_recordset_on_a_source_is_refused() -> None:
    made = ExcelApplication()
    made.add_workbook()
    made.add_module('Function P()\nDim r As Object\nSet r = CreateObject("ADODB.Recordset")\nr.Open "select 1"\nEnd Function', name='M')
    with pytest.raises(VBAUnsupportedError):
        made.run('M.P')
