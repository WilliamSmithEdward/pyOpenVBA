"""Native EnableEvents boundaries between sheet, workbook and application events."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/event_dispatch.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_event_dispatch_boundaries(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(RECORD['helpers'], name='EventLog')
    app.add_module(case['sink'], name='ExcelSink', kind='class')
    app.add_module(case['sheet_handler'], name=book.sheets_[0].code_name, kind='document')
    app.add_module(case['book_handler'], name=book.code_name, kind='document')
    app.add_module(case['source'], name='Harness')
    assert app.run('Harness.Probe') == case['result']
