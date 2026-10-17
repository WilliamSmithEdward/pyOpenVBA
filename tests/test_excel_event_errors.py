"""Native event-local recovery, unhandled callback boundaries and Err state."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/event_errors.json').read_text(encoding='utf-8'))


@pytest.mark.parametrize('case', RECORD['probes'], ids=lambda case: case['name'])
def test_native_error_context_and_event_boundaries(case: dict[str, Any]) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    app.add_module(RECORD['helpers'], name='EventLog')
    for module in case.get('modules', []):
        app.add_module(module['source'], name=module['name'], kind=module['kind'])
    if 'sink' in case:
        app.add_module(case['sink'], name='ExcelSink', kind='class')
        app.add_module(case['sheet_handler'], name=book.sheets_[0].code_name, kind='document')
        app.add_module(case['book_handler'], name=book.code_name, kind='document')
    app.add_module(case['source'], name='Harness')
    app.add_module(RECORD['read'], name='Readback')
    if case['unhandled']:
        with pytest.raises(VBARuntimeError) as failure:
            app.run('Harness.Probe')
        assert failure.value.number == 9
        assert failure.value.source == 'event-source'
        assert failure.value.description == 'event-error'
    else:
        app.run('Harness.Probe')
    assert app.run('Readback.ReadState') == case['result']
