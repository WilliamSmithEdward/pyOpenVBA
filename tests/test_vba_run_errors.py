"""An error inside a macro Application.Run started goes past every handler of its caller, where Excel shows its run-time error
dialog, and errors Run raises itself stay the caller's to trap, over 18 live cases."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/vba_run_errors.json').read_text(encoding='utf-8'))


def _answer(source: str) -> str:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    try:
        return str(app.run('Harness.Probe'))
    except VBARuntimeError as caught:
        # Past every handler is where Excel shows its dialog; any other error reached the host the ordinary way.
        return f"{'dialog' if caught.past_handlers else 'error'} {caught.number}"


@pytest.mark.parametrize(('case', 'excel'), list(zip(RECORD['cases'], RECORD['excel'])), ids=[case['name'] for case in RECORD['cases']])
def test_run_error_goes_where_excel_sends_it(case: dict[str, Any], excel: str) -> None:
    assert _answer(case['source']) == excel
