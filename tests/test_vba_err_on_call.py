"""What a call does to Err, replayed from live Excel (scripts/measure_err_on_call.py, tests/fixtures/err_on_call.json):
it stays through the call, however the call is written, until the called procedure runs an On Error statement."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/err_on_call.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    application = ExcelApplication()
    application.add_workbook()
    application.add_module(RECORD['source'], name='ErrProbe')
    return application


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[' '.join(case[:3]) for case in RECORD['cases']])
def test_each_call_as_excel_answered(app: ExcelApplication, index: int) -> None:
    assert app.run('ErrProbe.Probe', index, index) == RECORD['cases'][index][3]
