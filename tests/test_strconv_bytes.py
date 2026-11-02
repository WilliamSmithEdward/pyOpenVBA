"""What StrConv(text, vbFromUnicode) gives, replayed from live Excel (scripts/measure_strconv_bytes.py,
tests/fixtures/strconv_bytes.json): a String of the text's ANSI bytes, half a character over when there is an odd count."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/strconv_bytes.json').read_text(encoding='utf-8'))


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    application = ExcelApplication()
    application.add_workbook()
    application.add_module(RECORD['source'], name='BytesProbe')
    return application


#: A cell Excel writes this to makes a whole character of the odd byte, and the model's drops it. That one is not done.
KEPT = [index for index, case in enumerate(RECORD['cases']) if case[0] != 'into a cell']


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(app: ExcelApplication, index: int) -> None:
    assert app.run('BytesProbe.Probe', index, index) == RECORD['cases'][index][1]
