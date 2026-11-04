"""When a closing workbook's links take its cells for the last time, replayed from live Excel
(scripts/measure_close_links.py, tests/fixtures/close_links.json): after its window events, so an edit a handler makes in
any of them reaches the workbooks that read it."""
from __future__ import annotations

import base64
import json
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/close_links.json').read_text(encoding='utf-8'))


def _harness_free(answer: str) -> str:
    """The value and the events of the two books in the case, without those of the workbook the macro ran from."""
    value, *events = answer.split('|')
    return '|'.join([value, *(event for event in events if event.split()[1] in ('Alpha.xlsx', 'Dest.xlsx'))])


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('mode', list(RECORD['answers']))
def test_each_handler_as_excel_answered(tmp_path: Path, mode: str) -> None:
    for name, data in RECORD['files'].items():
        (tmp_path / name).write_bytes(base64.b64decode(data))
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['sink'], name='CloseSink', kind='class')
    app.add_module(RECORD['source'], name='CloseProbe')
    assert _harness_free(str(app.run('CloseProbe.Tried', mode, str(tmp_path)))) == _harness_free(RECORD['answers'][mode])
