"""The rest of Workbook.Connections, Scenarios.CreateSummary and Workbooks' file-naming members, replayed from live Excel
(scripts/measure_connection_members.py, tests/fixtures/connection_members.json): what each answers or raises."""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'connection_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What broke the run, and what the model has no object for: a file's refresh, an ODC file written, the hidden defaults.
LEFT_OUT = ('.SaveAsODC', 'a.Refresh', 'n.Connections.AddFromFile', 'w.CanCheckOut', '[_Default]', 'ListObjects(1).SourceType', 'w.CheckOut path',
            'r = ws.Scenarios.CreateSummary(xlStandardSummary)')
#: What the model reports instead of following.
REPORTED = ('FastCombine = True', 'xlPivotTable', 'CreateSummary: r = Dumped')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT + REPORTED)]
SHOWN = [index for index, name in enumerate(NAMES) if any(text in name for text in REPORTED)]


def _run(index: int, folder: Path) -> str:
    copy = folder / 'book.xlsx'
    shutil.copy(FIXTURES / RECORD['file'], copy)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='MemberProbe')
    return str(app.run('MemberProbe.Tried', index, str(copy)))


def _masked(text: str) -> str:
    """Who made a scenario, and when, and where this machine's documents are."""
    text = re.sub(r'Created by .*? on \d+/\d+/\d{4}', 'Created', text)
    return re.sub(r"find [^ ]*[\\/]Documents[\\/]", 'find <documents>\\\\', text)


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    assert _masked(_run(index, tmp_path)) == _masked(RECORD['cases'][index][1])


@pytest.mark.parametrize('index', SHOWN, ids=[NAMES[index] for index in SHOWN])
def test_what_is_not_followed_reports_itself(tmp_path: Path, index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index, tmp_path)


def test_workbooks_close_ended_the_macro_in_excel() -> None:
    """Measured, and still reported by the model: Excel's macro stops with its own workbook, before the next line."""
    assert RECORD['workbooks_close']['log'] == 'before 3\n'
