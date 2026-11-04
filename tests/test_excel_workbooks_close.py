"""Workbooks.Close, replayed from live Excel (scripts/measure_workbooks_close.py, tests/fixtures/workbooks_close.json)."""
from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/workbooks_close.json').read_text(encoding='utf-8'))
#: Excel asks about the changed book, which the driver cancelled: the model reports the prompt. Hiding a book's window reports itself.
REPORTED = ('changed prompt', 'hidden')


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('name', [name for name in RECORD['cases'] if name not in REPORTED])
def test_each_case_as_excel_answered(tmp_path: Path, name: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['sink'], name='CloseSink', kind='class')
    app.add_module(RECORD['source'], name='CloseProbe')
    answer = str(app.run('CloseProbe.Tried', list(RECORD['cases']).index(name), str(tmp_path)))
    wanted = RECORD['answers'][name]
    if name == 'argument':  # the wording of 450 is the interpreter's
        answer, wanted = answer.split('|', 1)[1], wanted.split('|', 1)[1]
    assert answer.replace(str(tmp_path), '<DIR>') == wanted


def test_a_changed_book_with_alerts_on_reports_the_prompt_and_closes_nothing() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Tried()\n    ThisWorkbook.IsAddin = True\n    Workbooks.Add\n    Workbooks(1).Saved = False\n    Application.DisplayAlerts = True\n    Workbooks.Close\nEnd Sub\n', name='M')
    with pytest.raises(VBAUnsupportedError, match='save prompt'):
        app.run('M.Tried')
    assert len(app.application.workbooks_.books) == 1


def test_a_macro_in_a_closing_book_ends_there() -> None:
    """The last case of the fixture: the line after Workbooks.Close never ran."""
    assert RECORD['answers']['own workbook'] == 'runner-error: before 2;'
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Tried()\n    Application.DisplayAlerts = False\n    Workbooks.Add\n    Debug.Print "before " & Workbooks.Count\n    Workbooks.Close\n    Debug.Print "after"\nEnd Sub\n', name='M')
    app.run('M.Tried')
    assert app.console == ['before 2']
    assert not app.application.workbooks_.books
