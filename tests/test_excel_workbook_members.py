"""Members of Workbook, Sheets, Worksheet and PageSetup a macro meets with no screen, server or mail system, replayed from
live Excel (scripts/measure_workbook_members.py, tests/fixtures/workbook_members.json): what each answers or raises."""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/workbook_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What the model reports instead of following: a theme's font changed, a custom XML part saved, read and write access
#: again where Excel reopens the file.
REPORTED = ('MinorFont(1).Name = ', 'MajorFont(2).Name = ', 'saved.CustomXMLParts.Add', 'saved.ChangeFileAccess xlReadOnly: saved.ChangeFileAccess xlReadWrite',
            'CustomXMLParts(1): r = Len', 'A1").Value = 1: saved.ChangeFileAccess xlReadOnly: saved.ChangeFileAccess')
#: Left out: what broke the run; members with no model yet; a pivot table made in the probe, which Excel refused; the
#: nodes of a custom XML part and the window a macro opens, which are other work; and what this machine answers.
LEFT_OUT = ('wb.NewWindow', 'ToggleFormsDesign', 'MailEnvelope', 'SelectSingleNode', 'DocumentElement', 'wb.[_CodeName]', 'wb.Theme.ThemeColorScheme.Colors(5) = 255',
            'CreatePivotTable', 'HPageBreaks.Add', 'CanCheckOut', '[_NewEnum]', 'Colors(5).Parent', 'Colors(5).RGB = -1', 'Colors(5).RGB = 16777216', 'r = wb.CustomXMLParts.Add().Count', 'r = ws.Scenarios.Merge(other)', 'ws.Scenarios.CreateSummary',
            'ws.PageSetup.Pages.Parent', 'wb.UpdateFromFile', 'saved.UpdateFromFile', 'ws.PageSetup.LeftHeaderPicture.Filename = ')
KEPT = [index for index, name in enumerate(NAMES)
        if not any(text in name for text in LEFT_OUT) and not any(text in name for text in REPORTED)]
SHOWN = [index for index, name in enumerate(NAMES) if any(text in name for text in REPORTED) and not any(text in name for text in LEFT_OUT)]


def _run(index: int, folder: Path) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='MemberProbe')
    return str(app.run('MemberProbe.Tried', index, str(folder)))


def _dated(text: str) -> str:
    return re.sub(r'\d+/\d+/\d{4}', '<date>', text)


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(tmp_path: Path, index: int) -> None:
    assert _dated(_run(index, tmp_path)) == _dated(RECORD['cases'][index][1])


@pytest.mark.skipif(sys.platform != 'win32', reason='the macros name files by Windows paths')
@pytest.mark.parametrize('index', SHOWN, ids=[NAMES[index] for index in SHOWN])
def test_what_is_not_followed_reports_itself(tmp_path: Path, index: int) -> None:
    with pytest.raises(VBAUnsupportedError):
        _run(index, tmp_path)
