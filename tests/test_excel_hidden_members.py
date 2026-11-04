"""The hidden members of Excel's collections and small members of objects otherwise in the model, replayed from live Excel
(scripts/measure_hidden_members.py, tests/fixtures/hidden_members.json): what each answers or raises."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/hidden_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Members still to come, each of which reports itself when asked for: a table style of a workbook's own, a table style's
#: elements, ranges that allow edits, a list column's XPath, Characters of a displayed format, a header's picture object.
LEFT_OUT = ('PhoneticCharacters', 'DisplayDataTypeIcons', 'TableStyleElements', '.Duplicate', 'TableStyles.Add', 'AllowEditRanges', '.XPath', 'DisplayFormat.Characters', 'r = Application.CellFormat',
            'FindFormat', 'OddPage', 'FormatConditions.Add(xlExpression, , "=TRUE").NumberFormat', 'ws.PageSetup.LeftHeader', 'ic.IconCriteria(1).Value', 'r = ActiveWindow.Panes.[_Default](1).Index',
            # Breaks a chart's extent brings into the sheet's used area; a group's items; Errors, which For Each refuses
            # in Excel and walks here; a window's panes, which report themselves.
            'o = ws.HPageBreaks: For', 'o = ws.VPageBreaks: For', 'SparklineGroups(1): For', 'Errors: For', 'ActiveWindow.Panes: For')


def _run(index: int) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='HiddenProbe')
    return str(app.run('HiddenProbe.Tried', index))


KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(index: int) -> None:
    assert _run(index) == RECORD['cases'][index][1]


#: Breaks and panes, for the reasons in LEFT_OUT, Errors' enumerator, which Excel has none of, and two that VBA binds late
#: however the variable is typed, as an index returns an Object.
NOT_TYPED = ('ws.HPageBreaks _Default', 'ws.VPageBreaks _Default', 'ws.Range("A1").Errors _NewEnum', 'wb.IconSets(1) _Default', 'wb.IconSets(1) _NewEnum',
             'ch.SeriesCollection(1).Trendlines _Default', 'ch.SeriesCollection(1).Trendlines _NewEnum')


def test_hidden_members_on_typed_variables() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    source = '\n'.join(line for line in RECORD['source'].split('\n') if 't = TypeName(ActiveWindow.Panes' not in line)
    app.add_module(source, name='HiddenProbe')
    answers = str(app.run('HiddenProbe.Typed')).split('%%')[1:]
    names = [f'{name} {hidden}' for name in RECORD['typed_names'] for hidden in ('_Default', '_NewEnum')
             if name != 'ActiveWindow.Panes']
    assert len(answers) == len(names)
    wrong = {name: pair for name, *pair in zip(names, answers, RECORD['typed'], strict=False)
             if pair[0] != pair[1] and name not in NOT_TYPED}
    assert not wrong


def test_what_is_not_followed_reports_itself() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Function Go() As Long\n    Go = ActiveWorkbook.TableStyles.Add("Mine").ShowAsAvailableTableStyle\nEnd Function\n', name='Reporting')
    with pytest.raises(VBAUnsupportedError):
        app.run('Reporting.Go')
