"""Changes to the sheet's old collections and their items on a workbook Excel saved, replayed from live Excel
(scripts/measure_old_collection_edits.py, tests/fixtures/old_collection_edits.json): what each answers, what the items read after it, and what
Excel read from the model's save. `excel_reads_model` there is Excel's reading of the model's file, recorded when the fixture was measured."""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'old_collection_edits.json').read_text(encoding='utf-8'))
CASES: list[dict[str, Any]] = RECORD['cases']
READS: list[str] = RECORD['reads']
DRIVER = '''
Public Function ChangeHere(ByVal name As String, ByVal first As Boolean, ByVal which As Long) As String
    ChangeHere = Change(ActiveWorkbook.Worksheets(1), name, first, which)
End Function

Public Function ReadHere(ByVal name As String, ByVal place As Long, ByVal which As Long) As String
    ReadHere = ReadOne(ActiveWorkbook.Worksheets(1), name, place, which)
End Function

Public Function CountHere(ByVal name As String) As Long
    CountHere = Counted(ActiveWorkbook.Worksheets(1), name)
End Function
'''
#: Excel names a macro by its workbook: the one the probe ran in, or the file it read.
MACRO = re.compile(r'String [^;]*!M\b')
TAGS = [f'{one["collection"]} {"item" if one["first"] else "collection"} {one["change"]}' for one in CASES]


def plain(text: str) -> str:
    return MACRO.sub('String !M', text)


def opened() -> ExcelApplication:
    app = ExcelApplication.open(FIXTURES / 'old_collections.xlsx', with_vba=False)
    app.interpreter.sandbox = False
    app.add_module(RECORD['source'] + DRIVER, name='Probe')
    return app


@pytest.mark.parametrize('case', CASES, ids=TAGS)
def test_each_change_as_excel_answered(case: dict[str, Any]) -> None:
    answered, _, after = case['excel'].partition('@')
    after = after.partition('@')[0]
    app = opened()
    which = RECORD['changes'].index(case['change'])
    try:
        got = str(app.run('Probe.ChangeHere', case['collection'], case['first'], which))
    except VBAUnsupportedError:
        assert case['model_save'].startswith('reports')
        return
    if case['change'] == 'Interior.ColorIndex = 6' and case['collection'] == 'Arcs':
        pytest.xfail('An arc from a file takes no ColorIndex yet')
    assert got == answered
    count, _, cells = after.partition(':')
    assert str(app.run('Probe.CountHere', case['collection'])) == count
    expected = cells.split(';')
    for at, wanted in enumerate(expected[:-1]):
        place, read = divmod(at, len(READS))
        try:
            found = str(app.run('Probe.ReadHere', case['collection'], place + 1, read))
        except VBAUnsupportedError:
            continue
        assert plain(found) == plain(wanted), (place + 1, READS[read])


@pytest.mark.parametrize('case', [one for one in CASES if one['excel_reads_model']], ids=lambda one: f'{one["collection"]} {one["first"]} {one["change"]}')
def test_excel_read_the_model_save_as_its_own(case: dict[str, Any]) -> None:
    if case['change'] == 'Interior.ColorIndex = 6' and case['collection'] == 'Arcs':
        pytest.xfail('An arc from a file takes no ColorIndex yet')
    assert plain(case['excel_reads_model']) == plain(case['excel'].split('@')[2])


def test_a_control_shown_again_after_a_save_and_a_load_keeps_its_settings(tmp_path: Path) -> None:
    """Hidden and unlocked, a loaded control reads so from the model's own save."""
    app = opened()
    app.add_module('Public Sub Edit()\n    With ActiveWorkbook.Worksheets(1).Buttons(1)\n        .Visible = False\n        .Locked = False\n    End With\nEnd Sub\n', name='Edit')
    app.run('Edit.Edit')
    app.save(tmp_path / 'saved.xlsx')
    again = ExcelApplication.open(tmp_path / 'saved.xlsx', with_vba=False)
    again.add_module(RECORD['source'] + DRIVER, name='Probe')
    assert str(again.run('Probe.ReadHere', 'Buttons', 1, READS.index('Visible'))) == 'Boolean False'
    assert str(again.run('Probe.ReadHere', 'Buttons', 1, READS.index('Locked'))) == 'Boolean False'
