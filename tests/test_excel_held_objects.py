"""What an object answers when a macro still holds it after it has gone, replayed from live Excel (scripts/measure_held_objects.py,
tests/fixtures/held_objects.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/held_objects.json').read_text(encoding='utf-8'))
#: The probe's run ended instead of answering: Excel hung, or died, on one of these members of something that had gone.
ENDED = {'TIMEOUT', 'RUNNER-ERROR', 'VBA-ERROR', 'MODAL-BLOCKED'}
CASES = [case for case in RECORD['cases'] if case[5] not in ENDED]
IDS = [f'{index}: {case[0]} | {case[2]} | {case[3]} | {case[4][0] if len(case[4]) == 1 else "*"}' for index, case in enumerate(CASES)]
_STATEMENTS = ('h.Delete', 'h.Unlist', 'h.Resize', 'h.Text "', 'h.Activate', 'h.Select', 'h.Modify', 'h.SetFirst', 'h.Refresh', 'h.BringTo')


def source(case: list[Any]) -> str:
    _, made, held, gone, tries = case[:5]
    lines = ['Public Function Go() As String', '    Dim wb As Object, ws As Object, t As Object, h As Object, s2 As Variant, t2 As Variant, v As Variant, s As String', '    Set wb = ActiveWorkbook: Set ws = wb.Worksheets(1): ws.Name = "Sheet1"',
             '    wb.Worksheets.Add(After:=ws).Name = "Sheet2"', '    ws.Activate', '    On Error Resume Next', f'    {made}'.rstrip(': '), f'    Set h = {held}', '    s = "made " & Err.Number: Err.Clear', f'    {gone}', '    s = s & "~gone " & Err.Number: Err.Clear',
             '    v = "": v = TypeName(h): s = s & "~" & v & "/" & Err.Number: Err.Clear', '    v = "": v = (h Is Nothing): s = s & "~" & v & "/" & Err.Number: Err.Clear']
    for one in tries:
        lines += [f'    {one}: s = s & "~" & Err.Number: Err.Clear' if ' = ' in one or one == 'h.Add' or one.startswith(_STATEMENTS) else f'    v = Empty: v = {one}: s = s & "~" & Err.Number: Err.Clear']
    return '\n'.join([*lines, '    Go = s', 'End Function']) + '\n'


def answered(case: list[Any]) -> str:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case), name='Held')
    return str(app.run('Held.Go'))


def labelled(case: list[Any], answer: str) -> dict[str, str]:
    return dict(zip(['made', 'gone', 'type', 'nothing', *case[4]], answer.split('~'), strict=True))


@pytest.mark.parametrize('case', CASES, ids=IDS)
def test_each_case_as_excel_answered(case: list[Any]) -> None:
    assert labelled(case, answered(case)) == labelled(case, case[5])
