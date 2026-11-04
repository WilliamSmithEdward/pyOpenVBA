"""A fraction given to a Long property through an Object, against live Excel (scripts/measure_vba_late_fractions.py)."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication

ANSWERS: dict[str, str] = json.loads((Path(__file__).parent / 'fixtures/vba_late_fractions.json').read_text(encoding='utf-8'))['answers']
#: Tab.ColorIndex's refusal and ChartStyle past its list answer by what the model holds, which these lines are not about.
MODELLED = sorted(key for key in ANSWERS if key.split('.')[0] in ('o', 'ch') and 'ChartStyle' not in key and not key.endswith('-1.5'))


@pytest.mark.parametrize('key', MODELLED)
def test_application_cuts_a_fraction_and_a_chart_rounds_it(key: str) -> None:
    target, value = key.split(' ')
    holder, member = target.split('.', 1)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Function Probe() As String
    Dim o As Object, ch As Object, ws As Object
    Set ws = ActiveSheet
    ws.Range("A1:A3").Value = 1
    Set ch = ws.ChartObjects.Add(1, 1, 200, 100).Chart
    Set o = Application
    On Error Resume Next
    {holder}.{member} = 1
    Err.Clear
    {holder}.{member} = {value}
    Probe = Err.Number & ":" & {holder}.{member}
End Function''', name='Harness')
    assert app.run('Harness.Probe') == ANSWERS[key]
