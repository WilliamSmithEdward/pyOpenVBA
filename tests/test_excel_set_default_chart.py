"""Chart.SetDefaultChart against live Excel (scripts/measure_set_default_chart.py, tests/fixtures/set_default_chart.json),
measured under a snapshot of the Office registry, which the run left as it found it.

xlBuiltIn is taken and Charts.Add goes on making a clustered column chart. Excel takes a chart type too, and Charts.Add
then makes that type; how such a chart saves was not measured, so the model refuses it."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD = json.loads((Path(__file__).parent / 'fixtures/set_default_chart.json').read_text(encoding='utf-8'))
SETUP = '''Public Function Go(ByVal statement As Long) As String
    Dim wb As Workbook, ws As Worksheet, c As Object, made As Object, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    ws.Range("A1").Select
    Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    c.SetSourceData ws.Range("A1:C3")
    s = "before " & TypeName(c.ChartType) & " " & c.ChartType
    On Error Resume Next
    Err.Clear
    If statement = 0 Then c.SetDefaultChart xlBuiltIn Else c.SetDefaultChart xlLine
    s = s & "|set " & Err.Number
    Err.Clear
    ws.Activate
    Set made = wb.Charts.Add
    s = s & "|add " & Err.Number & " " & made.ChartType
    Go = s
End Function
'''


def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(SETUP, name='P')
    return made


def test_the_built_in_default_is_taken_and_changes_nothing() -> None:
    answer = RECORD['answer'].split('|')
    assert app().run('P.Go', 0) == '|'.join([answer[0], 'set ' + answer[3].split(' ')[-1], answer[4]])


def test_a_chart_type_as_the_default_is_refused() -> None:
    assert RECORD['answer'].split('|')[1:3] == ['set line 0', 'add 0 4']
    with pytest.raises(VBAUnsupportedError, match='SetDefaultChart'):
        app().run('P.Go', 1)
