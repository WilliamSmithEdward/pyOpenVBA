"""A chart's budgets beside cells that move: a cell shift carries a series' reference past the cell or the text budget,
and what reads it then reports the gap and changes nothing. The budgets stand far past what a macro charts (a
million cells, a reference of 32,768 characters), so they are accepted limits; the test lowers them to reach them."""
from __future__ import annotations

import pytest

from pyopenvba.apps.excel import ExcelApplication, _charts
from pyopenvba.exceptions import VBAUnsupportedError

BUILD = '''Public Sub Build()
    Dim d As Worksheet, ch As Object
    Set d = ActiveWorkbook.Worksheets(1)
    d.Range("A1:B10").Value = 1
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.ChartType = xlColumnClustered
    ch.SetSourceData d.Range("B2:B9")
End Sub

Public Sub Edit()
    ActiveWorkbook.Worksheets(1).Rows(3).Insert
End Sub

Public Function Formula() As String
    Formula = ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Formula
End Function

Public Function Count() As Long
    Count = UBound(ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Values)
End Function
'''


def _app() -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(BUILD, name='Chart')
    app.run('Chart.Build')
    return app


@pytest.mark.parametrize('budget, value, gap', [('MAX_CHART_POINTS', 8, 'cell budget'),
                                                ('MAX_CHART_REFERENCE_TEXT', len('Sheet1!$B$2:$B$9'), 'text budget')])
def test_a_reference_moved_past_a_budget_reports_where_it_is_read(budget: str, value: int, gap: str,
                                                                  monkeypatch: pytest.MonkeyPatch) -> None:
    app = _app()
    assert app.run('Chart.Count') == 8
    monkeypatch.setattr(_charts, budget, value)
    app.run('Chart.Edit')
    formula = app.run('Chart.Formula')
    assert '$B$2:$B$10' in str(formula)
    with pytest.raises(VBAUnsupportedError, match=gap):
        app.run('Chart.Count')
    assert app.run('Chart.Formula') == formula
