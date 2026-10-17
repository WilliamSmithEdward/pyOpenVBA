"""A displayed range query shares statistical populations without retaining stale reads."""
from __future__ import annotations

from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel import _conditional_stats as statistics
from pyopenvba.apps.excel._conditional import Rule
from pyopenvba.apps.excel._model import Cell, Worksheet


@pytest.mark.parametrize('property_name', ['Font.Bold', 'Font.Color', 'Interior.Color',
                                          'Borders(xlEdgeLeft).Color', 'Borders.Color'])
def test_displayed_range_builds_population_once_per_property(property_name: str, monkeypatch: Any) -> None:
    app = ExcelApplication()
    book = app.add_workbook()
    sheet = book.sheets_[0]
    sheet.cells_ = {(row, 1): Cell(value=float(row)) for row in range(1, 1001)}
    app.add_module('''Sub Setup()
Dim f As Object
Set f=Range("A:A").FormatConditions.AddTop10
f.Rank=3
f.Font.Bold=True
f.Font.Color=255
f.Interior.Color=255
f.Borders(xlLeft).Color=255
End Sub
Function ReadDisplayed() As Variant
ReadDisplayed=Range("A1:A1000").DisplayFormat.''' + property_name + '''
End Function
Sub ChangeInput()
Range("A1").Value=2000
End Sub
Function ReadFirst() As Boolean
ReadFirst=Range("A1").DisplayFormat.Font.Bold
End Function''', name='Harness')
    app.run('Setup')
    original = getattr(statistics, '_population')
    calls: list[tuple[int, int]] = []

    def counted(owner: Worksheet, rule: Rule) -> statistics.Population:
        calls.append((id(owner), id(rule)))
        return original(owner, rule)

    monkeypatch.setattr(statistics, '_population', counted)
    app.run('ReadDisplayed')
    assert len(calls) == 1
    # A later property read is a fresh snapshot, even without a value change.
    app.run('ReadDisplayed')
    assert len(calls) == 2
    assert app.run('ReadFirst') is False
    app.run('ChangeInput')
    assert app.run('ReadFirst') is True
    assert len(calls) == 4
