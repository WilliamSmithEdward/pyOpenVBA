"""Application.DoubleClick, replayed from live Excel (scripts/measure_double_click.py, tests/fixtures/double_click.json):
the events a sink saw, the sheets, the active cell and selection, and the cells, after a double-click on a value, a
formula, nothing, a table's cell, and a pivot table's value."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/double_click.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = RECORD['cases']
#: A pivot table's value, whose detail Excel shows on a new sheet as Range.ShowDetail does: not built.
DETAIL = {index for index, (statement, _, _) in enumerate(CASES) if 'CreatePivotTable' in statement}


def _run(index: int) -> str:
    _, cancels, _ = CASES[index]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['log'], name='EventLog')
    app.add_module(RECORD['sink'], name='ExcelSink', kind='class')
    app.add_module(RECORD['source'], name='ProbeModule')
    return str(app.run('ProbeModule.Probe', index, cancels))


@pytest.mark.parametrize('index', sorted(set(range(len(CASES))) - DETAIL), ids=str)
def test_a_double_click_does_what_excels_does(index: int) -> None:
    assert _run(index) == CASES[index][2]


@pytest.mark.parametrize('index', sorted(DETAIL), ids=str)
def test_a_pivot_values_detail_reports_itself(index: int) -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    with pytest.raises(VBAUnsupportedError, match="pivot table's value"):
        _run(index)


def test_a_double_click_on_a_pivot_tables_label_reports_itself() -> None:
    from pyopenvba.exceptions import VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(
        'Function Run() As String\nDim ws As Worksheet, pt As PivotTable\nSet ws = ActiveWorkbook.Worksheets(1)\n'
        'ws.Range("A1:B3").Value = ws.Evaluate("{""k"",""v"";""a"",1;""b"",2}")\n'
        'Set pt = ActiveWorkbook.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("D1"))\n'
        'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v")\n'
        'ws.Range("D2").Select\nApplication.DoubleClick\nEnd Function\n', name='Label')
    with pytest.raises(VBAUnsupportedError, match='Show Detail dialog'):
        app.run('Run')
