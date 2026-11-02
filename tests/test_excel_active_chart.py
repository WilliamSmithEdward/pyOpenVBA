"""ActiveChart on Application, Workbook and Window, replayed from live Excel (scripts/measure_active_chart.py,
tests/fixtures/active_chart.json)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/active_chart.json').read_text(encoding='utf-8'))
#: Cases that leave the chart area selected, which the model does not answer as a Selection, or that select it
#: through ChartArea, which it does not have. The replay runs them with no statements and leaves their answers out.
REFUSED = ('co.Activate: ActiveChart.ChartType = xlLine', 'co.Activate: ws.Shapes(2).Select',
           'co.Activate: co.Chart.ChartArea.Select')
SOURCE = '\n'.join(line for line in RECORD['source'].split('\n') if line.strip() not in REFUSED)
KEPT = [index for index, (case, _) in enumerate(RECORD['cases']) if case not in REFUSED]


@pytest.fixture(scope='module')
def answers() -> list[str]:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='ActiveChartProbe')
    return str(app.run('ActiveChartProbe.Probe')).split('@@@')


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(answers: list[str], index: int) -> None:
    assert answers[index] == RECORD['cases'][index][1]


def test_changing_the_active_chart_leaves_a_selection_the_model_refuses() -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go() As String\n    Dim co As Object\n'
                   '    Set co = ActiveSheet.ChartObjects.Add(10, 10, 200, 100)\n'
                   '    co.Activate\n    ActiveChart.HasTitle = True\n    Go = TypeName(Selection)\nEnd Function\n',
                   name='Changed')
    with pytest.raises(VBAUnsupportedError, match='chart area'):
        app.run('Changed.Go')
