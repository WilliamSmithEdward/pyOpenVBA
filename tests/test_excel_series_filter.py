"""Series.IsFiltered, replayed from live Excel (scripts/measure_series_filter.py, tests/fixtures/series_filter.json):
what the chart and its series answer with a series filtered out, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/series_filter.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What gives odd answers around a filtered series, and is refused: a new series, new source data and a new plot order.
REFUSED = ('.NewSeries', 'ch.SetSourceData', '.PlotOrder = 2')
#: Left out: a scatter chart with every series filtered out reads ChartType 74, and a filtered series given literal
#: values keeps PlotBy where one given a literal name loses it, for no rule we have. A chart's references do not yet follow
#: rows and columns inserted or deleted, whether a series is filtered or not.
EVERY_SCATTER = ('scatter: ch.FullSeriesCollection(1).IsFiltered = True: ch.FullSeriesCollection(2).IsFiltered = True: '
                 'ch.FullSeriesCollection(3).IsFiltered = True')
LEFT_OUT = ('.Values = Array(7, 8, 9)', 'ws.Range("B:B").Delete', 'ws.Rows(2).Insert')
KEPT = [index for index, name in enumerate(NAMES)
        if not any(text in name for text in (*REFUSED, *LEFT_OUT)) and name != EVERY_SCATTER]


def _source() -> str:
    """The probe as recorded, with the statements of the refused cases taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept and not stripped.startswith('ch.ChartType = xl'):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('series_filter')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='FilterProbe')
    return {index: str(app.run('FilterProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statement', ['ch.SeriesCollection.NewSeries', 'ch.SetSourceData ActiveSheet.Range("B1:B3")',
                                       'ch.SeriesCollection(1).PlotOrder = 1'])
def test_what_gives_odd_answers_is_refused(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:C3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("B1:C3")\n    ch.FullSeriesCollection(1).IsFiltered = True\n'
                   f'    {statement}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
