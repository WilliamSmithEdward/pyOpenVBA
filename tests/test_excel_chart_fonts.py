"""Font on a chart's text, replayed from live Excel (scripts/measure_chart_fonts.py, tests/fixtures/chart_fonts.json):
what each case reads and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/chart_fonts.json').read_text(encoding='utf-8'))
#: Superscript reads back in Excel and writes nothing, which the model refuses rather than keep: the replay runs those
#: cases with no statement and leaves their answers out.
REFUSED = 'el.Font.Superscript = True'
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').replace(f'        {REFUSED}\n', '')
KEPT = [index for index, (case, _, _) in enumerate(RECORD['cases']) if REFUSED not in case]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_fonts')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='FontsProbe')
    return str(app.run('FontsProbe.Probe', str(folder), 0, len(RECORD['cases']) - 1)).split('@@@'), folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statements', [
    'ch.Legend.Font.Superscript = True',
    # Whether the chart area's font reaches the others was not measured.
    'ch.ChartArea.Font.Size = 14: x = ch.Legend.Font.Size',
    'ch.Legend.Font.Size = 14: ch.ChartArea.Font.Size = 14',
    # Taking away an element with a font: what it comes back with was not measured.
    'ch.Legend.Font.Size = 14: ch.HasLegend = False',
    'ch.ChartTitle.Font.Size = 14: ch.HasTitle = False',
    'ch.ChartTitle.Font.Size = 14: ch.ChartTitle.Delete',
    'ch.Axes(xlValue).AxisTitle.Font.Size = 14: ch.Axes(xlValue).AxisTitle.Delete',
    'ch.Axes(xlValue).TickLabels.Font.Size = 14: ch.HasAxis(xlValue) = False',
    'ch.Legend.Font.Size = 14: ch.ChartType = xlLine',
    'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.Font.Size = 14: '
    'ch.ChartType = xlLine',
    'ch.SeriesCollection(1).HasDataLabels = True: ch.SeriesCollection(1).DataLabels.Font.Size = 14: '
    'ch.SeriesCollection(1).HasDataLabels = False',
    # The settings no case measured.
    'ch.Legend.Font.Size = 450',
    'ch.Legend.Font.Underline = xlUnderlineStyleDouble',
    'ch.Legend.Font.Size = 10.25',
    'ch.Legend.Font.ColorIndex = xlColorIndexAutomatic',
    'ch.Legend.Font.Strikethrough = False',
])
def test_an_unmeasured_step_is_refused(statements: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object, x As Variant\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = xlColumnClustered\n'
                   '    ch.HasTitle = True: ch.HasLegend = True\n    ch.Axes(xlValue).HasTitle = True\n'
                   f'    {statements}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
