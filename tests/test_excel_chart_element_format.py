"""Format on a chart's elements other than its series, replayed from live Excel
(scripts/measure_chart_element_format.py, tests/fixtures/chart_element_format.json): what each case reads and the chart
part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_element_format.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_element_format')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='ElementFormatProbe')
    return str(app.run('ElementFormatProbe.Probe', str(folder), 0, len(RECORD['cases']) - 1)).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case for case, _, _ in RECORD['cases']])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])), ids=[case for case, _, _ in RECORD['cases']])
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statements', [
    # Taking away an element with a look: what it comes back with was not measured.
    'ch.Legend.Format.Fill.ForeColor.RGB = 255: ch.HasLegend = False',
    'ch.Legend.Format.Fill.ForeColor.RGB = 255: ch.Legend.Delete',
    'ch.ChartTitle.Format.Fill.ForeColor.RGB = 255: ch.HasTitle = False',
    'ch.ChartTitle.Format.Fill.ForeColor.RGB = 255: ch.ChartTitle.Delete',
    'ch.Axes(xlValue).AxisTitle.Format.Fill.ForeColor.RGB = 255: ch.Axes(xlValue).HasTitle = False',
    'ch.Axes(xlValue).AxisTitle.Format.Fill.ForeColor.RGB = 255: ch.Axes(xlValue).AxisTitle.Delete',
    'ch.Axes(xlValue).MajorGridlines.Format.Line.Weight = 3: ch.Axes(xlValue).HasMajorGridlines = False',
    'ch.Axes(xlValue).MajorGridlines.Format.Line.Weight = 3: ch.Axes(xlValue).MajorGridlines.Delete',
    'ch.Axes(xlValue).Format.Line.Weight = 3: ch.HasAxis(xlValue) = False',
    # Going on after a hide, and the kinds of chart and lines not measured.
    'ch.ChartArea.Format.Fill.Visible = msoFalse: ch.ChartArea.Format.Fill.Visible = msoTrue',
    'ch.ChartArea.Format.Line.Visible = msoFalse: ch.ChartArea.Format.Line.Weight = 3',
    'ch.ChartArea.Format.Fill.ForeColor.RGB = 255: ch.ChartArea.Format.Fill.Visible = msoFalse',
    'ch.ChartArea.Format.Fill.ForeColor.RGB = 255: ch.ChartType = xlLine',
    'ch.Axes(xlValue).MinorGridlines.Format.Line.Weight = 3',
    'ch.ChartArea.Format.Line.DashStyle = msoLineRoundDot',
])
def test_an_unmeasured_step_is_refused(statements: str) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = xlColumnClustered\n'
                   '    ch.HasTitle = True: ch.HasLegend = True\n    ch.Axes(xlValue).HasTitle = True\n'
                   '    ch.Axes(xlValue).HasMinorGridlines = True\n'
                   f'    {statements}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')


def test_a_shadow_beside_a_look_is_refused_at_the_save(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go(ByVal path As String)\n    Dim co As Object\n'
                   '    Set co = ActiveSheet.ChartObjects.Add(10, 10, 200, 100)\n'
                   '    co.Chart.ChartArea.Format.Fill.ForeColor.RGB = 255: co.Shadow = True\n'
                   '    ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook\nEnd Sub\n', name='Both')
    with pytest.raises(VBAUnsupportedError):
        app.run('Both.Go', str(tmp_path / 'both.xlsx'))
