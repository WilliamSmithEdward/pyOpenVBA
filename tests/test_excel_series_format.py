"""Series.Format, replayed from live Excel (scripts/measure_series_format.py, tests/fixtures/series_format.json): what
each case reads and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/series_format.json').read_text(encoding='utf-8'))
#: Cases the model refuses: one series' own chart type, a marker on a series with a look, whose line Excel copies into
#: it, and a line colour on a series with markers, or a scatter series, which reaches the marker too. The replay runs
#: them with no statements and leaves their answers out.
REFUSED = ['51: sr.Format.Fill.ForeColor.RGB = 255: sr.ChartType = xlLine',
           '4: sr.Format.Line.ForeColor.RGB = 255: sr.MarkerStyle = xlMarkerStyleCircle',
           '65: sr.Format.Line.ForeColor.RGB = 255', '-4169: sr.Format.Line.ForeColor.RGB = 255']
KEPT = [index for index, (case, _, _) in enumerate(RECORD['cases']) if case not in REFUSED]


def _source() -> str:
    """The probe as recorded, with the refused statements taken out and this platform's separator in its paths."""
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for index, (case, _, _) in enumerate(RECORD['cases']):
        if case in REFUSED:
            kind, statements = case.split(': ', 1)
            start = (f'    Case {index}\n        ch.ChartType = {kind}\n        Set sr = ch.SeriesCollection(1)\n'
                     '        Set s2 = ch.SeriesCollection(2)\n')
            source = source.replace(f'{start}        {statements}\n', start)
    return source


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('series_format')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_source(), name='SeriesFormatProbe')
    return str(app.run('SeriesFormatProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('case', REFUSED)
def test_a_refused_case_reports_itself(case: str) -> None:
    kind, statements = case.split(': ', 1)
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object, sr As Object\n'
                   '    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = {kind}\n'
                   f'    Set sr = ch.SeriesCollection(1)\n    {statements}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
