"""Series.Points and Point, replayed from live Excel (scripts/measure_series_points.py,
tests/fixtures/series_points.json): what each case reads and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/series_points.json').read_text(encoding='utf-8'))
#: What the model refuses, the replay running those cases with no statements and leaving their answers out: a look or
#: InvertIfNegative on a point of a line or scatter chart, whose reads move onto its marker members, Delete, which takes
#: away a series, and ChartType with a point set apart.
LINES = ('4', '65', '-4169')
LOOK = re.compile(r'pt\.(Format\.|InvertIfNegative)')
OTHER = ('51: pt.Delete', '51: pt.Format.Fill.ForeColor.RGB = 255: ch.ChartType = xlLine')


def refused(case: str) -> bool:
    kind, statements = case.split(': ', 1)
    return (kind in LINES and bool(LOOK.search(statements))) or case in OTHER


KEPT = [index for index, (case, _, _) in enumerate(RECORD['cases']) if not refused(case)]


def _source() -> str:
    """The probe as recorded, with the refused statements taken out and this platform's separator in its paths."""
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for index, (case, _, _) in enumerate(RECORD['cases']):
        if refused(case):
            kind, statements = case.split(': ', 1)
            start = (f'    Case {index}\n        ch.ChartType = {kind}\n        Set sr = ch.SeriesCollection(1)\n'
                     '        Set pt = sr.Points(2)\n')
            assert f'{start}        {statements}\n' in source
            source = source.replace(f'{start}        {statements}\n', start)
    return source


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('series_points')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='PointsProbe')
    return str(app.run('PointsProbe.Probe', str(folder), 0, len(RECORD['cases']) - 1)).split('@@@'), folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('case', [one for one, _, _ in RECORD['cases'] if refused(one)])
def test_a_refused_case_reports_itself(case: str) -> None:
    kind, statements = case.split(': ', 1)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ws As Object, ch As Object, sr As Object, pt As Object\n'
                   '    Set ws = ActiveSheet\n    ws.Range("B1:B3").Value = 1\n'
                   '    Set ch = ws.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ws.Range("B1:B3")\n    ch.ChartType = {kind}\n'
                   f'    Set sr = ch.SeriesCollection(1)\n    Set pt = sr.Points(2)\n    {statements}\nEnd Sub\n',
                   name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
