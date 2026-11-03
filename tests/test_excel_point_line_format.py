"""One point's Format and ClearFormats on line and scatter charts and its Delete, replayed from live Excel
(scripts/measure_point_line_format.py, tests/fixtures/point_line_format.json): what each case reads, and the chart part
its workbook saves."""
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
    (Path(__file__).parent / 'fixtures/point_line_format.json').read_text(encoding='utf-8'))
P = 'ch.SeriesCollection(1).Points(2)'
S = 'ch.SeriesCollection(1)'


def refused(case: str) -> bool:
    """What the model refuses: hiding a point's line, Solid on its marker, any line setting of a scatter chart's point,
    and a series' own line on a line chart with markers, which has not been modelled."""
    kind, statements = case.split(': ', 1)
    return (f'{P}.Format.Line.Visible = msoFalse' in statements or f'{P}.Format.Fill.Solid' in statements
            or (kind == 'scatter' and '.Format.Line.' in statements)
            or (kind == 'markers' and f'{S}.Format.Line.' in statements))


def _skipped(case: str) -> bool:
    """On a line chart a series' Fill reads back through its markers without a trace in the file, which the model does
    not hold, and a coloured automatic marker reads size 7 after ClearFormats, for no rule we have. A scatter chart left
    with no series writes a number format on its axis, and a pie turns its leader lines off. Selecting a pie's point
    writes an empty title in its file, which the model does not hold; tests/test_excel_chart_element_select.py replays
    what Selection answers there."""
    kind, statements = case.split(': ', 1)
    return ((kind == 'line' and f'{S}.Format.Fill.' in statements) or (kind == 'pie' and '.Select' in statements)
            or case == f'line: {P}.MarkerBackgroundColor = RGB(0, 255, 0): {P}.ClearFormats'
            or case in (f'scatter: {P}.Delete: {P}.Delete', f'pie: {P}.Delete: {P}.Delete'))


KEPT = [index for index, case in enumerate(RECORD['cases']) if not refused(case[0]) and not _skipped(case[0])]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept and not stripped.startswith('ch.ChartType'):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('point_line_format')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='FormatProbe')
    return {index: str(app.run('FormatProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('case', [case[0] for case in RECORD['cases'] if refused(case[0])])
def test_a_refused_case_reports_itself(case: str) -> None:
    kind, statements = case.split(': ', 1)
    constant = {'markers': 'xlLineMarkers', 'line': 'xlLine', 'scatter': 'xlXYScatter'}[kind]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ws As Object, ch As Object\n    Set ws = ActiveSheet\n'
                   '    ws.Range("B1:B3").Value = 1\n    Set ch = ws.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ws.Range("B1:B3")\n    ch.ChartType = {constant}\n    {statements}\nEnd Sub\n',
                   name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
