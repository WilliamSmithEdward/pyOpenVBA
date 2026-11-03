"""The Format of a chart's trendlines and error bars, replayed from live Excel (scripts/measure_chart_line_format.py,
tests/fixtures/chart_line_format.json): what each case reads, and the chart part it saves."""
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
    (Path(__file__).parent / 'fixtures/chart_line_format.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused: theme and scheme colours, arrowheads, showing a hidden line again, msoCTrue, transparency on a line with a
#: look, and the shadow, glow and soft edge of a line.
REFUSED = ('ObjectThemeColor = msoTheme', 'SchemeColor', 'ArrowheadStyle = msoArrow', 'L.Visible = msoCTrue',
           'L.Visible = 2', 'L.Visible = msoFalse: L.Visible = msoTrue', 'L.Weight = 2: L.Transparency = 0.5',
           'F.Shadow', 'F.Glow', 'F.SoftEdge')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in REFUSED)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case ') and not stripped.endswith(' which, ws, ch, r'):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_line_format')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='LineProbe')
    return {index: str(app.run('LineProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


def test_showing_a_hidden_line_again_is_refused() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Try()
    Dim d As Worksheet, ch As Object
    Set d = ActiveWorkbook.Worksheets(1)
    d.Range("A1:A3").Value = "k"
    d.Range("B1").Value = "v": d.Range("B2:B3").Value = 1
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData d.Range("A1:B3"), xlColumns
    ch.SeriesCollection(1).Trendlines.Add
    ch.SeriesCollection(1).Trendlines(1).Format.Line.Visible = msoFalse
    ch.SeriesCollection(1).Trendlines(1).Format.Line.Visible = msoTrue
End Sub
''', name='Trial')
    with pytest.raises(VBAUnsupportedError, match='Showing the line again'):
        app.run('Trial.Try')
