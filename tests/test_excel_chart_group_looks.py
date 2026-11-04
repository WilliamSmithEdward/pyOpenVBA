"""The look of a chart group's lines and bars, its settings, a legend key and walking the collections, replayed from live
Excel (scripts/measure_chart_group_looks.py, tests/fixtures/chart_group_looks.json): what each case reads, and the chart
part it saves."""
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
    (Path(__file__).parent / 'fixtures/chart_group_looks.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Kinds of chart the model does not draw.
UNDRAWN = ('bubble:', 'pieOfPie:', 'barOfPie:', 'radar:')
#: Refused, which each case below checks: members yet to be built; a look over a hidden fill or line; a type change with
#: a look on the chart; hidden error bars cleared, which Excel leaves painted; no 3-D shading, which writes a material on
#: each series; and a style through the key of a series that refuses one.
REFUSED = ('[_Default]',
           'Format.Fill.ForeColor.RGB = 255: ch.ChartGroups(1).UpBars.Format.Fill.Visible = msoFalse',
           'Format.Fill.ForeColor.RGB = 255: ch.ChartGroups(1).DownBars.Format.Fill.Visible = msoFalse',
           'Bars.Format.Fill.Visible = msoFalse: ch.ChartGroups(1).UpBars.Format.Fill',
           'Bars.Format.Fill.Visible = msoFalse: ch.ChartGroups(1).DownBars.Format.Fill',
           ': ch.ChartType = xlColumnClustered: ch.ChartType = xlLine', ': ch.ChartType = xlPie:',
           'Line.Visible = msoFalse: ch.ChartGroups(1).DropLines.Format.Line.Visible = msoTrue',
           'Line.Visible = msoFalse: ch.ChartGroups(1).HiLoLines.Format.Line.Visible = msoTrue',
           'Line.Visible = msoFalse: ch.ChartGroups(1).SeriesLines.Format.Line.Visible = msoTrue',
           'LineStyle = xlNone: ch.ChartGroups(1).DropLines.Format.Line.ForeColor',
           'LineStyle = xlNone: ch.ChartGroups(1).HiLoLines.Format.Line.ForeColor',
           'LineStyle = xlNone: ch.ChartGroups(1).SeriesLines.Format.Line.ForeColor',
           'ErrorBars.Format.Line.Visible = msoFalse: ch.SeriesCollection(1).ErrorBars.ClearFormats',
           'Has3DShading = False', 'column: ch.Legend.LegendEntries(1).LegendKey.MarkerStyle =')
#: A fill on a line series, which the model does not take, where Excel keeps it for the marker ClearFormats leaves.
UNTAKEN = 'line: ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255: r = ch.Legend'
DRAWN = [index for index, name in enumerate(NAMES) if not name.startswith(UNDRAWN)]
REFUSALS = [index for index in DRAWN if any(text in NAMES[index] for text in REFUSED) and '.[_Default]' not in NAMES[index]]
KEPT = [index for index in DRAWN if index not in REFUSALS and not NAMES[index].startswith(UNTAKEN)]


def _source(kept: set[int]) -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
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
    folder = tmp_path_factory.mktemp('chart_group_looks')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(set(KEPT)), name='GroupLooksProbe')
    return {index: str(app.run('GroupLooksProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_saves_the_part_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', REFUSALS, ids=[NAMES[index] for index in REFUSALS])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source({index}), name='GroupLooksProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('GroupLooksProbe.Probe', str(tmp_path), index, index)


def test_a_look_stays_when_the_book_is_opened_and_saved_again(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    book = app.add_workbook()
    app.add_module('''Public Sub Make(ByVal path As String)
    Dim ws As Worksheet, ch As Object
    Set ws = ActiveSheet
    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))
    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))
    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))
    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData ws.Range("A1:C4"), xlColumns
    ch.ChartType = xlLine
    ch.ChartGroups(1).HasUpDownBars = True
    ch.ChartGroups(1).UpBars.Format.Fill.ForeColor.RGB = 255
    ch.ChartGroups(1).HasDropLines = True
    ch.ChartGroups(1).DropLines.Border.Color = 255
    ActiveWorkbook.SaveAs path, xlOpenXMLWorkbook
End Sub
''', name='Make')
    first = tmp_path / 'first.xlsx'
    app.run('Make.Make', str(first))
    del book
    again = ExcelApplication(sandbox=False)
    opened = again.open_workbook(str(first))
    second = tmp_path / 'second.xlsx'
    opened.save_as(str(second))
    part = chart_part(second)
    assert '<c:upBars><c:spPr><a:solidFill><a:srgbClr val="FF0000"/>' in part
    assert '<c:dropLines><c:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/>' in part
