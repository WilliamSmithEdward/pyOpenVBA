"""One point's DataLabel, replayed from live Excel (scripts/measure_data_label.py, tests/fixtures/data_label.json): what
each case reads, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/data_label.json').read_text(encoding='utf-8'))
#: Refused, so left out and tested below: a field with a random id, and a label's place and size, which depend on the
#: display.
REFUSED = ['el.AutoText = False', 'el.Formula = "=Sheet1!$A$1"', 'el.FormulaR1C1 = "=Sheet1!R1C1"', 'el.Left = 50',
           'el.Top = 50', 'el.Width = 50', 'el.Height = 50']
#: Excel's reads lag behind a cell edit here until it draws the chart, which a model has no moment for.
LAGGING = ['Range("B2").Value = 9', 'Range("B3").Value = 9', 'Range("A3").Value = "z"']
LEFT_OUT = [*REFUSED, *LAGGING]
#: What the series' labels do to a point's when those are the only ones is refused too, and a pie's Select, which
#: brings the chart a title.
SERIES_CHANGES = ['ch.SeriesCollection(1).DataLabels.ShowCategoryName = True',
                  'el.Text = "typed": ch.SeriesCollection(1).DataLabels.ShowCategoryName = True',
                  'el.ShowCategoryName = True: ch.SeriesCollection(1).DataLabels.ShowValue = False',
                  'ch.SeriesCollection(1).DataLabels.NumberFormat = "0.00"',
                  'ch.SeriesCollection(1).DataLabels.Position = xlLabelPositionCenter',
                  'ch.SeriesCollection(1).DataLabels.Font.Size = 14', 'ch.SeriesCollection(1).HasDataLabels = False']
LEFT_OUT_HERE = {*(f'point: {one}' for one in SERIES_CHANGES), 'pie: el.Select'}


def _kept(name: str) -> bool:
    return name.split(': ', 1)[1] not in LEFT_OUT and name not in LEFT_OUT_HERE


KEPT = [index for index, case in enumerate(RECORD['cases']) if _kept(case[0])]


_SETTING_UP = ('ch.SeriesCollection(1).HasDataLabels = True', 'ch.SeriesCollection(1).Points(2).HasDataLabel = True',
               'ch.ChartType = xlPie', 'Set el = ')


def _source() -> str:
    """The probe with the statements of the cases left out taken away, since they would raise."""
    lines = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n')
    kept = set(KEPT)
    out: list[str] = []
    index = None
    for line in lines:
        if line.startswith('    Case '):
            index = int(line.split()[1])
            out.append(line)
        elif index is not None and index not in kept and line.startswith('        ') and not line.strip().startswith(
                _SETTING_UP):
            continue
        else:
            out.append(line)
        if line.startswith('    End Select'):
            index = None
    return '\n'.join(out)


SOURCE = _source()


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('data_label')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='LabelProbe')
    return {index: str(app.run('LabelProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('statement', REFUSED)
def test_an_unmeasured_step_is_refused(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object, el As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = xlColumnClustered\n'
                   '    ch.SeriesCollection(1).HasDataLabels = True\n'
                   '    Set el = ch.SeriesCollection(1).Points(2).DataLabel\n'
                   f'    {statement}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
