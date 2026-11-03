"""A chart's series when the cells it reads are hidden, replayed from live Excel (scripts/measure_chart_hidden_cells.py,
tests/fixtures/chart_hidden_cells.json): what each case reads, and the chart part its workbook saves."""
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
    (Path(__file__).parent / 'fixtures/chart_hidden_cells.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What the model refuses, because Excel does more than leave the cells out: a series or its categories wholly hidden,
#: and a point's look or label on a series with hidden cells, which stay with a cell.
REFUSED = ('ws.Rows("2:5")', 'ws.Columns(1).Hidden = True', 'ws.Columns(2).Hidden = True',
           'ws.Columns(3).Hidden = True', 'ws.Columns("B:C")', 'ws.Columns(2).ColumnWidth = 0', '">9"',
           'Points(2).Format', 'Points(3).Format',
           'Points(2).HasDataLabel')
#: With series in rows the same edits fall differently.
ROWS_REFUSED = ('ws.Rows(1).Hidden = True', 'ws.Rows(2).Hidden = True', 'ws.Columns(1).Hidden = True')
#: Left out: literal values keep PlotBy where the model loses it, and literal categories read back trimmed to the cells
#: shown though the file keeps them all.
LEFT_OUT = ('.Values = Array(7, 8, 9, 10)', '.XValues = Array("x", "y", "z", "t")')


def _kept(name: str) -> bool:
    if any(text in name for text in LEFT_OUT):
        return False
    if name.startswith('column rows'):
        return not name.endswith(ROWS_REFUSED)
    return 'PlotVisibleOnly = False: ws.Columns' in name or not any(text in name for text in REFUSED)


KEPT = [index for index, name in enumerate(NAMES) if _kept(name)]


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
        elif current >= 0 and current not in kept and not stripped.startswith(('ch.SetSourceData', 'ch.ChartType')):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_hidden_cells')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='HiddenProbe')
    return {index: str(app.run('HiddenProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


HIDDEN_LABEL = 'ch.SeriesCollection(1).Points(2).HasDataLabel = True: ws.Rows(3).Hidden = True'


@pytest.mark.parametrize('statement', ['ws.Columns(2).Hidden = True: ws.Columns(3).Hidden = True',
                                       'ws.Columns(1).Hidden = True', HIDDEN_LABEL])
def test_what_is_not_modelled_is_refused(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Function Go() As Long\n    Dim ws As Object, ch As Object\n    Set ws = ActiveSheet\n'
                   '    ws.Range("A1:C4").Value = 1\n    ws.Range("A1:C1").Value = Array("k", "v", "w")\n'
                   '    ws.Range("A2:A4").Value = "a"\n    Set ch = ws.ChartObjects.Add(300, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ws.Range("A1:C4"), xlColumns\n    {statement}\n'
                   '    Go = ch.SeriesCollection(1).Points.Count\nEnd Function\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
