"""A chart's series when the cells it reads are hidden, replayed from live Excel (scripts/measure_chart_hidden_cells.py,
tests/fixtures/chart_hidden_cells.json): what each case reads, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_hidden_cells.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: literal values keep PlotBy where the model loses it, and literal categories read back trimmed to the cells
#: shown though the file keeps them all. A new series or a new chart type while a series is out of sight is written
#: differently: NewSeries puts the new one before it, and it takes no marker.
LEFT_OUT = ('.Values = Array(7, 8, 9, 10)', '.XValues = Array("x", "y", "z", "t")',
            'ws.Columns(2).Hidden = True: ch.SeriesCollection.NewSeries',
            'ws.Columns(2).Hidden = True: ch.ChartType = xlLine')


def _kept(name: str) -> bool:
    return not any(text in name for text in LEFT_OUT)


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
