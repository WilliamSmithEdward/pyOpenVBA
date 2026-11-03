"""Chart.ApplyLayout, replayed from live Excel (scripts/measure_apply_layout.py, tests/fixtures/apply_layout.json): what
each layout reads afterwards, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/apply_layout.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: A layout with a ChartType, and a scatter chart's layouts 3 and 9, which add trendlines, are refused.
REFUSED = (', xl', 'scatter: ch.ApplyLayout 3', 'scatter: ch.ApplyLayout 9')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in REFUSED)]


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
        elif current >= 0 and current not in kept and not stripped.startswith('ch.ChartType = xl'):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('apply_layout')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='LayoutProbe')
    return {index: str(app.run('LayoutProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize(('chart_type', 'statement'), [('xlColumnClustered', 'ch.ApplyLayout 1, xlLine'),
                                                       ('xlXYScatter', 'ch.ApplyLayout 3'),
                                                       ('xlXYScatter', 'ch.ApplyLayout 9')])
def test_what_is_not_modelled_is_refused(chart_type: str, statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = {chart_type}\n'
                   f'    {statement}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
