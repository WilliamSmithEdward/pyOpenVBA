"""Chart.ChartStyle, ChartColor, ClearToMatchStyle, ApplyLayout and ChartWizard, replayed from live Excel
(scripts/measure_chart_layouts.py, tests/fixtures/chart_layouts.json): what each case reads, and the chart part its
workbook saves."""
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

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/chart_layouts.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
REFUSED = re.compile(r'ch\.ChartStyle = (?:2\d\d|3\d\d)|ch\.ChartColor = (?:1|10|14|26)$'
                     r'|ch\.ApplyLayout (?:[1-9]|1[01])\b'
                     r'|line: ch\.ApplyLayout 12'
                     r'|ChartWizard.*(?:CategoryLabels|SeriesLabels)|ChartWizard PlotBy:=xlRows'
                     r'|HasDataLabels = True.*ClearToMatchStyle')
#: Duplicating a chart adds a second part, which the reader of chart parts does not take.
LEFT_OUT = 'ch.Parent.Duplicate'
KEPT = [index for index, name in enumerate(NAMES) if not REFUSED.search(name) and LEFT_OUT not in name]


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
    folder = tmp_path_factory.mktemp('chart_layouts')
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


@pytest.mark.parametrize('statement', ['ch.ChartStyle = 201', 'ch.ChartColor = 10', 'ch.ApplyLayout 1',
                                       'ch.ChartWizard SeriesLabels:=0', 'ch.ChartWizard PlotBy:=xlRows'])
def test_what_needs_more_of_the_model_is_refused(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    {statement}\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
