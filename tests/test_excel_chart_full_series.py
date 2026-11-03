"""Chart.FullSeriesCollection, Series.IsFiltered, and Select on the other elements of a chart, replayed from live Excel
(scripts/measure_chart_full_series.py, tests/fixtures/chart_full_series.json): what each case answers, and the chart
part its workbook saves."""
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
    (Path(__file__).parent / 'fixtures/chart_full_series.json').read_text(encoding='utf-8'))
FILTERING = '.IsFiltered = True'


def refused(case: str) -> bool:
    """What the model refuses: filtering a series out, selecting a pie's label, a scatter chart's gridlines, and reading
    ChartTitle once HasTitle was set False, which the model answers by giving the chart its title back."""
    return (FILTERING in case or (case.startswith('pie: ') and 'DataLabel.Select' in case)
            or (case.startswith('scatter: ') and 'MajorGridlines' in case)
            or 'ch.HasTitle = False: ch.ChartTitle.Select' in case)


KEPT = [index for index, case in enumerate(RECORD['cases']) if not refused(case[0])]
#: Selecting anything of a pie writes an empty title in its file, and taking away a title that was selected sets
#: autoTitleDeleted; the model holds neither. What they answer is still compared.
NAMES = [case[0] for case in RECORD['cases']]
FILE_KEPT = [index for index in KEPT if not (NAMES[index].startswith('pie: ') and '.Select' in NAMES[index])
             and 'ch.ChartTitle.Select: ch.HasTitle = False' not in NAMES[index]]


def _source() -> str:
    """The probe as recorded, with the statements of the refused cases taken away."""
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
    folder = tmp_path_factory.mktemp('chart_full_series')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='FullProbe')
    return {index: str(app.run('FullProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', FILE_KEPT, ids=[RECORD['cases'][index][0] for index in FILE_KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('owner', ['ch.FullSeriesCollection(1)', 'ch.SeriesCollection(1)'])
def test_filtering_a_series_out_is_refused(owner: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    {owner}.IsFiltered = True\nEnd Sub\n',
                   name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
