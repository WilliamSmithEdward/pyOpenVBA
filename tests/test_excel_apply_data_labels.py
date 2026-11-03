"""ApplyDataLabels, replayed from live Excel (scripts/measure_apply_data_labels.py, tests/fixtures/apply_data_labels.json):
what each case reads, and the chart part its workbook saves."""
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
    (Path(__file__).parent / 'fixtures/apply_data_labels.json').read_text(encoding='utf-8'))
#: AutoText:=False writes a field with a random id on a point, and empties every label on a series. It is refused.
REFUSED = 'AutoText:=False'
KEPT = [index for index, case in enumerate(RECORD['cases']) if REFUSED not in case[0]]


def _source() -> str:
    lines = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n')
    return '\n'.join(line for line in lines if REFUSED not in line)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('apply_data_labels')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='ApplyProbe')
    return {index: str(app.run('ApplyProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('owner', ['ch', 'ch.SeriesCollection(1)', 'ch.SeriesCollection(1).Points(2)'])
def test_a_field_with_a_random_id_is_refused(owner: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.ChartType = xlColumnClustered\n'
                   f'    {owner}.ApplyDataLabels AutoText:=False\nEnd Sub\n', name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
