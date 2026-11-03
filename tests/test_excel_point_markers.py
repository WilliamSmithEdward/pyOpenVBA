"""One point's marker, replayed from live Excel (scripts/measure_point_markers.py, tests/fixtures/point_markers.json):
what each case reads, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/point_markers.json').read_text(encoding='utf-8'))
#: A point's Format and ClearFormats on a line or scatter chart are refused, and so is changing the type of a chart
#: whose points have settings of their own.
REFUSED = ('.Format.Fill', '.ClearFormats', ': ch.ChartType')
KEPT = [index for index, case in enumerate(RECORD['cases']) if not any(text in case[0] for text in REFUSED)]
SOURCE = '\n'.join(line for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n')
                   if not any(text in line for text in REFUSED))


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('point_markers')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(SOURCE, name='MarkerProbe')
    return {index: str(app.run('MarkerProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[RECORD['cases'][index][0] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
