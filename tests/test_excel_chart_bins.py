"""A chart group's histogram bin members, replayed from live Excel (scripts/measure_chart_bins.py,
tests/fixtures/chart_bins.json): what each case reads, and the chart part it saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/chart_bins.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_bins')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}'), name='BinsProbe')
    return {index: str(app.run('BinsProbe.Probe', str(folder), index, index)) for index in range(len(NAMES))}, folder


@pytest.mark.parametrize('index', range(len(NAMES)), ids=NAMES)
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(NAMES)), ids=NAMES)
def test_each_case_saves_the_part_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
