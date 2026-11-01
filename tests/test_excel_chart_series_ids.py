"""Series idx and unique ids replayed from live Excel (scripts/measure_chart_series_ids.py,
tests/fixtures/chart_series_ids.json): ChartObjects.Add numbers later unique ids past the series it would have plotted
from the selection, and SetSourceData, NewSeries and Delete go on from there as Excel's do."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_series_ids.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def series_ids(path: Path) -> list[list[int]]:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/charts/chart1.xml').decode('utf-8')
    return [[int(idx), int(order), int(unique, 16)] for idx, order, unique in re.findall(
        r'<c:ser><c:idx val="(\d+)"/><c:order val="(\d+)"/>.*?<c16:uniqueId val="\{([0-9A-F]{8})-', xml)]


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_series_ids')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='IdsProbe')
    return str(app.run('IdsProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_series_numbered_as_excel_numbered_it(run: tuple[list[str], Path], index: int) -> None:
    assert series_ids(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
