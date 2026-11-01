"""Chart.DisplayBlanksAs, PlotVisibleOnly, ShowDataLabelsOverMaximum and DisplayValueNotAvailableAsBlank replayed
from live Excel (scripts/measure_chart_display.py, tests/fixtures/chart_display.json): what each case answers, the
markup its saved chart holds after the plot area, and what a file that leaves the options out reads once opened."""
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
    (Path(__file__).parent / 'fixtures/chart_display.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def chart_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        name = next(entry for entry in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', entry))
        return package.read(name).decode('utf-8')


def display_markup(path: Path) -> str:
    found = re.search(r'</c:plotArea>(.*?)</c:chart>', chart_xml(path), re.S)
    assert found is not None
    return re.sub(r'<c:legend>.*?</c:legend>', '', found.group(1), flags=re.S)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[ExcelApplication, list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_display')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='DisplayProbe')
    return app, str(app.run('DisplayProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_case_as_excel_answered(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert run[1][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_as_excel_saved_it(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    assert display_markup(run[2] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', range(len(RECORD['reopened'])))
def test_left_out_options_read_as_excel_reads_them(run: tuple[ExcelApplication, list[str], Path], index: int) -> None:
    app, _, folder = run
    name, answer = RECORD['reopened'][index]
    pattern, replacement = RECORD['strips'][name]
    target = folder / f'reopened{index}.xlsx'
    with zipfile.ZipFile(folder / 'c0.xlsx') as old, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if re.fullmatch(r'xl/charts/chart\d+\.xml', item.filename):
                data = re.sub(pattern, replacement, data.decode('utf-8'), flags=re.S).encode('utf-8')
            new.writestr(item, data)
    assert app.run('DisplayProbe.Reopen', str(target)) == answer
