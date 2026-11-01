"""Chart parts replayed from live Excel (scripts/measure_chart_markup.py, tests/fixtures/chart_markup.json): each
chart made through ChartObjects.Add, saved after the case's edits, holds the part Excel saved, its random axis and
series ids aside, and reads back as Excel's did."""
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
    (Path(__file__).parent / 'fixtures/chart_markup.json').read_text(encoding='utf-8'))
#: The probe as recorded, its files joined to the folder with this platform's separator.
SOURCE = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')


def normalized(xml: str) -> str:
    numbers: dict[str, str] = {}

    def axis(match: re.Match[str]) -> str:
        return f'{match.group(1)} val="{numbers.setdefault(match.group(2), str(len(numbers) + 1))}"'

    xml = re.sub(r'(<c:(?:axId|crossAx)) val="(\d+)"', axis, xml)
    return re.sub(r'(<c16:uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]{27}\}', r'\1}', xml)


def chart_part(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        names = [entry for entry in package.namelist() if entry.startswith('xl/charts/')]
        assert len(names) == 1 and re.fullmatch(r'xl/charts/chart\d+\.xml', names[0]), names
        return normalized(package.read(names[0]).decode('utf-8'))


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[list[str], Path]:
    folder = tmp_path_factory.mktemp('chart_markup')
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(SOURCE, name='MarkupProbe')
    return str(app.run('MarkupProbe.Probe', str(folder))).split('@@@'), folder


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_reads_as_excel_read_it(run: tuple[list[str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', range(len(RECORD['cases'])))
def test_each_chart_part_as_excel_saved_it(run: tuple[list[str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
