"""Combo charts, replayed from live Excel: a series' own type and axes, the groups and secondary axes they bring, and the part Excel
saves (scripts/measure_combo_charts.py, tests/fixtures/combo_charts.json), then a workbook Excel saved, opened in the model
(scripts/measure_combo_charts_saved.py, tests/fixtures/combo_chart_saved.xlsx, tests/fixtures/combo_charts_saved.json)."""
from __future__ import annotations

import json
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'combo_charts.json').read_text(encoding='utf-8'))
SAVED: dict[str, Any] = json.loads((FIXTURES / 'combo_charts_saved.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Refused, and each case below checks it reports itself: series of types that force axes of their own, chart types the model does
#: not draw, charts of those kinds on secondary axes, setting PlotOrder across groups, and scales and crossings that need Excel's
#: layout. Their answers stay in the fixture for when they are built.
REFUSALS = [4, 5, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 48, 54, 55, 85, 86, 118, 120, 123, 124, 125, 126, 153, 154, 155, 156, 187, 227, 228, 229]
KEPT = [index for index in range(len(NAMES)) if index not in REFUSALS]


def _run(index: int, folder: Path, app: ExcelApplication) -> object:
    return app.run('P.Probe', str(folder), index, index)


def _plain(text: str) -> str:
    """A part without Excel's random ids, the axes' numbered as they first appear."""
    text = re.sub(r'\{[0-9A-F]{8}-[0-9A-F-]{27}\}', '{id}', text)
    seen: dict[str, str] = {}

    def numbered(found: re.Match[str]) -> str:
        seen.setdefault(found.group(2), str(len(seen)))
        return f'{found.group(1)}val="{seen[found.group(2)]}"'

    text = re.sub(r'(<c:(?:axId|crossAx) )val="(\d+)"', numbered, text)
    return re.sub(r'<c16r2:[^>]*>|</c16r2:[^>]*>|<c:date1904 val="0"/>|<c:lang [^>]*/>|<c:roundedCorners [^>]*/>', '', text)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('combo_charts')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    return {index: str(_run(index, folder, app)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[f'{index}: {NAMES[index][:70]}' for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', [one for one in KEPT if RECORD['cases'][one][2]], ids=[f'{one}: {NAMES[one][:70]}' for one in KEPT if RECORD['cases'][one][2]])
def test_each_case_saves_the_part_excel_saved(run: tuple[dict[int, str], Path], index: int) -> None:
    with zipfile.ZipFile(run[1] / f'c{index}.xlsx') as package:
        ours = package.read('xl/charts/chart1.xml').decode('utf-8')
    assert _plain(ours) == _plain(next(iter(RECORD['cases'][index][2].values())))


@pytest.mark.parametrize('index', REFUSALS, ids=[f'{one}: {NAMES[one][:70]}' for one in REFUSALS])
def test_what_is_not_built_says_so(tmp_path: Path, index: int) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['source'], name='P')
    with pytest.raises(VBAUnsupportedError):
        _run(index, tmp_path, app)


def _opened(tmp_path: Path) -> tuple[ExcelApplication, Path]:
    copy = tmp_path / 'saved.xlsx'
    shutil.copy(FIXTURES / 'combo_chart_saved.xlsx', copy)
    return ExcelApplication(sandbox=False), copy


def test_excel_saved_charts_read_as_excel_reads_them(tmp_path: Path) -> None:
    app, copy = _opened(tmp_path)
    app.add_workbook()
    app.add_module(SAVED['source'], name='Saved')
    assert app.run('Saved.Reads', str(copy)) == SAVED['answers']


def test_an_excel_saved_chart_is_not_changed_and_saves_as_it_came(tmp_path: Path) -> None:
    _, copy = _opened(tmp_path)
    app = ExcelApplication.open(copy, with_vba=False)
    app.add_module('Public Sub Go()\n    ActiveWorkbook.Worksheets(1).Range("H9").Value = 1\nEnd Sub\n', name='Elsewhere')
    app.run('Elsewhere.Go')
    out = tmp_path / 'out.xlsx'
    app.save(out)
    with zipfile.ZipFile(copy) as before, zipfile.ZipFile(out) as after:
        for name in before.namelist():
            if re.fullmatch(r'xl/charts/chart\d+\.xml', name):
                assert before.read(name) == after.read(name)


@pytest.mark.parametrize('edit', ['ch.SeriesCollection(1).ChartType = xlLine', 'ch.SeriesCollection(1).AxisGroup = xlSecondary', 'ch.Axes(2, 2).HasTitle = False', 'ch.ChartGroups(1).GapWidth = 33'])
def test_editing_an_excel_saved_chart_reports_itself(tmp_path: Path, edit: str) -> None:
    _, copy = _opened(tmp_path)
    app = ExcelApplication.open(copy, with_vba=False)
    app.add_module(f'Public Sub Go()\n    Dim ch As Object\n    Set ch = ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart\n    {edit}\nEnd Sub\n', name='Edited')
    with pytest.raises(VBAUnsupportedError):
        app.run('Edited.Go')
