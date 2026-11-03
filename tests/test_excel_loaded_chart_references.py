"""A chart read from a file when the cells it reads move or hide, replayed from live Excel
(scripts/measure_loaded_chart_references.py, tests/fixtures/loaded_chart_references.json): what each case reads, and the
references and cached points in the chart parts its workbook saves."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/loaded_chart_references.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
INDEXES = list(range(len(NAMES)))


def _parts(path: Path) -> list[list[list[object]]]:
    """Each chart part's references with their cached points, as the script gathers them from Excel's."""
    data: list[list[list[object]]] = []
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if re.fullmatch(r'xl/charts/chart[0-9]+\.xml', name):
                xml = package.read(name).decode('utf-8')
                found = re.findall(r'<c:f>(.*?)</c:f>(?:<c:(?:num|str)Cache>(.*?)</c:(?:num|str)Cache>)?', xml,
                                   re.DOTALL)
                data.append(sorted([reference, re.findall(r'<c:v>(.*?)</c:v>', cache)]
                                   for reference, cache in found))
    return sorted(data, key=len, reverse=True)


def _built(folder: Path) -> Path:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['build'], name='Builder')
    app.run('Builder.Build')
    path = folder / 'base.xlsx'
    app.save(path)
    return path


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('loaded_chart_references')
    base = _built(folder)
    answers: dict[int, str] = {}
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    for index in INDEXES:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(source, name='LoadedProbe')
        answers[index] = str(app.run('LoadedProbe.Probe', str(base), str(folder), index))
    return answers, folder


@pytest.mark.parametrize('index', INDEXES, ids=NAMES)
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', INDEXES, ids=NAMES)
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert _parts(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


def test_only_references_and_their_caches_change_in_the_part(run: tuple[dict[int, str], Path]) -> None:
    def part(name: str, changed: bool) -> str:
        with zipfile.ZipFile(run[1] / name) as package:
            xml = package.read('xl/charts/chart1.xml').decode('utf-8')
        xml = re.sub(r'<c:((?:num|str)Cache)>.*?</c:\1>', '', xml, flags=re.DOTALL)
        return re.sub(r'<c:f>.*?</c:f>', '', xml) if changed else xml
    assert part('base.xlsx', True) == part(f'c{NAMES.index("d.Rows(3).Insert")}.xlsx', True)
    assert part('base.xlsx', False) != part(f'c{NAMES.index("d.Rows(3).Insert")}.xlsx', False)


def _opened(path: Path, statement: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Sub Try(ByVal path As String)
    Dim wb As Workbook, d As Worksheet
    Set wb = Workbooks.Open(path)
    Set d = wb.Worksheets("Data")
    {statement}
End Sub
''', name='Trial')
    app.run('Trial.Try', str(path))
    return app


def test_a_second_save_keeps_the_rewritten_references(tmp_path: Path) -> None:
    _opened(_built(tmp_path), f'd.Rows(1).Insert: wb.SaveAs "{tmp_path / "once.xlsx"}"')
    _opened(tmp_path / 'once.xlsx', f'd.Rows(1).Insert: wb.SaveAs "{tmp_path / "twice.xlsx"}"')
    assert 'Data!$B$4:$B$7' in str(_parts(tmp_path / 'twice.xlsx'))


def test_taking_away_the_cells_a_loaded_chart_reads_is_refused(tmp_path: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match='Taking away the cells'):
        _opened(_built(tmp_path), 'd.Columns(3).Delete')


def test_moving_cells_between_a_loaded_charts_points_is_refused(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(RECORD['build'] + '''
Public Sub Paint()
    ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255
End Sub
''', name='Builder')
    app.run('Builder.Build')
    app.run('Builder.Paint')
    path = tmp_path / 'painted.xlsx'
    app.save(path)
    with pytest.raises(VBAUnsupportedError, match='between the points'):
        _opened(path, 'd.Rows(3).Insert')
    _opened(path, 'd.Rows(1).Insert')
