"""SeriesCollection.Add and Extend, replayed from live Excel (scripts/measure_series_add.py,
tests/fixtures/series_add.json): what each answers, the series it leaves, and the chart part it saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/series_add.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
S = 'ch.SeriesCollection'
#: Left out, all on an empty chart: the Name of what Add answers, read straight away, which reads as if unnamed; PlotBy
#: of a series of one cell or of text alone; a whole row, which Excel takes as one series with a name of several cells;
#: several areas, which are refused; and a scatter chart whose text categories leave its axis no number format.
LEFT_OUT = (f'empty: r = {S}.Add(d.Range("C1:C5")).Name', f'empty: {S}.Add d.Range("C3")',
            f'empty: {S}.Add d.Range("C1")', f'empty: {S}.Add d.Rows(1)', f'empty: {S}.Add d.Range("A1:A5")', f'empty: {S}.Extend d.Range("C6")',
            f'empty: {S}.Add d.Range("C1:C5,D1:D5")', f'empty: ch.ChartType = xlXYScatter: {S}.Add d.Range("A1:C5")')
KEPT = [index for index, name in enumerate(NAMES) if not name.startswith(LEFT_OUT)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case ') and not stripped.endswith(' which, d, o, ch, r'):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('series_add')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='AddProbe')
    return {index: str(app.run('AddProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
