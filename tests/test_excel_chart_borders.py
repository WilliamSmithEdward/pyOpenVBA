"""The Border of a chart's elements and a series' ClearFormats, replayed from live Excel
(scripts/measure_chart_borders.py, tests/fixtures/chart_borders.json): what each case reads, and the chart part it saves."""
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
    (Path(__file__).parent / 'fixtures/chart_borders.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: a trendline's and error bars' Border, which have no Format in the model to share.
LEFT_OUT = ('Trendlines(1).', '.ErrorBars.')
#: Refused: ClearFormats on a line series with a look, which Excel moves to its marker, on an area series with a point set
#: apart, which leaves it an empty dPt, and a line colour on these kinds, which Format.Line does not set. A fill on a line
#: series is not refused: the model takes none, where Excel keeps it for the marker ClearFormats leaves.
REFUSED = ('line: ch.SeriesCollection(1).Format.', 'lineMarkers: ch.SeriesCollection(1).Format.Line.',
           'scatter: ch.SeriesCollection(1).Format.Line.', 'area: ch.SeriesCollection(1).Points(2).')
KEPT = [index for index, name in enumerate(NAMES)
        if not any(text in name for text in LEFT_OUT) and not name.startswith(REFUSED)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case ') and not stripped.endswith(' which, ws, ch, r'):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept and not stripped.startswith(('ch.SetSourceData', 'ch.ChartType')):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_borders')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='BorderProbe')
    return {index: str(app.run('BorderProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


@pytest.mark.parametrize('index', [index for index, name in enumerate(NAMES)
                                   if name.startswith(REFUSED) and 'ClearFormats' in name
                                   and not name.startswith('line: ch.SeriesCollection(1).Format.Fill')])
def test_what_clear_formats_does_otherwise_is_refused(index: int, tmp_path: Path) -> None:
    source = RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='BorderProbe')
    with pytest.raises(VBAUnsupportedError):
        app.run('BorderProbe.Probe', str(tmp_path), index, index)
