"""Clearing a chart's parts, its legend entries and its chart objects, replayed from live Excel
(scripts/measure_chart_clears.py, tests/fixtures/chart_clears.json): what each case reads, and the chart part it
saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_clears.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: ChartArea.Clear, which takes the chart object away and leaves its chart answering; a legend entry's Format;
#: chart objects copied, cut, duplicated or selected together; a legend with a look taken away, which is refused, and
#: ClearContents on an axis with a look, which keeps a number format the model drops.
LEFT_OUT = ('ChartArea.Clear', 'LegendEntries(1).Format', 'Duplicate', '.Copy', '.Cut', 'ChartObjects.Add 10',
            'ChartObjects(Array', 'ch.Legend.Format.Fill.ForeColor.RGB = 255: ch.Legend.Clear')
KEPT = [index for index, name in enumerate(NAMES)
        if ('ChartArea.ClearContents' in name or 'ChartArea.ClearFormats' in name
            or not any(text in name for text in LEFT_OUT))
        and not ('ChartArea.ClearContents' in name and 'ch.Axes(xlValue).Format' in name)]


def _source() -> str:
    """The probe as recorded, with the statements of the cases left out taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept:
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_clears')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='ClearProbe')
    return {index: str(app.run('ClearProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
