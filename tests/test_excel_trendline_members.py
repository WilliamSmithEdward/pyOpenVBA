"""A trendline's Forward2, Backward2 and label, and Axis.MinorUnitScale, replayed from live Excel
(scripts/measure_trendline_members.py, tests/fixtures/trendline_members.json): what each case reads, and the chart
part."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/trendline_members.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: a trendline label's text, number format, font and Format set or read.
LEFT_OUT = ('Trendlines(1).DataLabel.Text = ', 'Trendlines(1).DataLabel.NumberFormat = ',
            'Trendlines(1).DataLabel.Font', 'Trendlines(1).DataLabel.Format')
KEPT = [index for index, name in enumerate(NAMES) if not any(text in name for text in LEFT_OUT)]


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
    folder = tmp_path_factory.mktemp('trendline_members')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='MemberProbe')
    return {index: str(app.run('MemberProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
