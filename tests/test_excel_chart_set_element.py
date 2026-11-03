"""Chart.SetElement, replayed from live Excel (scripts/measure_chart_set_element.py,
tests/fixtures/chart_set_element.json): what each element reads afterwards, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_set_element.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
REFUSED = ('msoElementDataLabelCallout', 'msoElementPlotAreaNone')
KEPT = [index for index, name in enumerate(NAMES) if not any(name.endswith(text) for text in REFUSED)]


def _source() -> str:
    """The probe as recorded, with the statements of the refused cases taken away."""
    kept = set(KEPT)
    written: list[str] = []
    current = -1
    for line in RECORD['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and current not in kept and not stripped.startswith('ch.ChartType'):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_set_element')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='ElementProbe')
    return {index: str(app.run('ElementProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]
