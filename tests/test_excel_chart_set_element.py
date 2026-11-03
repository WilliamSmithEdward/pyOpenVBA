"""Chart.SetElement, replayed from live Excel (scripts/measure_chart_set_element.py,
tests/fixtures/chart_set_element*.json): what each element reads afterwards, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError
from test_excel_chart_markup import chart_part

FIXTURES = ('chart_set_element', 'chart_set_element_bar_plain', 'chart_set_element_scatter_area_doughnut')
RECORDS: dict[str, dict[str, Any]] = {
    name: json.loads((Path(__file__).parent / f'fixtures/{name}.json').read_text(encoding='utf-8'))
    for name in FIXTURES}
#: What the model refuses, on any kind that takes it: a data label callout and taking away the plot area write markup
#: it cannot hold, and an area chart's error bars are not modelled.
REFUSED = ('msoElementDataLabelCallout', 'msoElementPlotAreaNone')
AREA_REFUSED = ('msoElementErrorBarPercentage', 'msoElementErrorBarStandardDeviation', 'msoElementErrorBarStandardError')


def refused(case: list[Any]) -> bool:
    name, answer = case[0], case[1]
    if str(answer).startswith('Err'):
        return False
    return name.endswith(REFUSED) or (name.startswith('area: ') and name.endswith(AREA_REFUSED))


CASES = [(file, index) for file in FIXTURES for index, case in enumerate(RECORDS[file]['cases']) if not refused(case)]
IDS = [f"{file[18:] or 'first'}-{RECORDS[file]['cases'][index][0]}" for file, index in CASES]


def _source(file: str) -> str:
    """The probe as recorded, with the statements of the refused cases taken away."""
    record = RECORDS[file]
    written: list[str] = []
    current = -1
    for line in record['source'].replace('folder & "\\', f'folder & "{os.sep}').split('\n'):
        stripped = line.strip()
        if line.startswith('    Case '):
            current = int(stripped[5:])
        elif stripped == 'End Select':
            current = -1
        elif current >= 0 and refused(record['cases'][current]) and not stripped.startswith('ch.ChartType'):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[tuple[str, int], str], Path]:
    folder = tmp_path_factory.mktemp('chart_set_element')
    answers: dict[tuple[str, int], str] = {}
    for file in FIXTURES:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(_source(file), name='ElementProbe')
        place = folder / file
        place.mkdir()
        for name, index in CASES:
            if name == file:
                answers[name, index] = str(app.run('ElementProbe.Probe', str(place), index, index))
    return answers, folder


@pytest.mark.parametrize(('file', 'index'), CASES, ids=IDS)
def test_each_case_as_excel_answered(run: tuple[dict[tuple[str, int], str], Path], file: str, index: int) -> None:
    assert run[0][file, index] == RECORDS[file]['cases'][index][1]


@pytest.mark.parametrize(('file', 'index'), CASES, ids=IDS)
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[tuple[str, int], str], Path], file: str, index: int) -> None:
    assert chart_part(run[1] / file / f'c{index}.xlsx') == RECORDS[file]['cases'][index][2]


@pytest.mark.parametrize('element', ['msoElementDataLabelCallout', 'msoElementPlotAreaNone'])
def test_markup_the_model_cannot_hold_is_refused(element: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("B1:B3").Value = 1\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   f'    ch.SetSourceData ActiveSheet.Range("B1:B3")\n    ch.SetElement {element}\nEnd Sub\n',
                   name='Refused')
    with pytest.raises(VBAUnsupportedError):
        app.run('Refused.Go')
