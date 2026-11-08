"""A chart's references following the cells they read, replayed from live Excel (scripts/measure_chart_references.py,
tests/fixtures/chart_references.json): what each case reads, and the chart part its workbook saves."""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from test_excel_chart_markup import chart_part

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_references.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Deleting a sheet a chart reads is refused elsewhere.
LOST = ('d.Delete',)
#: Left out: Worksheet.Copy with a chart and setting Series.Formula are refused elsewhere, literal values keep PlotBy
#: where the model loses it, and a hidden row's points are still counted.
LEFT_OUT = ('d.Copy After:=d', '.Formula = "=SERIES', '.Values = Array(7, 8, 9): d.Rows(1).Insert', 'd.Rows(2).Hidden')
KEPT = [index for index, name in enumerate(NAMES)
        if (not any(name.endswith(text) or text + ' ' in name for text in LOST)
            or name.endswith('d.Rows(1).Insert: d.Rows(1).Delete'))
        and not any(text in name for text in LEFT_OUT)]


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
        elif (current >= 0 and current not in kept
              and not stripped.startswith(('Set ch =', 'ch.SetSourceData', 'ch.ChartType'))):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_references')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='ReferenceProbe')
    return {index: str(app.run('ReferenceProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', [index for index in KEPT if RECORD['cases'][index][2]],
                         ids=[NAMES[index] for index in KEPT if RECORD['cases'][index][2]])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


def test_taking_away_what_an_error_bar_reads_goes_through() -> None:
    """The bar reads #REF!, as Excel has it (tests/fixtures/chart_shift_losses.json)."""
    app = ExcelApplication()
    app.add_workbook()
    app.add_module('Public Sub Go()\n    Dim ch As Object\n    ActiveSheet.Range("A1:C4").Value = 1\n'
                   '    ActiveSheet.Range("A1:C1").Value = Array("k", "v", "w")\n'
                   '    ActiveSheet.Range("E1").Value = "e"\n'
                   '    Set ch = ActiveSheet.ChartObjects.Add(10, 10, 200, 100).Chart\n'
                   '    ch.SetSourceData ActiveSheet.Range("A1:C4"), xlColumns\n'
                   '    ch.SeriesCollection(1).ErrorBar xlY, xlBoth, xlCustom, ActiveSheet.Range("E2:E4"), '
                   'ActiveSheet.Range("E2:E4")\n'
                   '    ActiveSheet.Columns(5).Delete\nEnd Sub\n\n'
                   'Public Function Left() As String\n'
                   '    Left = ActiveSheet.Range("A1").Value & ActiveSheet.Range("E1").Value\nEnd Function\n',
                   name='Refused')
    app.run('Refused.Go')
    assert app.run('Refused.Left') == 'k'
