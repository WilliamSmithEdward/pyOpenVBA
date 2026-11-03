"""A chart when the cells it reads are taken away, replayed from live Excel (scripts/measure_chart_lost_cells.py,
tests/fixtures/chart_lost_cells.json): what each case reads, and the chart part its workbook saves."""
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
    (Path(__file__).parent / 'fixtures/chart_lost_cells.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: What the model refuses, because Excel's answer depends on more than it holds.
REFUSED = ('d.Delete', 'ErrorBar xlY, xlBoth, xlCustom, d.Range("E2:E4"), d.Range("E2:E4"): d.Columns(5).Delete')


def _kept(name: str) -> bool:
    return not any(text in name for text in REFUSED)


KEPT = [index for index, name in enumerate(NAMES) if _kept(name)]


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
        elif current >= 0 and current not in kept and not stripped.startswith(('ch.SetSourceData', 'ch.ChartType')):
            continue
        written.append(line)
    return '\n'.join(written)


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_lost_cells')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='LostProbe')
    return {index: str(app.run('LostProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_part(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


def test_a_saved_chart_keeps_what_it_lost(tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Sub Try()
    Dim wb As Workbook, d As Worksheet, ch As Object
    Set wb = ActiveWorkbook
    Set d = wb.Worksheets(1): d.Name = "Data"
    d.Range("A1:C4").Value = 1
    d.Range("A1:A4").Value = "k"
    d.Range("B1").Value = "v": d.Range("C1").Value = "w"
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData d.Range("A1:C4"), xlColumns
    d.Columns(2).Delete
    wb.SaveAs "{tmp_path / 'once.xlsx'}"
    d.Range("A1").Value = "z"
    wb.SaveAs "{tmp_path / 'twice.xlsx'}"
End Sub

Public Function Reopened() As String
    Dim wb As Workbook
    Set wb = Workbooks.Open("{tmp_path / 'twice.xlsx'}")
    With wb.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1)
        Reopened = .Name & "|" & .Points.Count
    End With
End Function
''', name='Trial')
    app.run('Trial.Try')
    assert chart_part(tmp_path / 'once.xlsx').count('<c:v>#REF!</c:v>') == 1
    assert chart_part(tmp_path / 'twice.xlsx').count('<c:v>#REF!</c:v>') == 1
    assert app.run('Trial.Reopened') == '#REF!|1'


def test_taking_away_an_error_bars_cells_is_refused() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Sub Try()
    Dim d As Worksheet, ch As Object
    Set d = ActiveWorkbook.Worksheets(1)
    d.Range("A1:B3").Value = 1
    d.Range("E1:E3").Value = 1
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.SetSourceData d.Range("A1:B3"), xlColumns
    ch.SeriesCollection(1).ErrorBar xlY, xlBoth, xlCustom, d.Range("E2:E3"), d.Range("E2:E3")
    d.Columns(5).Delete
End Sub
''', name='Trial')
    with pytest.raises(VBAUnsupportedError, match='an error bar'):
        app.run('Trial.Try')
