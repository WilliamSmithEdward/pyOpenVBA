"""A chart duplicated, copied and pasted, replayed from live Excel (scripts/measure_chart_copies.py,
tests/fixtures/chart_copies.json): what each case reads, and every chart part it saves."""
from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError

sys.path.insert(0, str(Path(__file__).parent))
from test_excel_chart_markup import normalized  # noqa: E402

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_copies.json').read_text(encoding='utf-8'))
NAMES = [case[0] for case in RECORD['cases']]
#: Left out: a Chart read after its chart object went, which Excel still answers; Duplicate's Shape deleted or
#: duplicated again, which raises 1004; a cell's PasteSpecial of a chart, which adds no chart when it does not raise.
LEFT_OUT = ('.Cut', 'n.Delete', 'Set n = n.Duplicate', 'co.Delete', 'Application.CutCopyMode', 'PasteSpecial')
#: Excel's Paste of a chart raises 1004 at random, about two runs in five, and now and then hangs. So a failed paste
#: in the fixture tells nothing, but for these, which failed every time in repeated runs.
ALWAYS_FAIL = ('co.Copy: o.Paste o.Range("B2")', 'ch.ChartArea.Copy: d.Paste')
KEPT = [index for index, (name, answer, _) in enumerate(RECORD['cases']) if not any(text in name for text in LEFT_OUT)
        and (not ('.Paste' in name and answer.startswith('Err')) or name.startswith(ALWAYS_FAIL))]


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


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: normalized(package.read(name).decode('utf-8')) for name in package.namelist()
                if re.fullmatch(r'xl/charts/chart[0-9]+\.xml', name)}


@pytest.fixture(scope='module')
def run(tmp_path_factory: pytest.TempPathFactory) -> tuple[dict[int, str], Path]:
    folder = tmp_path_factory.mktemp('chart_copies')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(), name='CopyProbe')
    return {index: str(app.run('CopyProbe.Probe', str(folder), index, index)) for index in KEPT}, folder


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_case_as_excel_answered(run: tuple[dict[int, str], Path], index: int) -> None:
    assert run[0][index] == RECORD['cases'][index][1]


@pytest.mark.parametrize('index', KEPT, ids=[NAMES[index] for index in KEPT])
def test_each_chart_part_as_excel_saved_it(run: tuple[dict[int, str], Path], index: int) -> None:
    assert chart_parts(run[1] / f'c{index}.xlsx') == RECORD['cases'][index][2]


def _trial(body: str) -> Any:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Function Try() As Variant
    Dim d As Object, co As Object
    Set d = ActiveWorkbook.Worksheets(1)
    d.Range("A1:A3").Value = "k"
    d.Range("B1").Value = "v": d.Range("B2:B3").Value = 1
    Set co = d.ChartObjects.Add(300, 100, 300, 200)
    co.Chart.SetSourceData d.Range("A1:B3"), xlColumns
    {body}
End Function
''', name='Trial')
    return app.run('Trial.Try')


# Each of these answers is from repeated runs in Excel, where Paste fails at random: what it does when it does not.
@pytest.mark.parametrize('statement', ['co.Cut', 'd.Shapes(1).Cut'])
def test_a_chart_cut_goes_at_once_and_can_be_pasted(statement: str) -> None:
    assert _trial(f'{statement}: Try = d.ChartObjects.Count') == 0
    assert _trial(f'{statement}: d.Paste d.Range("A10"): '
                  'Try = d.ChartObjects.Count & d.ChartObjects(1).Name & d.ChartObjects(1).Left') == '1Chart 20'


def test_a_chart_copied_outlives_a_cell_selected_and_the_chart() -> None:
    assert _trial('co.Copy: d.Range("H20").Select: d.Paste: Try = d.ChartObjects(2).Name & " " '
                  '& d.ChartObjects(2).Left & " " & d.ChartObjects(2).Top') == 'Chart 2 336 285'
    assert _trial('co.Copy: co.Delete: d.Paste d.Range("A10"): Try = d.ChartObjects.Count & d.ChartObjects(1).Name'
                  ) == '1Chart 2'


def test_only_the_active_sheet_takes_a_chart() -> None:
    other = 'Set o = ActiveWorkbook.Worksheets.Add(After:=d): d.Activate: co.Copy: '
    with pytest.raises(VBARuntimeError, match='Paste method'):
        _trial(f'Dim o As Object: {other}o.Paste o.Range("B2")')
    assert _trial(f'Dim o As Object: {other}o.Activate: o.Range("B2").Select: o.Paste: '
                  'Try = o.ChartObjects(1).Name & " " & o.ChartObjects(1).Left') == 'Chart 1 48'


def test_cells_copied_end_a_chart_copied() -> None:
    assert _trial('co.Copy: d.Range("A1").Copy: d.Paste d.Range("F9"): Try = d.ChartObjects.Count & d.Range("F9").Value'
                  ) == '1k'


def test_a_chart_read_from_a_file_is_not_duplicated(tmp_path: Path) -> None:
    path = tmp_path / 'one.xlsx'
    _trial(f'ActiveWorkbook.SaveAs "{path}"')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(f'''Public Sub Try()
    Workbooks.Open("{path}").Worksheets(1).ChartObjects(1).Duplicate
End Sub
''', name='Trial')
    with pytest.raises(VBAUnsupportedError, match='read from a file'):
        app.run('Trial.Try')
