"""Chart.Location, replayed from live Excel (scripts/measure_chart_location.py, tests/fixtures/chart_location.json): what each
case answers, the sheets it leaves, the chart parts it saves, and what earns a chart its automatic title on the way."""

from __future__ import annotations

import importlib.util
import json
import os
import re
import zipfile
from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURE: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/chart_location.json').read_text(encoding='utf-8'))
CASES: list[list[Any]] = FIXTURE['cases']
#: What is left out, each saying so: a move off a chart sheet, whose geometry comes from the window, and a sheet's copy with another shape
#: beside its chart.
REFUSED = {
    68: 'Moving a chart off a chart sheet', 69: 'Moving a chart off a chart sheet', 70: 'Moving a chart off a chart sheet',
    71: 'Moving a chart off a chart sheet', 170: 'Worksheet.Copy with drawings',
}


def plain(text: str) -> str:
    """A chart part without what differs on every save: its ids, and the marks the model does not write."""
    text = re.sub(r'\{[0-9A-F]{8}-[0-9A-F-]{27}\}', '{id}', text)
    seen: dict[str, str] = {}

    def numbered(found: re.Match[str]) -> str:
        seen.setdefault(found.group(2), str(len(seen)))
        return f'{found.group(1)}val="{seen[found.group(2)]}"'

    text = re.sub(r'(<c:(?:axId|crossAx) )val="(\d+)"', numbered, text)
    return re.sub(r'<c16r2:[^>]*>|</c16r2:[^>]*>|<c:date1904 val="0"/>|<c:lang [^>]*/>|<c:roundedCorners [^>]*/>', '', text)


@pytest.fixture(scope='module')
def app() -> ExcelApplication:
    made = ExcelApplication(sandbox=False)
    made.add_workbook()
    made.add_module(FIXTURE['source'], name='P')
    return made


@pytest.fixture(scope='module')
def folder(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return tmp_path_factory.mktemp('chart_location')


@pytest.fixture(autouse=True)
def only_the_host_book(app: ExcelApplication) -> Iterator[None]:
    yield
    for book in app.application.workbooks_.books[1:]:
        book.vba_call('Close', [False])


ANSWERED = [index for index in range(len(CASES)) if index not in REFUSED]


@pytest.mark.parametrize('index', ANSWERED, ids=lambda index: CASES[index][0])
def test_each_case_answers_and_leaves_the_sheets_as_excel_did(index: int, app: ExcelApplication, folder: Path) -> None:
    assert str(app.run('P.Probe', str(folder), index, index)) == CASES[index][1]


@pytest.mark.parametrize('index', [index for index in ANSWERED if any(name.startswith('xl/charts/chart') for name in CASES[index][2])],
                         ids=lambda index: CASES[index][0])
def test_each_chart_part_as_excel_saved_it(index: int, app: ExcelApplication, folder: Path) -> None:
    app.run('P.Probe', str(folder), index, index)
    with zipfile.ZipFile(folder / f'c{index}.xlsx') as package:
        ours = [package.read(name).decode('utf-8') for name in package.namelist() if name.startswith('xl/charts/chart')]
    theirs = [text for name, text in CASES[index][2].items() if name.startswith('xl/charts/chart')]
    assert sorted(map(plain, ours)) == sorted(map(plain, theirs))


def test_a_book_from_a_template_constant_is_named_by_a_session_count() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(FIXTURE['book_names_source'], name='P')
    assert str(app.run('P.Names')) == FIXTURE['book_names']


SHEET_PARTS = re.compile(r'xl/chartsheets/(_rels/)?sheet\d+\.xml(\.rels)?|docProps/app\.xml')


@pytest.mark.parametrize('index', [index for index in ANSWERED if any('chartsheets' in name for name in CASES[index][2])],
                         ids=lambda index: CASES[index][0])
def test_each_chart_sheet_part_as_excel_saved_it(index: int, app: ExcelApplication, folder: Path) -> None:
    def same_id(text: str) -> str:
        return re.sub(r'\{[0-9A-F-]{36}\}', '{id}', text)

    app.run('P.Probe', str(folder), index, index)
    with zipfile.ZipFile(folder / f'c{index}.xlsx') as package:
        for name, text in CASES[index][2].items():
            if SHEET_PARTS.fullmatch(name):
                assert same_id(package.read(name).decode('utf-8')) == same_id(text), name


@pytest.mark.parametrize('index', sorted(REFUSED), ids=lambda index: CASES[index][0])
def test_what_is_left_out_reports_itself(index: int, app: ExcelApplication, folder: Path) -> None:
    with pytest.raises(VBAUnsupportedError, match=REFUSED[index]):
        app.run('P.Probe', str(folder), index, index)


READ_BACK = '''
Public Function Probe(ByVal path As String) As String
    Dim wb As Workbook, t As Chart
    Set wb = Workbooks.Open(path)
    Set t = wb.Charts(1)
    Probe = wb.Sheets(1).Name & " " & TypeName(wb.Sheets(1)) & " " & wb.Sheets(2).Name & " " & t.SeriesCollection.Count & " " & t.SeriesCollection(1).Formula
    wb.Close False
End Function
'''


@pytest.mark.skipif(os.environ.get('RUN_LIVE_EXCEL') != '1' or importlib.util.find_spec('pyvbaharness') is None,
                    reason='requires explicitly enabled native Excel')
def test_excel_opens_the_chart_sheet_a_save_wrote(tmp_path: Path) -> None:
    harness = pytest.importorskip('pyvbaharness')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''
Public Sub Make(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet, ch As Chart
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    Set ch = ws.ChartObjects.Add(48, 96, 240, 144).Chart
    ch.SetSourceData ws.Range("A1:C2")
    ch.Location xlLocationAsNewSheet, "Mine"
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Sub
''', name='P')
    path = tmp_path / 'made.xlsx'
    app.run('P.Make', str(path))
    with harness.ExcelSession() as excel:
        done = excel.run_vba(READ_BACK, 'Probe', (str(path),), timeout=120.0, line_numbers=False)
    assert done.value == 'Mine Chart One 1 =SERIES(One!$A$2,One!$B$1:$C$1,One!$B$2:$C$2,1)'
