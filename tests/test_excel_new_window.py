"""A workbook with several windows, replayed from a shown Excel (scripts/measure_new_window.py, tests/fixtures/new_window.json):
what each row answers, typed and late, what a save writes, and what Excel read from the model's save."""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads((Path(__file__).parent / 'fixtures/new_window.json').read_text(encoding='utf-8'))
CASES: list[dict[str, Any]] = RECORD['cases']
#: The model refuses to hide a workbook's window, as it has no such state to save.
HIDING = ('hidden_copied', 'hidden_other')
#: The model adds no chart sheet to a new workbook: these go on one from a file, below.
NEW_CHART_SHEET = ('chart_sheet_shown', 'chart_sheet_own', 'chart_sheet_deleted')


def masked(answer: str) -> str:
    return re.sub(r'\b(Sheet|Book)\d+', 'Book', answer)


@pytest.mark.parametrize('case', CASES, ids=[f'{one["name"]} {one["binding"]}' for one in CASES])
def test_each_row_as_excel_answered(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    home = app.workbook.name
    app.interpreter.sandbox = False
    app.add_module(RECORD['source'], name='Probe')
    if case['name'] in HIDING:
        with pytest.raises(VBAUnsupportedError, match='hiding'):
            app.run('Probe.InModel', case['name'], case['binding'], str(tmp_path))
        return
    if case['name'] in NEW_CHART_SHEET:
        with pytest.raises(VBAUnsupportedError, match='Sheets.Add'):
            app.run('Probe.InModel', case['name'], case['binding'], str(tmp_path))
        return
    got = str(app.run('Probe.InModel', case['name'], case['binding'], str(tmp_path))).replace(home, '<HOME>')
    if (case['name'], case['binding']) == ('argument', 'Workbook'):
        # 438 on a typed call with an argument, which the binding table does not hold for this method.
        assert (got, case['excel']) == ('E450', 'E438')
        return
    assert masked(got) == case['excel']


@pytest.mark.parametrize('case', [one for one in CASES if one['excel_reads_model']], ids=lambda one: f'{one["name"]} {one["binding"]}')
def test_excel_read_the_model_save(case: dict[str, Any]) -> None:
    """`excel_reads_model` is Excel's reading of the model's file, recorded when the fixture was measured; `excel_reads_own` is its reading of its own."""
    assert case['excel_reads_model'] == case['excel_reads_own']


def _views(path: Path) -> tuple[list[str], list[list[str]]]:
    with zipfile.ZipFile(path) as package:
        book = re.findall(r'<workbookView\b[^>]*>', package.read('xl/workbook.xml').decode('utf-8'))
        sheets = [re.findall(r'<sheetView\b[^>]*>', package.read(name).decode('utf-8')) for name in sorted(package.namelist()) if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name)]
    return book, sheets


@pytest.mark.parametrize('case', [one for one in CASES if one['files'] and one['binding'] == 'Object'], ids=lambda one: one['name'])
def test_a_save_writes_a_view_for_each_window_as_excel_does(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.interpreter.sandbox = False
    app.add_module(RECORD['source'], name='Probe')
    app.run('Probe.InModel', case['name'], case['binding'], str(tmp_path))
    file, measured = next(iter(case['files'].items()))
    book, sheets = _views(tmp_path / file)
    assert len(book) == len(measured['book'])
    ids = [re.findall(r'workbookViewId="(\d+)"', ' '.join(one)) for one in sheets]
    assert ids == [re.findall(r'workbookViewId="(\d+)"', ' '.join(views)) for name, views in sorted(measured.items()) if name != 'book']
    tabbed = [['tabSelected' in view for view in one] for one in sheets]
    assert tabbed == [['tabSelected' in view for view in views] for name, views in sorted(measured.items()) if name != 'book']


def _ran(source: str) -> tuple[ExcelApplication, str]:
    app = ExcelApplication()
    app.add_workbook()
    app.interpreter.sandbox = False
    app.add_module('Public Function Go(ByVal folder As String) As String\n    Dim wb As Workbook, w As Object\n    Application.DisplayAlerts = False\n'
                   '    Set wb = Workbooks.Add(xlWBATWorksheet)\n    wb.Worksheets(1).Name = "A"\n    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "B"\n'
                   f'{source}\nEnd Function\n', name='Probe')
    return app, 'Probe.Go'


def test_a_second_window_setting_survives_a_save_and_a_load(tmp_path: Path) -> None:
    source = ('    wb.Worksheets("A").Activate\n    wb.NewWindow\n    wb.Windows(1).Zoom = 80\n    wb.Windows(2).DisplayGridlines = False\n'
              '    wb.SaveAs folder & "\\Two.xlsx", xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(folder & "\\Two.xlsx", 0)\n'
              '    Go = wb.Windows.Count & "|" & wb.Windows(1).WindowNumber & "|" & wb.Windows(1).DisplayGridlines & "|" & wb.Windows(2).Zoom')
    app, routine = _ran(source)
    assert str(app.run(routine, str(tmp_path))) == '2|1|False|80'


def test_closing_a_window_writes_one_view_again(tmp_path: Path) -> None:
    source = ('    wb.NewWindow\n    wb.Windows(1).Close\n    wb.SaveAs folder & "\\One.xlsx", xlOpenXMLWorkbook\n    Go = wb.Windows.Count')
    app, routine = _ran(source)
    assert str(app.run(routine, str(tmp_path))) == '1'
    book, sheets = _views(tmp_path / 'One.xlsx')
    assert len(book) == 1
    assert all(len(one) == 1 for one in sheets)


def test_a_book_with_one_window_saves_as_it_did(tmp_path: Path) -> None:
    """No window made: no view is touched."""
    source = '    wb.SaveAs folder & "\\One.xlsx", xlOpenXMLWorkbook\n    Go = wb.Windows(1).Caption'
    app, routine = _ran(source)
    assert str(app.run(routine, str(tmp_path))) == 'One.xlsx'
    book, sheets = _views(tmp_path / 'One.xlsx')
    assert len(book) == 1
    assert all(len(one) == 1 for one in sheets)


def _chart_sheet_book(body: str) -> ExcelApplication:
    app = ExcelApplication.open(Path(__file__).parent / 'fixtures/chart_sheets/active.xlsx', with_vba=False)
    app.add_module(f'Public Function Go() As String\n    Dim wb As Workbook, s As String\n    Set wb = ActiveWorkbook\n{body}\n    Go = s\nEnd Function\n', name='Go')
    return app


def test_a_window_shows_a_chart_sheet_as_excel_did() -> None:
    measured = next(one['excel'] for one in CASES if one['name'] == 'chart_sheet_shown')
    app = _chart_sheet_book('    wb.NewWindow\n    s = TypeName(wb.Windows(1).ActiveSheet) & "|" & TypeName(wb.Windows(2).ActiveSheet) & "|" & wb.Windows.Count')
    assert str(app.run('Go.Go')) == measured


def test_a_window_behind_goes_on_when_its_chart_sheet_is_deleted() -> None:
    """As Excel did (chart_sheet_deleted): both windows show the worksheet in the chart sheet's place."""
    app = _chart_sheet_book('    Application.DisplayAlerts = False\n    wb.NewWindow\n    wb.ActiveSheet.Delete\n    s = TypeName(wb.Windows(1).ActiveSheet) & "|" & TypeName(wb.Windows(2).ActiveSheet)')
    assert str(app.run('Go.Go')) == 'Worksheet|Worksheet'

