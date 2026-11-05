"""Wrapped text and its row's height, replayed through the model from live Excel at 96 DPI: where a line stops wrapping (scripts/measure_wrapped_heldout.py), how tall text makes a row at each column width
(scripts/measure_wrapped_rows.py, scripts/measure_wrapped_curves.py, scripts/measure_wrapped_round3.py), and when Excel works the height out (scripts/measure_wrapped_events.py)."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.apps.excel._dimensions import row_points_read
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'


def _load(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))


EVENTS = _load('wrapped_events')
ROWS = _load('wrapped_rows')
HELDOUT = _load('wrapped_heldout_predictions')
HELDOUT_ANSWERS = _load('wrapped_heldout')
CURVES = _load('wrapped_curves_predictions')
CURVE_ANSWERS = _load('wrapped_curves')
ROUND3 = _load('wrapped_round3_predictions')
ROUND3_ANSWERS = _load('wrapped_round3')
#: The model has widths for these alone, and refuses the rest.
TABLED = ("Aptos Narrow", "Calibri", "Arial", "Times New Roman")


def _sheet(text: str, font: str, size: float, *, bold: bool = False, italic: bool = False, indent: int = 0) -> tuple[ExcelApplication, Any]:
    app = ExcelApplication()
    app.add_workbook()
    source = ('Public Sub Make(ByVal text As String, ByVal font As String, ByVal size As Double, ByVal bold As Boolean, ByVal italic As Boolean, ByVal indent As Long)\n    With ActiveSheet.Range("A1")\n'
              '        .Value = "\'" & text\n        .WrapText = True\n        .Font.Name = font\n        .Font.Size = size\n        .Font.Bold = bold\n        .Font.Italic = italic\n        .IndentLevel = indent\n    End With\nEnd Sub\n')
    app.add_module(source, name='Probe')
    app.run('Probe.Make', text, font, size, bold, italic, indent)
    return app, app.workbook.sheets_[0]


def _height_at(sheet: Any, column: int) -> float:
    sheet.dims.set_column_width(1, column)
    sheet.dims.autofit_row(1)
    return row_points_read(sheet.dims.row_quarters(1))


@pytest.mark.parametrize('case', [one for one in ROWS['cases'] if one['font'] in TABLED], ids=lambda one: f"{one['font']}-{one['size']}-{one['text'][:18]}")
def test_each_curve_as_excel_answered(case: dict[str, Any]) -> None:
    _, sheet = _sheet(case['text'], case['font'], case['size'])
    wrong = [(column, tall, got) for column, tall in case['heights'] if (got := _height_at(sheet, column)) != tall]
    assert not wrong


@pytest.mark.parametrize('number', range(len(HELDOUT['cases'])), ids=lambda number: f"{HELDOUT['cases'][number]['font']}-{HELDOUT['cases'][number]['size']}-{HELDOUT['cases'][number]['gdi']}")
def test_each_held_out_line_stops_wrapping_where_excel_says(number: int) -> None:
    case, answer = HELDOUT['cases'][number], HELDOUT_ANSWERS['answers'][number]
    _, sheet = _sheet(case['text'], case['font'], case['size'])
    if case['font'] not in TABLED:
        with pytest.raises(VBAUnsupportedError):
            _height_at(sheet, answer['over'])
        return
    assert _height_at(sheet, answer['over']) == answer['one_line']
    assert _height_at(sheet, answer['under']) > answer['one_line']


@pytest.mark.parametrize('number', range(len(CURVES['cases'])), ids=lambda number: f"{CURVES['cases'][number]['font']}-{CURVES['cases'][number]['size']}-{CURVES['cases'][number]['text'][:14]}")
def test_each_held_out_curve_as_excel_answered(number: int) -> None:
    case = CURVES['cases'][number]
    _, sheet = _sheet(case['text'], case['font'], case['size'], bold=case['bold'])
    wrong = [(column, tall, got) for column, tall in CURVE_ANSWERS['curves'][number].items() if (got := _height_at(sheet, int(column))) != tall]
    assert not wrong


@pytest.mark.parametrize('number', range(len(ROUND3['cases'])), ids=lambda number: f"{ROUND3['cases'][number]['font']}-{ROUND3['cases'][number]['size']}-{number}")
def test_each_third_round_curve_as_excel_answered(number: int) -> None:
    case = ROUND3['cases'][number]
    _, sheet = _made(case)
    wrong = [(column, tall, got) for column, tall in ROUND3_ANSWERS['curves'][number].items() if (got := _height_at(sheet, int(column))) != tall]
    assert not wrong


SAVES = _load('wrapped_model_saves')
ROUND4 = _load('wrapped_round4_predictions')
ROUND4_ANSWERS = _load('wrapped_round4')
_HYPHEN_THEN_LINE_FEED = '-\n'


@pytest.mark.parametrize('number', range(len(ROUND4['cases'])), ids=lambda number: f"{ROUND4['cases'][number]['font']}-{ROUND4['cases'][number]['size']}-{number}")
def test_each_fourth_round_curve_as_excel_answered(number: int) -> None:
    case = ROUND4['cases'][number]
    _, sheet = _made(case)
    if case['text'].endswith(_HYPHEN_THEN_LINE_FEED):
        with pytest.raises(VBAUnsupportedError):
            _height_at(sheet, 100)
        return
    wrong = [(column, tall, got) for column, tall in ROUND4_ANSWERS['curves'][number].items() if (got := _height_at(sheet, int(column))) != tall]
    assert not wrong


def _made(case: dict[str, Any]) -> tuple[ExcelApplication, Any]:
    """A text with line feeds in it, which the plain setter takes through the sheet's cell."""
    app, sheet = _sheet('x', case['font'], case['size'], bold=case['bold'], italic=case['italic'], indent=case['indent'])
    app.add_module('Public Sub Text(ByVal text As String)\n    ActiveSheet.Range("A1").Value = "\'" & Replace(text, "~", vbLf)\nEnd Sub\n', name='Texter')
    app.run('Texter.Text', case['text'].replace('\n', '~'))
    return app, sheet


_READ = '    r = r & "|" & ws.Rows(1).RowHeight & "/" & ws.Rows(2).RowHeight'


def _event_source(statements: list[str]) -> str:
    lines = ['Public Function Go() As String', '    Dim ws As Object, c As Object, d As Object, r As String', '    Set ws = ActiveSheet', '    ws.Cells.Clear', '    ws.Rows("1:3").RowHeight = 15', '    ws.Rows("1:3").UseStandardHeight = True', '    ws.Columns("A:B").ColumnWidth = 8.43', '    Set c = ws.Range("A1")', '    Set d = ws.Range("B1")', _READ]
    for statement in statements:
        lines += [f'    {statement}', _READ]
    return '\n'.join([*lines, '    Go = r', 'End Function']) + '\n'


#: What the model reports rather than answers, and where in each sequence: a carriage return's width is not measured, nor AutoFit of a column over wrapped text.
REFUSED_STATEMENTS = ('c.Value = "ab" & vbCr & "cd"', 'ws.Columns(1).AutoFit')


@pytest.mark.parametrize('name', EVENTS['events'])
def test_each_event_works_the_height_out_as_excel_does(name: str) -> None:
    statements: list[str] = EVENTS['events'][name]
    # The first reading comes before any statement; each statement has the reading after it.
    kept = [one for one in statements if one not in REFUSED_STATEMENTS]
    answers = [EVENTS['rows'][name][0], *[found for one, found in zip(statements, EVENTS['rows'][name][1:], strict=True) if one not in REFUSED_STATEMENTS]]
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(_event_source(kept), name='Probe')
    assert str(app.run('Probe.Go')).split('|')[1:] == answers


@pytest.mark.parametrize('statement', REFUSED_STATEMENTS)
def test_what_is_not_measured_reports_itself(statement: str) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Go()\n    Dim ws As Object, c As Object\n    Set ws = ActiveSheet\n    Set c = ws.Range("A1")\n    c.WrapText = True\n    c.Value = "one two"\n    {statement}\n    x = ws.Rows(1).RowHeight\nEnd Sub\n', name='Probe')
    with pytest.raises(VBAUnsupportedError):
        app.run('Probe.Go')


HISTORY: dict[str, dict[str, str]] = CURVE_ANSWERS['history']
LONG = 'one two three four five six seven'
#: scripts/measure_wrapped_curves.py's cases: what happens before the save, and what after it, with @.
HISTORY_CASES = {'stands': [], 'wider then saved': ['ws.Columns(1).ColumnWidth = 40'], 'fixed': ['ws.Rows(1).RowHeight = 40'], 'autofit after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Rows(1).AutoFit'],
                 'value after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("A1").Value = "one two"'], 'other cell after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("B1").Value = "x"'], 'font after opening': ['ws.Columns(1).ColumnWidth = 40', '@ws.Range("A1").Font.Bold = True'],
                 'width after opening': ['@ws.Columns(1).ColumnWidth = 40'], 'fixed then value after opening': ['ws.Rows(1).RowHeight = 40', '@ws.Range("A1").Value = "one two"']}


def _row_tag(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        found = re.search(r'<row r="1"[^>]*>', package.read('xl/worksheets/sheet1.xml').decode('utf-8'))
    return found.group() if found else ''


@pytest.mark.parametrize('name', HISTORY_CASES)
def test_a_saved_height_carries_as_in_excel(name: str, tmp_path: Path) -> None:
    before = [one for one in HISTORY_CASES[name] if not one.startswith('@')]
    after = [one[1:] for one in HISTORY_CASES[name] if one.startswith('@')]
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    source = ['Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, r As String', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1").WrapText = True', f'    ws.Range("A1").Value = "{LONG}"', *[f'    {one}' for one in before], '    r = ws.Rows(1).RowHeight', f'    wb.SaveAs folder & "{os.sep}saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              f'    Set wb = Workbooks.Open(folder & "{os.sep}saved.xlsx")', '    Set ws = wb.Worksheets(1)', '    r = r & "|" & ws.Rows(1).RowHeight', *[f'    {one}' for one in after], '    r = r & "|" & ws.Rows(1).RowHeight', '    wb.Close False', '    Go = r', 'End Function']
    app.add_module('\n'.join(source) + '\n', name='Probe')
    assert app.run('Probe.Go', str(tmp_path)) == HISTORY[name]['set, opened, after']
    assert _row_tag(tmp_path / 'saved.xlsx') == HISTORY[name]['row']


@pytest.mark.parametrize('name', SAVES['cases'])
def test_what_excel_read_of_the_model_s_saves_still_holds(name: str) -> None:
    """scripts/check_wrapped_model_saves.py: Excel opened each file the model saved, and read what the model did, and Excel's own save of the same steps has the same rows."""
    case = SAVES['cases'][name]
    assert case['model read'] == case['excel read'] == case['excel opening the model']
    assert case['model rows'] == case['excel rows']


def test_a_wrapped_row_with_no_height_in_its_file_is_the_standard_row(tmp_path: Path) -> None:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(f'Public Sub Make()\n    ActiveSheet.Range("A1").WrapText = True\n    ActiveSheet.Range("A1").Value = "{LONG}"\nEnd Sub\n', name='Probe')
    app.run('Probe.Make')
    app.save(tmp_path / 'saved.xlsx')
    with zipfile.ZipFile(tmp_path / 'saved.xlsx') as package:
        parts = {info.filename: package.read(info.filename) for info in package.infolist()}
    parts['xl/worksheets/sheet1.xml'] = re.sub(r'(<row r="1"[^>]*?) ht="[^"]*"', r'\1', parts['xl/worksheets/sheet1.xml'].decode('utf-8')).encode('utf-8')
    with zipfile.ZipFile(tmp_path / 'bare.xlsx', 'w') as package:
        for name, data in parts.items():
            package.writestr(name, data)
    opened = ExcelApplication.open(tmp_path / 'bare.xlsx', with_vba=False)
    assert row_points_read(opened.workbook.sheets_[0].dims.row_quarters(1)) == 15
