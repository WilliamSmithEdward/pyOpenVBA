"""Column AutoFit, measured by scripts/measure_column_autofit.py.

AutoFit makes a column as wide as the widest text its cells show: each
character as wide as GDI makes it in the cell's font, a number in General
a digit wide for each digit, point and exponent sign, italic reaching a
little further, and the text's width made the column's by a fixed rule up
to 255 characters. Hidden rows count, merged cells and empty text do not,
a hidden column is shown, Range.Columns.AutoFit fits the range's cells
alone, and a file keeps the width with bestFit="1". What AutoFit was not
measured over reports itself unsupported.
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

RECORD: dict[str, Any] = json.loads(
    (Path(__file__).parent / 'fixtures/column_autofit.json').read_text(encoding='utf-8'))
PROBES: dict[str, dict[str, Any]] = {probe['name']: probe for probe in RECORD['probes']}
NORMALS = {'aptos_narrow': 'Aptos Narrow', 'calibri': 'Calibri'}
#: Probes replayed as they ran in Excel.
WHOLE = [name for name in PROBES if name.startswith(('display_', 'runs_', 'latin_', 'ranges_', 'states_'))]
#: The cells cases AutoFit was not measured over, which report themselves unsupported.
UNMEASURED_CELLS = {'wrapped', 'line_break', 'indent', 'turned', 'big_font', 'mixed_runs', 'wrapped_beside',
                    'line_break_unwrapped', 'line_break_wraps'}


def _run(source: str) -> object:
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source, name='Harness')
    return app.run('Harness.Probe')


def _module(normal: str, body: str) -> str:
    """A probe as the measurement wrote them: a workbook of its own whose Normal font is ``normal``."""
    return ('Function Fitted(ws As Worksheet, column As Long) As Long\nws.Columns(column).AutoFit\n'
            'Fitted = CLng(ws.Columns(column).Width * 4 / 3)\nEnd Function\n\n'
            'Function Probe() As String\nDim wb As Workbook, ws As Worksheet, s As String, i As Long, items As Variant\n'
            f'Application.DisplayAlerts = False\nSet wb = Workbooks.Add(xlWBATWorksheet)\n'
            f'wb.Styles("Normal").Font.Name = "{normal}"\nwb.Styles("Normal").Font.Size = 11\n'
            f'Set ws = wb.Worksheets(1)\n{body}\nwb.Close False\nApplication.DisplayAlerts = True\n'
            'Probe = Mid$(s, 2)\nEnd Function\n')


def _vba(text: str) -> str:
    """A VBA expression for ``text``, ChrW for anything past ASCII and the quote."""
    pieces = [f'"{char}"' if 32 <= ord(char) < 127 and char != '"' else f'ChrW({ord(char)})' for char in text]
    return ' & '.join(pieces)


def _latin(text: str) -> bool:
    return all(ord(char) <= 0xFF for char in text)


@pytest.mark.parametrize('name', WHOLE)
def test_autofit_as_excel_did(name: str) -> None:
    probe = PROBES[name]
    assert _run(probe['macro']) == probe['result']


@pytest.mark.parametrize('name', [name for name in PROBES if name.startswith('glyphs_')])
def test_each_character_alone_and_beside_another(name: str) -> None:
    # Windows-1252's 27 characters past Latin-1 are left out of the replay: GDI's widths for them depend on their
    # neighbours, and AutoFit over them reports itself unsupported (test_characters_past_latin_1).
    macro = re.sub(r'items = Array\([^)]*\)\nFor i = LBound\(items\)[^\n]*\n', '', PROBES[name]['macro'])
    kept = ' '.join(piece for piece in PROBES[name]['result'].split()
                    if not piece.split(':')[0].isdigit() or int(piece.split(':')[0]) <= 0xFF)
    assert _run(macro) == kept


@pytest.mark.parametrize('normal', list(NORMALS))
def test_texts_of_mixed_characters(normal: str) -> None:
    measured = PROBES[f'mixed_{normal}']['result'].split()
    texts = [(text, width) for text, width in zip(RECORD['strings'], measured) if _latin(text)]
    body = '\n'.join(f'ws.Range("A1").Value = "\'" & {_vba(text)}: s = s & " " & Fitted(ws, 1)' for text, _ in texts)
    assert _run(_module(NORMALS[normal], body)) == ' '.join(width for _, width in texts)


def test_characters_past_latin_1() -> None:
    with pytest.raises(VBAUnsupportedError, match='past Latin-1'):
        _run(_module('Aptos Narrow', 'ws.Range("A1").Value = "don" & ChrW(8217) & "t": s = s & " " & Fitted(ws, 1)'))


def _numbers(cases: list[list[str]], *, bold: bool) -> str:
    return f'ws.Range("A1").Font.Bold = {bold}\n' + '\n'.join(
        f'ws.Columns(1).ColumnWidth = 0.83: ws.Range("A1").NumberFormat = "{code.replace(chr(34), chr(34) * 2)}": '
        f'ws.Range("A1").Formula = "{formula.replace(chr(34), chr(34) * 2)}": '
        f's = s & Chr(9) & Fitted(ws, 1) & "~" & ws.Range("A1").Text' for formula, code in cases)


@pytest.mark.parametrize('normal', list(NORMALS))
@pytest.mark.parametrize('kind', ['numbers', 'bold_numbers'])
def test_numbers_in_their_formats(normal: str, kind: str) -> None:
    measured = PROBES[f'{kind}_{normal}']['result'].split('\t')
    cases = [(case, found) for case, found in zip(RECORD[kind], measured) if '*' not in case[1]]
    got = _run(_module(NORMALS[normal], _numbers([case for case, _ in cases], bold=kind == 'bold_numbers')))
    assert str(got).split('\t') == [found for _, found in cases]


def test_a_number_format_with_a_fill() -> None:
    with pytest.raises(VBAUnsupportedError, match='fill'):
        _run(_module('Aptos Narrow', _numbers([['=1234.5', '_($* #,##0.00_)']], bold=False)))


def _cell(name: str, setup: str) -> str:
    return (f'If ws.AutoFilterMode Then ws.AutoFilterMode = False\nws.Cells.UnMerge: ws.Cells.Clear: '
            f'ws.Columns(1).ColumnWidth = 8.43: ws.Rows("1:3").Hidden = False\n{setup}\n'
            f's = s & "#{name}=" & Fitted(ws, 1)')


@pytest.mark.parametrize('normal', list(NORMALS))
@pytest.mark.parametrize('name', list(RECORD['cells']))
def test_what_autofit_looks_at(normal: str, name: str) -> None:
    measured = dict(piece.split('=', 1) for piece in PROBES[f'cells_{normal}']['result'].split('#'))
    source = _module(NORMALS[normal], _cell(name, RECORD['cells'][name]))
    if name in UNMEASURED_CELLS:
        with pytest.raises(VBAUnsupportedError):
            _run(source)
        return
    assert _run(source) == f'{name}={measured[name]}'


def test_another_size() -> None:
    with pytest.raises(VBAUnsupportedError, match='11pt'):
        _run(_module('Aptos Narrow', 'ws.Range("A1").Value = "\'abc": ws.Range("A1").Font.Size = 14: '
                                     's = s & " " & Fitted(ws, 1)'))


def test_the_columns_a_file_keeps(tmp_path: Path) -> None:
    path = tmp_path / 'fitted.xlsx'
    _run(RECORD['file']['macro'].replace('{path}', str(path)))
    with zipfile.ZipFile(path) as package:
        sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
    found = re.search(r'<cols>.*?</cols>', sheet)
    assert found is not None and found.group(0) == RECORD['file']['cols']
