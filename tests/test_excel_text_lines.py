"""TextRange2.Lines against Excel: the rounds written down before it was asked (scripts/predict_shape_lines.py, scripts/predict_shape_lines2.py), and the member (scripts/measure_shape_lines_members.py)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'


def _load(name: str) -> dict[str, Any]:
    return json.loads((FIXTURES / f'{name}.json').read_text('utf-8'))


FIRST, SECOND, MEMBERS, EVENTS = _load('shape_lines'), _load('shape_lines2'), _load('shape_lines_members'), _load('wrapped_events')
LAYOUT = 'LAYOUT'
#: What is not measured to leave the rule as it is.
REFUSED = ('.Rotation = 90', '.Rotation = 33', '.Orientation', '.Column.Number', '.LeftIndent', '.FirstLineIndent', '.RightIndent', '.IndentLevel', '.Bullet.Visible', '.Font.Spacing', '.Allcaps', '.Smallcaps', '.Superscript', 'Characters(1, 9).Font.Size', 'Characters(11, 5).Font',
           'Font.Name = "No Such Font"', '.Font.Name = "+m', 'vbTab', 'ChrW(8212)', 'ChrW(20013)', 'ChrW(160)', 'ChrW(173)', '.AutoShapeType', 'T.Characters(5, 0)', '.InsertBefore', '.InsertAfter " abc"', 'AutoSize = 1', 'AddLabel')
#: The same of shapes in a file: other shapes, an upright or set-upright text, and fonts that differ.
REFUSED_LOADED = ('Oval', 'Upward', 'Runs', 'Fonts', 'Paint', 'Paras', 'Broken', 'Clipped')


def _source(setup: str, read: str) -> str:
    return MEMBERS['source'].replace('%TEXT%', MEMBERS['text']).replace('%SETUP%', setup or "'").replace('%READ%', 'Layout(s)' if read == LAYOUT else read)


def _app(source: str) -> ExcelApplication:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source, name='Probe')
    return app


def _replay(data: dict[str, Any], case: dict[str, Any]) -> list[tuple[str, str, str, int, list[list[int]]]]:
    codes = ','.join(str(ord(character)) for character in case['text'])
    widths = ','.join(str(answer['pixel'] * 0.75) for answer in case['answers'])
    if 'kind' in case:
        arguments = (codes, case['font'], case['size'], case['bold'], case['italic'], case['kind'], case['margins'][0], case['margins'][1], case['wrap'], widths)
    else:
        arguments = (codes, case['font'], case['size'], case['bold'], case['italic'], widths)
    found = str(_app(data['source']).run('Probe.Go', *arguments)).split('|')
    assert found[0] == '15~15'
    rows: list[tuple[str, str, str, int, list[list[int]]]] = []
    for one in found[1:]:
        width, kind, count_kind, count, parts = one.split(':')
        rows.append((width, kind, count_kind, int(count), [[int(part.split('+')[0]) - 1, int(part.split('+')[1])] for part in parts.split(',') if part]))
    return rows


@pytest.mark.parametrize('data, case', [(data, case) for data in (FIRST, SECOND) for case in data['cases']], ids=lambda value: value['text'][:24] + '-' + value['font'] + str(value['size']) if 'text' in value else 'round')
def test_each_text_breaks_where_excel_broke_it(data: dict[str, Any], case: dict[str, Any]) -> None:
    assert _replay(data, case) == [(one['width'], one['type'], one['count_type'], one['count'], one['lines']) for one in case['answers']]


def test_the_rounds_were_measured_at_96_dpi() -> None:
    assert FIRST['standard_height'] == SECOND['standard_height'] == MEMBERS['standard_height'] == '15'


@pytest.mark.parametrize('text', EVENTS['shapes'])
def test_the_first_curves_still_hold(text: str) -> None:
    source = ('Public Function Go(ByVal text As String, ByVal widths As String) As String\n    Dim s As Object, r As String, one As Variant\n    Set s = ActiveSheet.Shapes.AddShape(1, 10, 10, 100, 300)\n'
              '    s.TextFrame2.TextRange.Text = text\n    For Each one In Split(widths, ",")\n        s.Width = Val(one)\n        r = r & "|" & s.TextFrame2.TextRange.Lines.Count\n    Next\n    Go = r\nEnd Function\n')
    found = str(_app(source).run('Probe.Go', text, ','.join(str(one) for one in EVENTS['shape_widths']))).split('|')[1:]
    assert found == EVENTS['shapes'][text]


@pytest.mark.parametrize('case', MEMBERS['cases'], ids=[f"{one['setup'][:60]}-{one['read'][-30:]}" for one in MEMBERS['cases']])
def test_each_member_as_excel_answered(case: dict[str, str]) -> None:
    app = _app(_source(case['setup'], case['read']))
    if any(mark.replace('T.', '') in (case['setup'] + ' ' + case['read']).replace('s.TextFrame2.TextRange', 'T') for mark in REFUSED):
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go')
        return
    assert str(app.run('Probe.Go')).split('~')[2:] == case['answer'].split('~')


@pytest.mark.parametrize('name', MEMBERS['loaded'])
def test_a_shape_in_a_file_breaks_as_excel_read_it(name: str) -> None:
    source = f'''Public Function Go(ByVal path As String) As String
    Dim s As Object, one As String, i As Long
    Set s = Workbooks.Open(path).Worksheets(1).Shapes("{name}")
    On Error Resume Next
    one = s.TextFrame2.TextRange.Lines.Count & ":"
    If Err.Number <> 0 Then
        one = "error " & Err.Number
    Else
        For i = 1 To s.TextFrame2.TextRange.Lines.Count
            one = one & s.TextFrame2.TextRange.Lines(i).Start & "+" & s.TextFrame2.TextRange.Lines(i).Length & ","
        Next
    End If
    Go = s.Width & "=" & one
End Function
'''
    app = _app(source)
    book = str(FIXTURES / 'loaded_shape_text.xlsx')
    if name in REFUSED_LOADED:
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go', book)
        return
    assert str(app.run('Probe.Go', book)).split('=') == MEMBERS['loaded'][name]
