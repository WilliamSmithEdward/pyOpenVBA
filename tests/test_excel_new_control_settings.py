"""A new form control through Shape, replayed from live Excel (scripts/measure_new_control_settings.py): what it reads from the first, what each setting reads and leaves in the file,
and what reads after a save and a reopen. All at 96 DPI, where StandardHeight reads 15, which is what the box is rounded for."""
from __future__ import annotations

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'new_control_settings.json').read_text(encoding='utf-8'))
READS: list[str] = RECORD['reads']
KINDS: dict[str, int] = RECORD['kinds']
BOXES: dict[str, Any] = json.loads((FIXTURES / 'new_control_boxes.json').read_text(encoding='utf-8'))
GRID: dict[str, Any] = json.loads((FIXTURES / 'new_control_grid.json').read_text(encoding='utf-8'))
GEOMETRY: set[str] = set()
#: What the model reports rather than answers: a scroll bar's size, and a list box's height.
REFUSED = re.compile(r'^ScrollBar: s\.(?:Width|Height)|^ListBox: s\.Height')


def _block(name: str, form: str) -> list[str]:
    lines = [f'Private Function {name}(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E"' + (' & Err.Number' if form == 'value' else ''), '    Else', '        r = r & "|" & ' + ('CStr(v)' if form == 'value' else 'TypeName(v)'), '    End If']
    return [*lines, f'    {name} = r', 'End Function', '']


def _source(case: dict[str, Any]) -> str:
    left, top, width, height = case['box']
    lines = [*_block('Reads', 'value'), *_block('TypesOf', 'type'), 'Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, s As Object, r As String, name As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f"    Set s = ws.Shapes.AddFormControl({KINDS[case['kind']]}, {left}, {top}, {width}, {height})", '    name = s.Name', '    On Error Resume Next',
             f"    {case['statement']}" if case['statement'] else "    ' Nothing is set.", '    If Err.Number Then r = "E" & Err.Number & "!"', '    Err.Clear', '    On Error GoTo 0', '    r = r & Reads(s) & "##" & TypesOf(s) & "@@"', f'    wb.SaveAs folder & "{os.sep}saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False',
             f'    Set wb = Workbooks.Open(folder & "{os.sep}saved.xlsx")', '    r = r & Reads(wb.Worksheets(1).Shapes(name)) & "##" & TypesOf(wb.Worksheets(1).Shapes(name))', '    wb.Close False', '    Application.DisplayAlerts = True', '    Go = r', 'End Function']
    return '\n'.join(lines) + '\n'


def _fields(answer: str) -> dict[str, str]:
    """An answer by what it says, without where the box lands."""
    first, second = answer.split('@@')
    error = ''
    if not first.startswith('|'):
        error, first = first.split('!', 1)
    found = {'error': error}
    for side, text in (('before', first), ('after', second)):
        values, types = text.split('##')
        for read, value, kind in zip(READS, values.split('|')[1:], types.split('|')[1:], strict=True):
            if read not in GEOMETRY:
                found[f'{side} {read}'] = f'{kind} {value}'
    return found


def _facts(drawing: str, vml: str, controls: str) -> dict[str, object]:
    """What the file says of these settings, wherever each part says it."""
    edit = re.search(r'<xdr:twoCellAnchor\b([^>]*)>', drawing)
    anchor = re.search(r'<anchor\b([^>]*)>', controls)
    head = re.search(r'<controlPr\b([^>]*)>', controls)
    alt = re.search(r'\balt="([^"]*)"', vml)
    alt_text = re.search(r'\baltText="([^"]*)"', controls)
    return {'editAs': edit.group(1).strip() if edit else None, 'anchor': anchor.group(1).strip() if anchor else None,
            'locked': 'locked="0"' in (head.group(1) if head else ''), 'altText': alt_text.group(1) if alt_text else None, 'alt': alt.group(1) if alt else None,
            'hidden': 'visibility:hidden' in vml, 'x:Locked': '<x:Locked>' in vml, 'x:MoveWithCells': '<x:MoveWithCells' in vml, 'x:SizeWithCells': '<x:SizeWithCells' in vml, 'aspect': 'aspectratio="t"' in vml}


def _file_facts(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        vml = next((package.read(one).decode('utf-8') for one in package.namelist() if one.endswith('.vml')), '')
        sheet = next(package.read(one).decode('utf-8') for one in package.namelist() if one.startswith('xl/worksheets/sheet') and one.endswith('.xml'))
        found = re.search(r'<controls>.*?</controls>', sheet, re.S)
        return _facts(package.read('xl/drawings/drawing1.xml').decode('utf-8'), vml, found.group() if found else '')


_CASES = [*RECORD['cases'], *BOXES['cases']]


@pytest.mark.parametrize('case', _CASES, ids=[f"{one['kind']}-{one['statement']}" for one in _CASES])
def test_each_setting_as_excel_answered_and_saved(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(case), name='Probe')
    if REFUSED.search(f"{case['kind']}: {case['statement']}"):
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go', str(tmp_path))
        return
    answer = str(app.run('Probe.Go', str(tmp_path)))
    assert _fields(answer) == _fields(case['answer'])
    assert _file_facts(tmp_path / 'saved.xlsx') == _facts(case['drawing'], case.get('vml', ''), case.get('controls', ''))


def test_the_fixtures_were_measured_at_96_dpi() -> None:
    assert RECORD['standard_height'] == BOXES['standard_height'] == '15'


def test_each_box_lands_where_excel_put_it() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(GRID_SOURCE, name='Probe')
    wrong: list[tuple[list[float], str, str]] = []
    for start in range(0, len(GRID['boxes']), 300):
        chunk = GRID['boxes'][start:start + 300]
        found = str(app.run('Probe.Go', ';'.join(','.join(str(one) for one in box) for box in chunk))).split('|')
        wrong += [(box, truth, got) for box, truth, got in zip(chunk, GRID['excel'][start:start + 300], found, strict=True) if truth != got]
    assert not wrong


GRID_SOURCE = '''Public Function Go(ByVal text As String) As String
    Dim ws As Object, s As Object, r As String, line As Variant, n As Variant
    Set ws = ActiveSheet
    For Each line In Split(text, ";")
        n = Split(line, ",")
        Set s = ws.Shapes.AddFormControl(Val(n(4)), Val(n(0)), Val(n(1)), Val(n(2)), Val(n(3)))
        r = r & "|" & s.Left & "," & s.Top & "," & s.Width & "," & s.Height
        s.Delete
    Next
    Go = Mid(r, 2)
End Function
'''


def test_a_setting_on_a_control_in_a_file_comes_back(tmp_path: Path) -> None:
    """Saved, opened, changed, saved and opened again."""
    source = 'Public Function Go() As String\n    Set s = ActiveWorkbook.Worksheets(1).Shapes("Button 1")\n    {0}\n    Go = s.AlternativeText & "," & s.Placement & "," & s.Locked & "," & s.Visible\nEnd Function\n'
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('Public Sub Make()\n    ActiveSheet.Shapes.AddFormControl 0, 10, 10, 100, 40\nEnd Sub\n', name='Probe')
    app.run('Probe.Make')
    app.save(tmp_path / 'first.xlsx')
    opened = ExcelApplication.open(tmp_path / 'first.xlsx', with_vba=False)
    opened.add_module(source.format('s.AlternativeText = "Alt": s.Placement = xlFreeFloating: s.Locked = False: s.Visible = False'), name='Probe')
    assert opened.run('Probe.Go') == 'Alt,3,False,0'
    opened.save(tmp_path / 'second.xlsx')
    again = ExcelApplication.open(tmp_path / 'second.xlsx', with_vba=False)
    again.add_module(source.format(''), name='Probe')
    assert again.run('Probe.Go') == 'Alt,3,False,0'


PLACES = json.loads((Path(__file__).parent / 'fixtures/new_control_places.json').read_text('utf-8'))
_CONSTANTS = {'xlButtonControl': 0, 'xlCheckBox': 1, 'xlDropDown': 2, 'xlListBox': 6, 'xlGroupBox': 4, 'xlLabel': 5, 'xlOptionButton': 7, 'xlScrollBar': 8, 'xlSpinner': 9}


@pytest.mark.parametrize('case', PLACES['cases'], ids=[f"{one['kind']}-{one['member']}-{one['value']}" for one in PLACES['cases']])
def test_a_place_or_size_set_lands_where_excel_put_it(case: dict[str, Any], tmp_path: Path) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(PLACES['source'], name='Probe')
    arguments = (_CONSTANTS[case['kind']], case['member'], case['value'], str(tmp_path / 'book.xlsx'))
    if (case['kind'], case['member']) in (('xlScrollBar', 'Width'), ('xlScrollBar', 'Height'), ('xlListBox', 'Height')):
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go', *arguments)
        return
    assert str(app.run('Probe.Go', *arguments)) == f"15~15~{case['answer']}"


OLD_PLACES = json.loads((Path(__file__).parent / 'fixtures/new_control_old_places.json').read_text('utf-8'))
#: A height set through the old object moves the Shape's Top by an EMU or not at all, by a rounding not worked out over 989 heights
#: (tests/fixtures/new_control_old_height_heldout.json), so the model's Shape.Top reports itself there.


@pytest.mark.parametrize('case', OLD_PLACES['cases'], ids=[f"{one['kind']}-{one['set']}" for one in OLD_PLACES['cases']])
def test_an_old_object_reads_its_box_as_excel_does(case: dict[str, str]) -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(OLD_PLACES['source'].replace('%MAKE%', OLD_PLACES['makers'][case['kind']]).replace('%SET%', case['set']), name='Probe')
    member = case['set'][2:].split(' = ')[0]
    if case['kind'] in ('group box', 'label', 'spinner', 'scroll bar') or (case['kind'] == 'list box' and member == 'Height') or (case['set'].startswith('o.Height') and case['kind'] not in ('rectangle', 'text box')):  # Their old objects are not built, and a list box's height goes its own way.
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go')
        return
    assert str(app.run('Probe.Go')).split('~', 2)[2] == case['answer']


OLD_HEIGHTS = json.loads((Path(__file__).parent / 'fixtures/new_control_old_heights.json').read_text('utf-8'))
_OLD_READS = '''Public Function Go(ByVal kind As Long, ByVal top As Double, ByVal tall As Double, ByVal through As String, ByVal path As String) As String
    Dim wb As Workbook, s As Object, o As Object, r As String
    Set wb = Workbooks.Add
    Set s = wb.Worksheets(1).Shapes.AddFormControl(kind, 12, top, 100, 40)
    Set o = wb.Worksheets(1).DrawingObjects(1)
    If through = "o" Then o.Height = tall Else s.Height = tall
    r = CStr(CDbl(s.Height) * 12700) & "," & o.Top & "," & o.Height
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(path)
    Set s = wb.Worksheets(1).Shapes(1)
    Go = r & "|" & CStr(CDbl(s.Top) * 12700) & "," & CStr(CDbl(s.Height) * 12700)
End Function
'''


@pytest.mark.parametrize('case', OLD_HEIGHTS['cases'], ids=[f"{one['kind']}-{one['through']}-{one['top']}-{one['height']}" for one in OLD_HEIGHTS['cases']])
def test_a_height_set_either_way_reads_and_saves_as_excel_does(case: dict[str, Any], tmp_path: Path) -> None:
    """Everything tests/fixtures/new_control_old_heights.json read but the Shape's Top straight after an old object's height, which reports itself."""
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_OLD_READS, name='Probe')
    _, after, saved = case['answer'].split('|')
    expected = f"{after.split(',', 1)[1]}|{saved}"
    assert str(app.run('Probe.Go', 0 if case['kind'] == 'button' else 1, case['top'], case['height'], case['through'], str(tmp_path / 'book.xlsx'))) == expected


def test_the_shape_s_top_after_an_old_object_s_height_reports_itself_until_it_is_set() -> None:
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module('''Public Function Go(ByVal again As Boolean) As String
    Dim s As Object
    Set s = ActiveSheet.Shapes.AddFormControl(xlButtonControl, 12, 12, 100, 40)
    ActiveSheet.DrawingObjects(1).Height = 46
    If again Then s.Top = 12
    Go = s.Top
End Function
''', name='Probe')
    with pytest.raises(VBAUnsupportedError, match='EMU'):
        app.run('Probe.Go', False)
    assert str(app.run('Probe.Go', True)) == '12'
