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
#: What the model reports rather than answers: a scroll bar's width, and a list box's height.
REFUSED = re.compile(r'^ScrollBar: s\.Width|^ListBox: s\.Height')


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
    wrong = []
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

