"""Shape settings across a save and a reopen, replayed from live Excel (scripts/measure_shape_settings_kept.py): what each reads before and after,
and the anchor the file holds."""
from __future__ import annotations

import json
import os
import re
import shutil
import struct
import zipfile
import zlib
from pathlib import Path
from typing import Any

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).parent / 'fixtures'
RECORD: dict[str, Any] = json.loads((FIXTURES / 'shape_settings_kept.json').read_text(encoding='utf-8'))
NEW: dict[str, str] = RECORD['new']
CONTROLS = ('CheckBox', 'Button', 'DropDown', 'Spinner')
#: What the model reports rather than answers: a picture's old object, which is only partly measured.
REFUSED = re.compile(r'^(Picture|Pic): s\.DrawingObject\.(Interior\.(Color|ColorIndex)|Border\.)')


def _png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * 4 for _ in range(4))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 4, 4, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')


def _reader() -> str:
    lines = ['Private Function Shown(v As Variant) As String', '    If IsObject(v) Then', '        Shown = "object"', '    ElseIf IsArray(v) Or IsNull(v) Then', '        Shown = TypeName(v)', '    Else', '        Shown = CStr(v)', '    End If', 'End Function', '',
             'Private Function Reads(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next']
    for read in RECORD['reads']:
        lines += ['    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E" & Err.Number', '    Else', '        r = r & "|" & Shown(v)', '    End If']
    return '\n'.join([*lines, '    Reads = r', 'End Function', ''])


def _source(target: str, statement: str) -> str:
    made = target in NEW
    lines = [_reader(), 'Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, s As Object, r As String, name As String', '    Application.DisplayAlerts = False']
    lines += ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    {NEW[target]}', '    s.Name = "Target"'] if made else [
        '    Set wb = Workbooks.Open(folder & "\\book.xlsx")', '    Set ws = wb.Worksheets(1)', f'    Set s = ws.Shapes("{target}")']
    lines += ['    name = s.Name', '    On Error Resume Next', f'    {statement}' if statement else "    ' Nothing is set.", '    If Err.Number Then r = "E" & Err.Number & "!"', '    Err.Clear', '    On Error GoTo 0', '    r = r & Reads(s) & "@@"',
              '    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Set wb = Workbooks.Open(folder & "\\saved.xlsx")', '    r = r & Reads(wb.Worksheets(1).Shapes(name))', '    wb.Close False', '    Application.DisplayAlerts = True', '    Go = r', 'End Function']
    return '\n'.join(lines).replace('folder & "\\', f'folder & "{os.sep}')+ '\n'


def _anchor(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    for found in re.finditer(r'<xdr:(?:twoCell|oneCell|absolute)Anchor\b.*?</xdr:(?:twoCell|oneCell|absolute)Anchor>', xml, re.S):
        if f'name="{name}"' in found.group():
            return found.group()
    return ''


def _without_ids(markup: str) -> str:
    """A new shape's creation id is its own each time."""
    return re.sub(r'(creationId [^>]*id=)"[^"]*"', r'\1""', markup)


_CASES = [one for one in RECORD['cases'] if one['target'] not in CONTROLS]


@pytest.mark.parametrize('case', _CASES, ids=[f"{one['target']}-{one['statement']}" for one in _CASES])
def test_each_setting_as_excel_answered_and_saved(case: dict[str, Any], tmp_path: Path) -> None:
    (tmp_path / 'p.png').write_bytes(_png())
    shutil.copyfile(FIXTURES / 'loaded_shape_text.xlsx', tmp_path / 'book.xlsx')
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(_source(case['target'], case['statement']), name='Probe')
    label = f"{case['target']}: {case['statement']}"
    if REFUSED.search(label):
        with pytest.raises(VBAUnsupportedError):
            app.run('Probe.Go', str(tmp_path))
        return
    assert app.run('Probe.Go', str(tmp_path)) == case['answer']
    assert _without_ids(_anchor(tmp_path / 'saved.xlsx', 'Target' if case['target'] in NEW else case['target'])) == _without_ids(case['anchor'])
