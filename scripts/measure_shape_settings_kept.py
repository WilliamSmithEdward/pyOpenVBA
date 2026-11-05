"""Shape settings a set, save and reopen found the file losing in the model: what Excel reads after each, on a new shape and on one in a file.

    python scripts/measure_shape_settings_kept.py

Each case makes or opens the shape, sets one thing and reads it, saves, closes, opens what it saved and reads it again. It uses no hyperlink, mail or send member.

Writes tests/fixtures/shape_settings_kept.json. tests/test_excel_shape_settings_kept.py replays it.
"""
from __future__ import annotations

import json
import re
import shutil
import struct
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / 'tests/fixtures/loaded_shape_text.xlsx'
OUT = ROOT / 'tests/fixtures/shape_settings_kept.json'


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * 4 for _ in range(4))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 4, 4, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b'')


#: What makes each new shape, s.
NEW = {
    'Rect': 'Set s = ws.Shapes.AddShape(1, 100, 50, 120, 60)',
    'TextBox': 'Set s = ws.Shapes.AddTextbox(1, 100, 50, 120, 60)',
    'Connector': 'Set s = ws.Shapes.AddConnector(1, 100, 50, 120, 60)',
    'Straight': 'Set s = ws.Shapes.AddLine(100, 50, 220, 110)',
    'Picture': 'Set s = ws.Shapes.AddPicture(folder & "\\p.png", False, True, 100, 50, 60, 60)',
    'CheckBox': 'Set s = ws.Shapes.AddFormControl(xlCheckBox, 100, 50, 120, 60)',
    'Button': 'Set s = ws.Shapes.AddFormControl(xlButtonControl, 100, 50, 120, 60)',
    'DropDown': 'Set s = ws.Shapes.AddFormControl(xlDropDown, 100, 50, 120, 20)',
    'Spinner': 'Set s = ws.Shapes.AddFormControl(xlSpinner, 100, 50, 20, 60)',
}
LOADED = ('Plain', 'Box', 'Joined', 'Elbow', 'Line', 'Pic', 'Tuned', 'Clear')
LOCKS = ('s.LockAspectRatio = msoTrue', 's.LockAspectRatio = msoFalse')
OLD = ('s.DrawingObject.Interior.Color = 255', 's.DrawingObject.Interior.ColorIndex = 5', 's.DrawingObject.Interior.Pattern = xlNone', 's.DrawingObject.Border.Color = 255', 's.DrawingObject.Border.ColorIndex = 5',
       's.DrawingObject.Border.LineStyle = xlNone', 's.DrawingObject.Border.LineStyle = xlDash', *(f's.DrawingObject.Border.Weight = {weight}' for weight in ('xlHairline', 'xlThin', 'xlMedium', 'xlThick')),
       *(f's.Line.Weight = {weight}' for weight in ('0.25', '0.5', '0.75', '1', '1.25', '1.5', '2', '2.25', '2.5', '3', '3.5', '4.5', '6', '10')))
CONTROL = ('s.AlternativeText = "Alt"', 's.Title = "Ttl"', 's.Rotation = 30', 's.Visible = msoFalse', 's.Placement = xlMove', 's.Placement = xlFreeFloating', 's.Placement = xlMoveAndSize', 's.Decorative = msoTrue', 's.Locked = msoFalse')
READS = ('s.LockAspectRatio', 's.AlternativeText', 's.Title', 's.Rotation', 's.Visible', 's.Placement', 's.Decorative', 's.Locked', 's.Line.Weight', 's.DrawingObject.Interior.Color', 's.DrawingObject.Interior.ColorIndex', 's.DrawingObject.Interior.Pattern',
         's.DrawingObject.Border.Color', 's.DrawingObject.Border.ColorIndex', 's.DrawingObject.Border.LineStyle', 's.DrawingObject.Border.Weight', 's.Fill.ForeColor.RGB', 's.Fill.Visible', 's.Line.ForeColor.RGB', 's.Line.Visible', 's.Line.DashStyle')
CONTROLS = ('CheckBox', 'Button', 'DropDown', 'Spinner')
CASES: list[tuple[str, str]] = []
for target in (*NEW, *LOADED):
    CASES += [(target, statement) for statement in LOCKS]
for target in ('Rect', 'TextBox', 'Connector', 'Picture', 'Plain', 'Box', 'Joined', 'Pic'):
    CASES += [(target, statement) for statement in OLD]
for target in CONTROLS:
    CASES += [(target, statement) for statement in CONTROL]
#: Nothing set: what each reads from the first.
CASES += [(target, '') for target in (*NEW, *LOADED)]

_READER = '\n'.join(['Private Function Shown(v As Variant) As String', '    If IsObject(v) Then', '        Shown = "object"', '    ElseIf IsArray(v) Or IsNull(v) Then', '        Shown = TypeName(v)', '    Else', '        Shown = CStr(v)', '    End If', 'End Function', '',
                     'Private Function Reads(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next',
                     *[line for read in READS for line in ('    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E" & Err.Number', '    Else', '        r = r & "|" & Shown(v)', '    End If')],
                     '    Reads = r', 'End Function', '',
                     'Private Function TypesOf(s As Object) As String', '    Dim v As Variant, r As String', '    On Error Resume Next',
                     *[line for read in READS for line in ('    Err.Clear', f'    v = {read}', '    If Err.Number Then', '        r = r & "|E"', '    Else', '        r = r & "|" & TypeName(v)', '    End If')],
                     '    TypesOf = r', 'End Function', ''])


def source(target: str, statement: str) -> str:
    """One case, on Excel's Worksheets(1)."""
    made = target in NEW
    lines = [_READER, 'Public Function Go(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Object, s As Object, r As String, name As String', '    Application.DisplayAlerts = False']
    lines += ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    {NEW[target]}', '    s.Name = "Target"'] if made else [
        '    Set wb = Workbooks.Open(folder & "\\book.xlsx")', '    Set ws = wb.Worksheets(1)', f'    Set s = ws.Shapes("{target}")']
    lines += ['    name = s.Name', '    On Error Resume Next', f'    {statement}' if statement else "    ' Nothing is set.", '    If Err.Number Then r = "E" & Err.Number & "!"', '    Err.Clear', '    On Error GoTo 0', '    r = r & Reads(s) & "##" & TypesOf(s) & "@@"',
              '    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Set wb = Workbooks.Open(folder & "\\saved.xlsx")', '    r = r & Reads(wb.Worksheets(1).Shapes(name)) & "##" & TypesOf(wb.Worksheets(1).Shapes(name))', '    wb.Close False', '    Application.DisplayAlerts = True', '    Go = r', 'End Function']
    return '\n'.join(lines) + '\n'


def parts_of(path: Path, name: str) -> dict[str, str]:
    """The shape's anchor, and any VML part, with the text that names the file it was saved from gone."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        found: dict[str, str] = {}
        drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
        for one in re.finditer(r'<xdr:(?:twoCell|oneCell|absolute)Anchor\b.*?</xdr:(?:twoCell|oneCell|absolute)Anchor>', drawing, re.S):
            if f'name="{name}"' in one.group():
                found['anchor'] = one.group()
        for part in names:
            if part.endswith('.vml'):
                found['vml'] = package.read(part).decode('utf-8')
            if part.startswith('xl/worksheets/sheet') and part.endswith('.xml') and 'mc:AlternateContent' in (text := package.read(part).decode('utf-8')):
                control = re.search(r'<controls>.*?</controls>', text, re.S)
                if control:
                    found['controls'] = control.group()
    return found


def main() -> None:
    records = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        base = Path(folder)
        (base / 'p.png').write_bytes(png())
        shutil.copyfile(BOOK, base / 'book.xlsx')
        excel.new_document()
        for target, statement in CASES:
            result = excel.run_vba(source(target, statement), 'Go', (folder,), timeout=60.0, line_numbers=False)
            assert result.ok and not result.dialogs, (target, statement, result)
            name = 'Target' if target in NEW else target
            records.append({'target': target, 'statement': statement, 'answer': str(result.value), **parts_of(base / 'saved.xlsx', name)})
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'new': NEW, 'cases': records}, indent=0) + '\n', encoding='utf-8')
    print(len(records), 'cases;', sum(one['answer'].startswith('E') for one in records), 'raised')


if __name__ == '__main__':
    main()
