"""What a shape, a ShapeRange or a part of a shape answers when held after the shape has gone, for each kind of shape and each way
of deleting one: its TypeName, Is Nothing, and the error from reading, setting and calling a member.

    python scripts/measure_shape_held_deleted.py

writes tests/fixtures/shape_held_deleted.json, which tests/test_excel_shape_held_deleted.py replays. Each case is
[made, held, gone, answers], one answer for each member tried on what is held, as the number of the error or ok.
"""
from __future__ import annotations

import json
import struct
import tempfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_held_deleted.json'
RECT = 'Set sh = ws.Shapes.AddShape(1, 10, 10, 100, 40)'
KINDS = {
    'rect': RECT, 'oval': 'Set sh = ws.Shapes.AddShape(9, 10, 10, 100, 40)', 'textbox': 'Set sh = ws.Shapes.AddTextbox(1, 10, 10, 100, 40)',
    'line': 'Set sh = ws.Shapes.AddLine(10, 10, 50, 30)', 'connector': 'Set sh = ws.Shapes.AddConnector(1, 10, 10, 50, 30)',
    'picture': 'Set sh = ws.Shapes.AddPicture(folder & "\\p.png", msoFalse, msoTrue, 10, 10, 20, 20)',
    'group': f'{RECT}: Set o = ws.Shapes.AddShape(1, 200, 10, 50, 50): Set sh = ws.Shapes.Range(Array(sh.Name, o.Name)).Group',
    'chart': 'ws.ChartObjects.Add 10, 10, 100, 100: Set sh = ws.Shapes(1)', 'button': 'ws.Buttons.Add 10, 10, 100, 40: Set sh = ws.Shapes(1)',
    'checkbox': 'ws.CheckBoxes.Add 10, 10, 100, 40: Set sh = ws.Shapes(1)',
    'freeform': 'Set sh = ws.Shapes.BuildFreeform(0, 10, 10).ConvertToShape',
}
FREEFORM_REAL = 'With ws.Shapes.BuildFreeform(0, 10, 10): .AddNodes 0, 0, 50, 10: .AddNodes 0, 0, 50, 50: Set sh = .ConvertToShape: End With'
KINDS['freeform'] = FREEFORM_REAL
#: What can be held, and what is tried on it: a read, a set and a call that work while the shape is there.
SHAPE = ('sh', 'h.Name', 'h.Name = "Z"', 'h.ZOrder 0')
PARTS = [
    SHAPE, ('sh.Fill', 'h.Visible', 'h.Visible = msoTrue', 'h.Solid'), ('sh.Fill.ForeColor', 'h.RGB', 'h.RGB = 255', None), ('sh.Fill.BackColor', 'h.RGB', 'h.RGB = 255', None),
    ('sh.Line', 'h.Weight', 'h.Weight = 2', None), ('sh.Line.ForeColor', 'h.RGB', 'h.RGB = 255', None),
    ('sh.TextFrame', 'h.MarginLeft', 'h.MarginLeft = 1', None), ('sh.TextFrame.Characters', 'h.Text', 'h.Text = "x"', None), ('sh.TextFrame.Characters(1, 1)', 'h.Count', 'h.Text = "x"', 'h.Delete'),
    ('sh.TextFrame.Characters.Font', 'h.Bold', 'h.Bold = True', None),
    ('sh.TextFrame2', 'h.MarginLeft', 'h.MarginLeft = 1', 'h.DeleteText'), ('sh.TextFrame2.TextRange', 'h.Text', 'h.Text = "x"', 'h.Delete'), ('sh.TextFrame2.TextRange.Font', 'h.Bold', 'h.Bold = msoTrue', None),
    ('sh.TextFrame2.TextRange.Characters(1, 1)', 'h.Length', 'h.Text = "x"', 'h.InsertAfter "y"'), ('sh.TextFrame2.TextRange.Font.Fill', 'h.Visible', 'h.Visible = msoTrue', 'h.Solid'),
    ('sh.TextFrame2.TextRange.Font.Line', 'h.Visible', 'h.Visible = msoTrue', None),
    ('sh.Shadow', 'h.Visible', 'h.Visible = msoTrue', None), ('sh.Glow', 'h.Radius', 'h.Radius = 1', None), ('sh.SoftEdge', 'h.Radius', 'h.Radius = 1', None), ('sh.Reflection', 'h.Type', 'h.Type = 1', None),
    ('sh.ThreeD', 'h.Visible', 'h.Visible = msoFalse', None), ('sh.Adjustments', 'h.Count', None, None), ('sh.ConnectorFormat', 'h.Type', None, None), ('sh.Fill.PictureEffects', 'h.Count', None, None),
    ('ws.Shapes.Range(Array(sh.Name))', 'h.Name', 'h.Left = 5', 'h.ZOrder 0'), ('ws.Shapes.Range(Array(sh.Name)).Fill', 'h.Visible', 'h.Visible = msoTrue', None), ('ws.Shapes.Range(Array(sh.Name)).Line', 'h.Weight', 'h.Weight = 2', None),
]
RANGE = 'ws.Shapes.Range(Array(sh.Name))'
#: Further members of a shape, to show no member stands apart.
MEMBERS = [('sh', one, None, None) for one in ('h.Left', 'h.Type', 'h.Parent.Name', 'h.ID', 'h.Fill.Visible', 'h.TextFrame2.HasText', 'h.Application.Name', 'h.Creator', 'h.NoSuchMember', 'h.Connector')]
MEMBERS += [('sh', 'h.Name', one, None) for one in ('h.Left = 5', 'h.NoSuchMember = 5', 'h.OnAction = "x"')] + [('sh', 'h.Name', None, one) for one in ('h.Delete', 'h.Select', 'h.Duplicate', 'h.NoSuchMember')]
MEMBERS += [('sh.Fill', one, None, None) for one in ('h.Parent.Name', 'h.Application.Name', 'h.Creator', 'h.NoSuchMember')] + [('sh.TextFrame2', one, None, None) for one in ('h.Parent.Name', 'h.Application.Name', 'h.Creator')]
MEMBERS += [('sh.TextFrame2.TextRange', one, None, None) for one in ('h.Parent.Name', 'h.Application.Name', 'h.Count')]
MEMBERS += [(RANGE, one, None, None) for one in ('h.Count', 'h.Item(1).Name', 'h(1).Name', 'h.Application.Name', 'h.Creator', 'h.Parent.Name', 'h.Left')]
CASES: list[tuple[str, tuple[str, str | None, str | None, str | None], str]] = [
    *[(RECT, part, 'sh.Delete') for part in PARTS],
    *[(make, SHAPE, 'sh.Delete') for make in KINDS.values() if make != RECT],
    *[(make, part, 'sh.Delete') for make in KINDS.values() if make != RECT for part in PARTS[1:2] + PARTS[4:5] + PARTS[10:11] + PARTS[24:25]],
    *[(RECT, part, 'sh.Delete') for part in MEMBERS],
    *[(RECT, part, gone) for part in (SHAPE, PARTS[1], PARTS[10], PARTS[24]) for gone in (
        'ws.Shapes(sh.Name).Delete', 'ws.Shapes.Range(Array(sh.Name)).Delete', 'ws.Shapes(1).Delete', 'h.Parent.Delete: Err.Clear', 'ws.Delete', 'sh.Cut', 'ws.Shapes.SelectAll: Selection.Delete', 'ws.DrawingObjects.Delete')],
    # What is still there: a copy, a shape named as the one that went, and a range of which one member remains.
    (RECT, SHAPE, 'sh.Copy'), (RECT, SHAPE, 's2 = sh.Name: sh.Delete: Set o = ws.Shapes.AddShape(1, 1, 1, 5, 5): o.Name = s2'),
    *[(f'{RECT}: Set o = ws.Shapes.AddShape(1, 200, 10, 50, 50)', ('ws.Shapes.Range(Array(sh.Name, o.Name))', one, None, None), 'sh.Delete') for one in ('h.Count', 'h.Name', 'h.Left', 'h.Item(1).Name', 'h.Item(2).Name', 'h(1).Name')],
    (f'{RECT}: Set o = ws.Shapes.AddShape(1, 200, 10, 50, 50)', ('ws.Shapes.Range(Array(sh.Name, o.Name))', 'h.Name', 'h.Left = 5', 'h.ZOrder 0'), 'sh.Delete'),
    (f'{RECT}: Set o = ws.Shapes.AddShape(1, 200, 10, 50, 50)', ('ws.Shapes.Range(Array(sh.Name, o.Name))', 'h.Name', 'h.Left = 5', 'h.ZOrder 0'), 'ws.Shapes.Range(Array(sh.Name, o.Name)).Delete'),
    # A group, its members and Ungroup.
    (KINDS['group'], SHAPE, 'sh.Ungroup'), (KINDS['group'], ('sh.GroupItems(1)', 'h.Name', 'h.Left = 5', None), 'sh.Ungroup'), (KINDS['group'], ('sh.GroupItems(1)', 'h.Name', 'h.Left = 5', None), 'sh.Delete'),
    (KINDS['group'], ('sh.GroupItems', 'h.Count', None, None), 'sh.Delete'), (KINDS['group'], ('sh.GroupItems', 'h.Count', None, None), 'sh.Ungroup'), (KINDS['group'], ('sh.GroupItems(1)', 'h.Name', 'h.Left = 5', None), 'sh.GroupItems(1).Delete'),
    (KINDS['group'], SHAPE, 'sh.GroupItems(1).Delete'), (KINDS['group'], SHAPE, 'sh.GroupItems(1).Delete: sh.GroupItems(1).Delete'), (KINDS['group'], ('sh.GroupItems(2)', 'h.Name', 'h.Left = 5', None), 'sh.GroupItems(1).Delete'),
    # Charts and controls.
    (KINDS['chart'], ('ws.ChartObjects(1)', 'h.Name', 'h.Left = 5', 'h.Select'), 'sh.Delete'), (KINDS['chart'], ('ws.ChartObjects(1)', 'h.Name', 'h.Left = 5', 'h.Select'), 'ws.ChartObjects(1).Delete'), (KINDS['chart'], SHAPE, 'ws.ChartObjects(1).Delete'),
    (KINDS['chart'], ('ws.ChartObjects(1).Chart', 'h.Name', 'h.HasTitle = True', None), 'sh.Delete'), (KINDS['chart'], ('sh.Chart', 'h.Name', 'h.HasTitle = True', None), 'sh.Delete'),
    (KINDS['button'], ('sh.ControlFormat', 'h.Enabled', 'h.Enabled = True', None), 'sh.Delete'), (KINDS['button'], ('ws.Buttons(1)', 'h.Name', 'h.Left = 5', None), 'sh.Delete'), (KINDS['button'], SHAPE, 'ws.Buttons(1).Delete'),
    (KINDS['checkbox'], ('sh.ControlFormat', 'h.Value', 'h.Value = 1', None), 'sh.Delete'), (KINDS['picture'], ('sh.PictureFormat', 'h.Brightness', 'h.Brightness = 0.5', 'h.IncrementBrightness 0.1'), 'sh.Delete'),
    (KINDS['freeform'], ('sh.Nodes', 'h.Count', None, None), 'sh.Delete'), (KINDS['group'], ('sh.GroupItems(1).Fill', 'h.Visible', 'h.Visible = msoTrue', None), 'sh.Delete'), (KINDS['group'], ('sh.GroupItems(1).Fill', 'h.Visible', 'h.Visible = msoTrue', None), 'sh.GroupItems(1).Delete'),
    # And in plain code: a With block, a For Each variable, and a function's argument.
    (RECT, ('sh', 'h.Name', None, None), 'For Each o In ws.Shapes: o.Delete: Next'), (RECT, ('sh', 'h.Name', None, None), 'ws.Shapes.Range(Array(1)).Delete'),
]


def _probe() -> str:
    lines: list[str] = []
    chunks = [range(at, min(at + 30, len(CASES))) for at in range(0, len(CASES), 30)]
    for number, chunk in enumerate(chunks):
        lines += [f'Private Function Act{number}(ByVal which As Long, ByVal ws As Worksheet, ByVal folder As String) As String', '    Dim sh As Object, h As Object, o As Object, s2 As Variant, v As Variant, s As String',
                  '    On Error Resume Next', '    Select Case which']
        for index in chunk:
            made, (held, read, write, call), gone = CASES[index]
            lines += [f'    Case {index}', f'        {made}' if not made.startswith('With') else '        ' + made.replace(': ', '\n        '), f'        Set h = {held}',
                      '        s = "made " & Err.Number: Err.Clear', f'        {gone}', '        s = s & "~gone " & Err.Number: Err.Clear',
                      '        v = "": v = TypeName(h): s = s & "~" & v & "/" & Err.Number: Err.Clear', '        v = "": v = (h Is Nothing): s = s & "~" & v & "/" & Err.Number: Err.Clear']
            for try_ in (read,):
                lines += [f'        v = "": v = {try_}: s = s & "~" & Err.Number: Err.Clear' if try_ else '        s = s & "~"']
            for try_ in (write, call):
                lines += [f'        {try_}: s = s & "~" & Err.Number: Err.Clear' if try_ else '        s = s & "~"']
        lines += ['    End Select', f'    Act{number} = s', 'End Function', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
              '    wb.Worksheets.Add After:=ws', '    Select Case which \\ 30']
    for number in range(len(chunks)):
        lines += [f'    Case {number}: Tried = Act{number}(which, ws, folder)']
    lines += ['    End Select', '    wb.Close False', 'End Function', '', 'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.CutCopyMode = False', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\0\0\xff\0\0' * 2)) + chunk(b'IEND', b'')


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        (Path(folder) / 'p.png').write_bytes(png())
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=560.0, line_numbers=False)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source.replace('\\\\', '\\'),
              'cases': [[made, list(part), gone, answer] for (made, part, gone), answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
