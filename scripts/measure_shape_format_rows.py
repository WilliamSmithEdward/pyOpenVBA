"""The shape formatting rows left MISSING: ShapeRange's PictureFormat, Nodes, Vertices, Callout, TextEffect, ThreeD and
_Default over one shape and over two; Application, Creator and Parent on ShadowFormat, GlowFormat, SoftEdgeFormat, ReflectionFormat and a
header picture's Graphic; and a Graphic's Brightness, ColorType, Contrast and crops on a sheet with no header picture.
Pictures come from a PNG under a TEMP folder. Needs Excel:

    python scripts/measure_shape_format_rows.py

writes tests/fixtures/shape_format_rows.json, which tests/test_excel_shape_format_rows.py replays.

Predictions, written before Excel ran:
  ShapeRange of one shape. P: each member answers what the shape's own member answers, the same object type and values.
    Rival Q: the range answers an object of its own kind, or raises, where the shape answers.
  ShapeRange of two shapes. P: PictureFormat, Callout and ThreeD give the object, and a value that differs reads -2
    (mixed), as Fill does. Nodes and Vertices are out of range (-2147024809), as PickUp is.
    Rival Q: every member over two shapes is out of range.
  _Default(1) bound early answers the first shape, as r(1) does. Bound late it raises 438, as Shapes' and ChartObjects' do.
  Trio: Application answers Application on all five. Creator is 1480803660. Parent of the shadow is the shape, as
    tests/fixtures/shape_effects.json reads. Assigning Application or Creator raises 450.
    Rival: the effect formats raise 438 for Application, as GlowFormat.Parent does.
  Graphic with no header picture. P: every read and every set raises 1004, as Filename, Height and Width do.
    Rival Q: the reads answer defaults (Brightness 0.5, Contrast 0.5, ColorType 1, crops 0) and the sets do nothing.
"""

from __future__ import annotations

import json
import struct
import tempfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_object_trio import _TRIO

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_format_rows.json'

#: What each case's range holds: statements making it from ws and folder, giving r, a ShapeRange.
RANGES = {
    'rect': 'ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 40, 30).Name = "A": Set r = ws.Shapes.Range(Array("A"))',
    'free': ('Set b = ws.Shapes.BuildFreeform(1, 10, 10): b.AddNodes 0, 0, 110, 10: b.AddNodes 0, 0, 60, 80: '
             'b.ConvertToShape.Name = "A": Set r = ws.Shapes.Range(Array("A"))'),
    'picture': ('ws.Shapes.AddPicture(folder & "\\a.png", msoFalse, msoTrue, 10, 10, 20, 10).Name = "A": '
                'Set r = ws.Shapes.Range(Array("A"))'),
    'callout': 'ws.Shapes.AddCallout(msoCalloutTwo, 10, 10, 40, 30).Name = "A": Set r = ws.Shapes.Range(Array("A"))',
    'two': ('ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 40, 30).Name = "A": '
            'ws.Shapes.AddShape(msoShapeOval, 60, 10, 40, 30).Name = "B": Set r = ws.Shapes.Range(Array("A", "B"))'),
    'two pictures': ('ws.Shapes.AddPicture(folder & "\\a.png", msoFalse, msoTrue, 10, 10, 20, 10).Name = "A": '
                     'ws.Shapes.AddPicture(folder & "\\a.png", msoFalse, msoTrue, 40, 10, 20, 10).Name = "B": '
                     'ws.Shapes("B").PictureFormat.Brightness = 0.7: Set r = ws.Shapes.Range(Array("A", "B"))'),
}
#: Each read, over r (the range) and sh (its first shape, for comparison).
READS = [
    'TypeName(r.PictureFormat)', 'r.PictureFormat.Brightness', 'r.PictureFormat.CropLeft', 'r.PictureFormat.ColorType',
    'sh.PictureFormat.Brightness',
    'TypeName(r.Nodes)', 'r.Nodes.Count', 'sh.Nodes.Count', 'TypeName(r.Vertices)', 'UBound(r.Vertices, 1)',
    'r.Vertices(1, 1)', 'TypeName(r.Callout)', 'r.Callout.Type', 'sh.Callout.Type', 'r.Callout.AutoAttach',
    'TypeName(r.TextEffect)', 'r.TextEffect.Text', 'TypeName(sh.TextEffect)', 'TypeName(r.ThreeD)', 'r.ThreeD.Visible',
    'TypeName(sh.ThreeD)', 'sh.ThreeD.Visible', 'r.[_Default](1).Name', 'r(1).Name', 'o.[_Default](1).Name',
]
#: A Graphic's members on a sheet with no header picture: reads, then sets.
GRAPHIC_READS = ['g.Brightness', 'g.ColorType', 'g.Contrast', 'g.CropBottom', 'g.CropLeft', 'g.CropRight', 'g.CropTop',
                 'TypeName(g.Application)']
GRAPHIC_SETS = ['g.Brightness = 0.6', 'g.ColorType = msoPictureGrayscale', 'g.Contrast = 0.6', 'g.CropBottom = 1',
                'g.CropLeft = 1', 'g.CropRight = 1', 'g.CropTop = 1']
#: Expressions for the trio, from a sheet ws with a rectangle sh.
TRIO_REACHES = [('ShadowFormat', 'sh.Shadow'), ('GlowFormat', 'sh.Glow'), ('SoftEdgeFormat', 'sh.SoftEdge'),
                ('ReflectionFormat', 'sh.Reflection'),
                ('Graphic', 'ws.PageSetup.LeftHeaderPicture'), ('Graphic', 'ws.PageSetup.CenterFooterPicture')]


def png(width: int, height: int) -> bytes:
    """A small opaque red PNG."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * width for _ in range(height))
    return b''.join([b'\x89PNG\r\n\x1a\n', chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0)),
                     chunk(b'IDAT', zlib.compress(rows, 9)), chunk(b'IEND', b'')])


def _read_lines(read: str) -> list[str]:
    return ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']


def source() -> str:
    lines = ['Private Function Ranged(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, r As ShapeRange, sh As Shape, o As Object, b As Object',
             '    Dim s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', '    Select Case which']
    for index, statements in enumerate(RANGES.values()):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    Set sh = r(1): Set o = r', '    On Error Resume Next']
    for read in READS:
        lines += _read_lines(read)
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Ranged = Mid(s, 2)', 'End Function', '',
              'Private Function Graphics() As String',
              '    Dim wb As Workbook, ws As Worksheet, g As Object, s As String, v As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
              '    Set g = ws.PageSetup.LeftHeaderPicture', '    On Error Resume Next']
    for read in GRAPHIC_READS:
        lines += _read_lines(read)
    for statement in GRAPHIC_SETS:
        lines += ['    Err.Clear', f'    {statement}', '    s = s & "|" & Err.Number']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Graphics = Mid(s, 2)', 'End Function', '',
              'Private Function Trios() As String',
              '    Dim wb As Workbook, ws As Worksheet, sh As Shape, s As String',
              '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
              '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 40, 30)']
    for _, reach in TRIO_REACHES:
        lines += [f'    s = s & "@" & Replace(Trio({reach}), wb.Name, "<book>")']
    lines += ['    wb.Close False', '    Trios = Mid(s, 2)', 'End Function', _TRIO,
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(RANGES) - 1}',
              '        s = s & "@@@" & Ranged(i, folder)', '    Next',
              '    s = s & "@@@" & Graphics() & "@@@" & Trios()', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    code = source()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        (Path(folder) / 'a.png').write_bytes(png(8, 4))
        excel.new_document()
        done = excel.run_vba(code, 'Probe', (folder,), timeout=300.0, line_numbers=False)
        assert done.outcome == 'passed' and not done.dialogs, done
    parts = str(done.value).split('@@@')
    ranged, graphic, trio = parts[:len(RANGES)], parts[len(RANGES)], parts[len(RANGES) + 1]
    record = {
        'measured_at': datetime.now(timezone.utc).isoformat(), 'source': code, 'reads': READS,
        'ranges': [[name, answer.split('|')] for name, answer in zip(RANGES, ranged, strict=True)],
        'graphic': {'reads': GRAPHIC_READS, 'sets': GRAPHIC_SETS, 'answers': graphic.split('|')},
        'trio': [[name, reach, answer] for (name, reach), answer in zip(TRIO_REACHES, trio.split('@'), strict=True)],
    }
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answers in record['ranges']:
        print(name, dict(zip(READS, answers)))
    print('graphic', record['graphic']['answers'])
    for row in record['trio']:
        print(row)


if __name__ == '__main__':
    main()
