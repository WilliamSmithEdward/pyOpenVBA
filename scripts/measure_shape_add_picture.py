"""Shapes.AddPicture: what a picture reads, the errors, and what a saved workbook holds for it, from small PNG, JPEG-free
test images this script writes.

    python scripts/measure_shape_add_picture.py

writes tests/fixtures/shape_add_picture.json, which tests/test_excel_shape_add_picture.py replays.
"""
from __future__ import annotations

import base64
import json
import re
import struct
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_add_picture.json'


def png(width: int, height: int, dpi: int | None = None) -> bytes:
    """A small opaque red PNG, with a pHYs density when given."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * width for _ in range(height))
    parts = [b'\x89PNG\r\n\x1a\n', chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))]
    if dpi:
        metres = round(dpi / 0.0254)
        parts.append(chunk(b'pHYs', struct.pack('>IIB', metres, metres, 1)))
    parts += [chunk(b'IDAT', zlib.compress(rows, 9)), chunk(b'IEND', b'')]
    return b''.join(parts)


IMAGES = {'a.png': png(8, 4), 'b.png': png(8, 4, 96), 'c.png': png(8, 4, 192), 'd.png': png(20, 10, 72)}
ADD = 'ws.Shapes.AddPicture(folder & "\\{name}", {link}, {save}, {box})'
BOX = '10, 10, 100, 40'
CASES = [
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box="10, 10, -1, -1")}',
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box="10, 10, -1, -1")}',
    f'Set sh = {ADD.format(name="c.png", link="msoFalse", save="msoTrue", box="10, 10, -1, -1")}',
    f'Set sh = {ADD.format(name="d.png", link="msoFalse", save="msoTrue", box="10, 10, -1, -1")}',
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box="10, 10, 50, -1")}',
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box="10, 10, -1, 50")}',
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box="10, 10, 0, 0")}',
    f'Set sh = {ADD.format(name="missing.png", link="msoFalse", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoFalse", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoTrue", save="msoFalse", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoTrue", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: '
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: '
    f'Set sh = {ADD.format(name="b.png", link="msoFalse", save="msoTrue", box=BOX)}',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.Delete: Set sh = ws.Shapes.AddShape(1, 1, 1, 9, 9)',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.Name = "Mine": sh.Left = 30: sh.Width = 60',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.Line.Visible = msoTrue',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.Line.Weight = 3',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.Fill.ForeColor.RGB = 255',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: sh.LockAspectRatio = msoFalse',
    f'Set sh = {ADD.format(name="a.png", link="msoFalse", save="msoTrue", box=BOX)}: Set sh = sh.Duplicate',
    f'Set sh = ws.Shapes.AddPicture2(folder & "\\a.png", msoFalse, msoTrue, {BOX}, 1)',
    f'Set sh = ws.Shapes.AddPicture2(folder & "\\a.png", msoFalse, msoTrue, {BOX}, 3)',
]
READS = ['TypeName(sh)', 'sh.Name', 'sh.Type', 'sh.AutoShapeType', 'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height', 'sh.LockAspectRatio',
         'sh.ID', 'sh.Fill.Visible', 'sh.Line.Visible', 'sh.Line.Weight', 'TypeName(sh.PictureFormat)', 'sh.PictureFormat.Brightness',
         'sh.PictureFormat.Contrast', 'sh.PictureFormat.Crop.ShapeWidth', 'sh.PictureFormat.CropLeft', 'sh.Placement', 'ws.Shapes.Count']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String, ByVal out As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select', '    s = "ok"', 'Measured:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "," & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs out & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function',
              '', 'Public Function Probe(ByVal folder As String, ByVal out As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder, out)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def saved(path: Path, images: Path) -> dict[str, object]:
    """The parts a saved workbook holds for its pictures, with the images' folder taken out."""
    found: dict[str, object] = {}
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        found['media'] = {name: base64.b64encode(package.read(name)).decode('ascii') for name in names if '/media/' in name}
        for name in names:
            if name == 'xl/drawings/drawing1.xml' or name == 'xl/drawings/_rels/drawing1.xml.rels' or name == '[Content_Types].xml':
                found[name] = package.read(name).decode('utf-8').replace(str(images), '<folder>').replace(
                    str(images).replace('\\', '/'), '<folder>')
    return found


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as images, tempfile.TemporaryDirectory() as out, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        for name, data in IMAGES.items():
            (Path(images) / name).write_bytes(data)
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (images, out), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [saved(Path(out) / f'c{index}.xlsx', Path(images)) for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'images': {name: base64.b64encode(data).decode('ascii') for name, data in IMAGES.items()},
              'cases': [[statements, answer, part] for statements, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    images_bytes = {base64.b64encode(data).decode('ascii'): name for name, data in IMAGES.items()}
    for statements, answer, part in record['cases']:
        print(statements.replace('ws.Shapes.AddPicture(folder & "\\', 'AP("')[:90], '->', answer)
        media = part['media']
        print('    media', {k: images_bytes.get(v, f'{len(v)}b64') for k, v in media.items()})
        print('    rels', re.sub(r'xmlns="[^"]*"', '', str(part.get('xl/drawings/_rels/drawing1.xml.rels', '')))[50:330])
        drawing = re.sub(r'<xdr:wsDr[^>]*>', '', str(part.get('xl/drawings/drawing1.xml', '')))
        print('    ', re.search(r'<xdr:pic.*?</xdr:pic>|<xdr:sp.*?</xdr:sp>|$', drawing, re.S).group(0)[:760])


if __name__ == '__main__':
    main()
