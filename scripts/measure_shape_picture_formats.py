"""Shapes.AddPicture of JPEG, GIF and BMP files, and PictureFormat's setters: what a picture reads, the errors, and what
a saved workbook holds for it. The images are small ones this script makes, kept in the fixture.

    python scripts/measure_shape_picture_formats.py

writes tests/fixtures/shape_picture_formats.json, which tests/test_excel_shape_picture_formats.py replays.
"""
from __future__ import annotations

import base64
import io
import json
import re
import struct
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_picture_formats.json'


def saved_image(kind: str, size: tuple[int, int] = (8, 4), **options: object) -> bytes:
    out = io.BytesIO()
    Image.new('RGB', size, (200, 30, 30)).save(out, kind, **options)  # pyright: ignore[reportUnknownMemberType]
    return out.getvalue()


def bmp(width: int, height: int, ppm: int) -> bytes:
    """An uncompressed 24-bit BMP with this density."""
    row = (b'\x1e\x1e\xc8' * width).ljust((width * 3 + 3) // 4 * 4, b'\0')
    pixels = row * height
    header = struct.pack('<IiiHHIIiiII', 40, width, height, 1, 24, 0, len(pixels), ppm, ppm, 0, 0)
    return b'BM' + struct.pack('<IHHI', 54 + len(pixels), 0, 0, 54) + header + pixels


IMAGES = {
    'none.jpg': saved_image('JPEG'),
    'dpi96.jpg': saved_image('JPEG', dpi=(96, 96)),
    'dpi72.jpeg': saved_image('JPEG', size=(20, 10), dpi=(72, 72)),
    'dpi192.jpg': saved_image('JPEG', dpi=(192, 192)),
    'skew.jpg': saved_image('JPEG', dpi=(96, 192)),
    'plain.gif': saved_image('GIF'),
    'big.gif': saved_image('GIF', size=(20, 10)),
    'zero.bmp': bmp(8, 4, 0),
    'dpi96.bmp': bmp(8, 4, 3780),
    'shown.png': saved_image('PNG'),
    'misnamed.png': saved_image('JPEG', dpi=(96, 96)),
    'misnamed.jpg': saved_image('PNG'),
    'noext': saved_image('GIF'),
    'text.png': b'not an image at all',
}
ADD = 'ws.Shapes.AddPicture(folder & "\\{name}", msoFalse, msoTrue, {box})'
NATURAL = '10, 10, -1, -1'
BOX = '10, 10, 100, 40'
PNG = ADD.format(name='shown.png', box=BOX)
FORMAT = 'sh.PictureFormat.'
CASES = [
    *[f'Set sh = {ADD.format(name=name, box=NATURAL)}' for name in IMAGES],
    *[f'Set sh = {ADD.format(name=name, box=BOX)}' for name in ('none.jpg', 'plain.gif', 'zero.bmp')],
    f'Set sh = {ADD.format(name="none.jpg", box=BOX)}: Set sh = {ADD.format(name="none.jpg", box=BOX)}',
    f'Set sh = {ADD.format(name="none.jpg", box=BOX)}: Set sh = {ADD.format(name="plain.gif", box=BOX)}: Set sh = {PNG}',
    *[f'Set sh = {PNG}: {FORMAT}{one}' for one in (
        'Brightness = 0.7', 'Brightness = 0', 'Brightness = 1', 'Brightness = 0.5', 'Brightness = 2', 'Brightness = -1',
        'Contrast = 0.7', 'Contrast = 0', 'Contrast = 1', 'Contrast = 0.5', 'Contrast = 2',
        'IncrementBrightness 0.1', 'IncrementBrightness -0.1', 'IncrementBrightness 0.9', 'IncrementContrast 0.1', 'IncrementContrast -0.1',
        'ColorType = 1', 'ColorType = 2', 'ColorType = 3', 'ColorType = 4', 'ColorType = 0', 'ColorType = 5', 'ColorType = -2',
        'CropLeft = 10', 'CropRight = 10', 'CropTop = 10', 'CropBottom = 10', 'CropLeft = -10', 'CropLeft = 200', 'CropLeft = 0',
        'TransparentBackground = msoTrue', 'TransparentBackground = msoFalse', 'TransparencyColor = 255',
        'TransparencyColor = 255: sh.PictureFormat.TransparentBackground = msoTrue',
        'TransparentBackground = msoTrue: sh.PictureFormat.TransparencyColor = 255',
        'Brightness = 0.7: sh.PictureFormat.Contrast = 0.7', 'Contrast = 0.7: sh.PictureFormat.Brightness = 0.7',
        'ColorType = 2: sh.PictureFormat.Brightness = 0.7', 'Brightness = 0.7: sh.PictureFormat.ColorType = 2',
        'CropLeft = 10: sh.PictureFormat.CropRight = 5: sh.PictureFormat.CropTop = 3: sh.PictureFormat.CropBottom = 2',
        'CropLeft = 10: sh.Width = 200', 'CropLeft = 10: sh.PictureFormat.CropLeft = 0',
        'Crop.ShapeWidth = 50', 'Crop.ShapeHeight = 20', 'Crop.ShapeLeft = 20', 'Crop.ShapeTop = 20', 'Crop.PictureWidth = 50',
        'Crop.PictureHeight = 20', 'Crop.PictureOffsetX = 5', 'Crop.PictureOffsetY = 5')],
    f'Set sh = {ADD.format(name="shown.png", box=NATURAL)}: {FORMAT}CropLeft = 1',
    f'Set sh = {ADD.format(name="none.jpg", box=BOX)}: {FORMAT}Brightness = 0.7',
    'Set sh = ws.Shapes.AddShape(1, 10, 10, 100, 40): sh.PictureFormat.Brightness = 0.7',
    'Set sh = ws.Shapes.AddShape(1, 10, 10, 100, 40): sh.PictureFormat.CropLeft = 5',
    'Set sh = ws.Shapes.AddLine(10, 10, 110, 50): sh.PictureFormat.Brightness = 0.7',
]
READS = ['TypeName(sh)', 'sh.Name', 'sh.Type', 'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height', 'sh.LockAspectRatio', 'sh.ID', 'ws.Shapes.Count',
         FORMAT + 'Brightness', FORMAT + 'Contrast', FORMAT + 'ColorType', FORMAT + 'CropLeft', FORMAT + 'CropRight', FORMAT + 'CropTop',
         FORMAT + 'CropBottom', FORMAT + 'TransparentBackground', FORMAT + 'TransparencyColor', 'TypeName(sh.PictureFormat.Crop)',
         FORMAT + 'Crop.ShapeWidth', FORMAT + 'Crop.ShapeHeight', FORMAT + 'Crop.ShapeLeft', FORMAT + 'Crop.ShapeTop', FORMAT + 'Crop.PictureWidth',
         FORMAT + 'Crop.PictureHeight', FORMAT + 'Crop.PictureOffsetX', FORMAT + 'Crop.PictureOffsetY']


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
                  '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s',
              '    wb.SaveAs out & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    s = "Err " & Err.Number', '    Resume Measured', 'End Function',
              '', 'Public Function Probe(ByVal folder As String, ByVal out As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder, out)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def saved(path: Path, images: Path) -> dict[str, object]:
    """The parts a saved workbook holds for its pictures. No path of this machine stays in them."""
    found: dict[str, object] = {}
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        found['media'] = {name: base64.b64encode(package.read(name)).decode('ascii') for name in names if '/media/' in name}
        for name in names:
            if name in ('xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels', '[Content_Types].xml'):
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
    lookup = {value: key for key, value in record['images'].items()}
    for statements, answer, part in record['cases']:
        short = re.sub(r'ws\.Shapes\.AddPicture\(folder & "\\([^"]*)", msoFalse, msoTrue, ([^)]*)\)', r'AP(\1 \2)', statements).replace('sh.PictureFormat.', '')
        print(short[:70], '->', answer.replace('E-2147024809', 'EI').replace('E-2147467259', 'EU'))
        media = {key: lookup.get(value, f'reencoded {len(value)}') for key, value in part['media'].items()}
        drawing = str(part.get('xl/drawings/drawing1.xml', ''))
        blip = re.findall(r'<xdr:blipFill>.*?</xdr:blipFill>', drawing, re.S)
        cleaned = [re.sub(r'<a:extLst>.*?</a:extLst>|xmlns:\w+="[^"]*"', '', one) for one in blip]
        types = re.findall(r'<Default Extension="(\w+)"', str(part.get('[Content_Types].xml', '')))
        print('    ', media, types, cleaned[-1:] if cleaned else '')


if __name__ == '__main__':
    main()
