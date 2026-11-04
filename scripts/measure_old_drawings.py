"""Worksheet.TextBoxes, Pictures, Rectangles, Ovals, Lines and Arcs, and DrawingObjects: which shapes each collection holds, what its
Add makes, and what an item and the collection read, set and call, with the drawing each saved workbook holds.

    python scripts/measure_old_drawings.py

writes tests/fixtures/old_drawings.json, which tests/test_excel_old_drawings.py replays. Each case is [statements, answers,
saved shapes]; the answers are the reads in the record's reads, joined by ~, with E and a number for an error.
"""
from __future__ import annotations

import json
import re
import struct
import tempfile
import time
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/old_drawings.json'
ADDS = {
    'TextBoxes': 'ws.TextBoxes.Add(10, 10, 100, 40)', 'Rectangles': 'ws.Rectangles.Add(10, 10, 100, 40)', 'Ovals': 'ws.Ovals.Add(10, 10, 100, 40)',
    'Lines': 'ws.Lines.Add(10, 10, 110, 50)', 'Arcs': 'ws.Arcs.Add(10, 10, 110, 50)', 'Pictures': 'ws.Pictures.Insert(folder & "\\p.png")',
}
#: Which collection holds which shape.
OTHERS = ['ws.Shapes.AddShape(1, 1, 1, 9, 9)', 'ws.Shapes.AddShape(5, 1, 1, 9, 9)', 'ws.Shapes.AddShape(9, 1, 1, 9, 9)', 'ws.Shapes.AddShape(25, 1, 1, 9, 9)', 'ws.Shapes.AddShape(7, 1, 1, 9, 9)', 'ws.Shapes.AddTextbox(1, 1, 1, 9, 9)',
          'ws.Shapes.AddLine(1, 1, 9, 9)', 'ws.Shapes.AddConnector(1, 1, 1, 9, 9)', 'ws.Shapes.AddConnector(2, 1, 1, 9, 9)', 'ws.Shapes.AddPicture(folder & "\\p.png", msoFalse, msoTrue, 1, 1, 9, 9)', 'ws.Buttons.Add(1, 1, 9, 9)']
COUNTS = ['ws.TextBoxes.Count', 'ws.Rectangles.Count', 'ws.Ovals.Count', 'ws.Lines.Count', 'ws.Arcs.Count', 'ws.Pictures.Count', 'ws.DrawingObjects.Count', 'ws.Shapes.Count']
ITEM_READS = ['TypeName(h)', 'h.Name', 'h.Index', 'h.Left', 'h.Top', 'h.Width', 'h.Height', 'h.Visible', 'h.Locked', 'h.Placement', 'h.PrintObject', 'h.Enabled', 'h.OnAction', 'h.ZOrder', 'TypeName(h.Parent)', 'h.Application.Name', 'h.Creator',
              'h.TopLeftCell.Address', 'h.BottomRightCell.Address', 'TypeName(h.ShapeRange)', 'h.ShapeRange.Count', 'h.ShapeRange.Name']
TEXT_READS = ['h.Text', 'h.Caption', 'h.Formula', 'h.HorizontalAlignment', 'h.VerticalAlignment', 'h.Orientation', 'h.AutoSize', 'h.AddIndent', 'h.LockedText', 'h.RoundedCorners', 'h.ReadingOrder',
              'TypeName(h.Characters)', 'h.Characters.Text', 'h.Characters.Count', 'TypeName(h.Font)', 'h.Font.Bold', 'h.Font.Size', 'h.Font.Name', 'h.Font.Color']
PAINT_READS = ['TypeName(h.Interior)', 'h.Interior.ColorIndex', 'h.Interior.Color', 'h.Interior.Pattern', 'TypeName(h.Border)', 'h.Border.ColorIndex', 'h.Border.Color', 'h.Border.LineStyle', 'h.Border.Weight', 'h.Shadow']
LINE_READS = ['h.ArrowHeadStyle', 'h.ArrowHeadLength', 'h.ArrowHeadWidth', 'h.Border.LineStyle', 'h.Border.Weight', 'h.Border.Color']
PICTURE_READS = ['h.Formula', 'h.Interior.ColorIndex', 'h.Border.LineStyle', 'h.ShapeRange.Type']
READS = [*COUNTS, *ITEM_READS, *TEXT_READS, *PAINT_READS, *LINE_READS, *PICTURE_READS, 's2']
SETS = ['h.Name = "Mine"', 'h.Left = 20', 'h.Top = 20', 'h.Width = 50', 'h.Height = 25', 'h.Visible = False', 'h.Locked = False', 'h.Placement = xlMove', 'h.Placement = xlFreeFloating', 'h.PrintObject = False', 'h.Enabled = False', 'h.OnAction = "Macro1"',
        'h.Text = "Hello"', 'h.Caption = "Hello"', 'h.Formula = "=A1"', 'h.Characters.Text = "Hello"', 'h.Characters(1, 2).Font.Bold = True', 'h.Font.Bold = True', 'h.Font.Size = 14', 'h.Font.Name = "Arial"', 'h.Font.Color = 255', 'h.Font.ColorIndex = 3',
        'h.HorizontalAlignment = xlCenter', 'h.VerticalAlignment = xlCenter', 'h.Orientation = xlUpward', 'h.AutoSize = True', 'h.LockedText = False', 'h.RoundedCorners = True',
        'h.Interior.ColorIndex = 3', 'h.Interior.Color = 255', 'h.Interior.Pattern = xlNone', 'h.Interior.ColorIndex = xlNone', 'h.Border.ColorIndex = 3', 'h.Border.Color = 255', 'h.Border.LineStyle = xlNone', 'h.Border.LineStyle = xlDash', 'h.Border.Weight = xlThick', 'h.Shadow = True',
        'h.ArrowHeadStyle = xlArrowHeadStyleTriangle', 'h.ArrowHeadLength = xlArrowHeadLengthLong', 'h.ArrowHeadWidth = xlArrowHeadWidthWide', 'h.Delete', 'h.Select', 'h.BringToFront', 'h.SendToBack', 'Set o = h.Duplicate: s2 = TypeName(o) & "|" & o.Name', 'h.Copy', 'h.Cut']
TEXTUAL = ('TextBoxes', 'Rectangles', 'Ovals')
CASES: list[str] = [
    # What each Add makes, and its errors.
    's2 = 0', *[f'Set h = {add}' for add in ADDS.values()], *[f'Set h = ws.{name}.Add' for name in ('TextBoxes', 'Rectangles', 'Ovals', 'Lines', 'Arcs')], 'Set h = ws.TextBoxes.Add(10, 10)', 'Set h = ws.Lines.Add(110, 50, 10, 10)', 'Set h = ws.Lines.Add(10, 50, 110, 10)', 'Set h = ws.Arcs.Add(110, 50, 10, 10)',
    'Set h = ws.Rectangles.Add(10, 10, -5, 40)', 'Set h = ws.Rectangles.Add(10, 10, 0, 0)', 'Set h = ws.Pictures.Insert(folder & "\\none.png")', 'Set h = ws.Pictures.Insert', 'Set h = ws.Pictures.Insert(folder & "\\p.png", True)',
    # Which collection holds which shape.
    *[f'Set h = {other}: s2 = ws.Shapes(1).Name' for other in OTHERS],
    # Items.
    *[f'Set h = {ADDS[name]}: Set o = {ADDS[name]}: s2 = ws.{name}.Count & "|" & ws.{name}(2).Name & "|" & ws.{name}(ws.{name}(1).Name).Name & "|" & ws.{name}.Item(1).Name' for name in ADDS],
    *[f'Set h = {ADDS[name]}: s2 = TypeName(ws.{name}) & "|" & TypeName(ws.{name}(1)) & "|" & TypeName(ws.{name}(9))' for name in ADDS],
    *[f'Set h = {ADDS[name]}: s2 = ws.{name}("None").Name' for name in ADDS], *[f'Set h = {ADDS[name]}: s2 = ws.{name}(0).Name' for name in ADDS],
    *[f'Set h = {ADDS[name]}: s2 = TypeName(ws.{name}(Array(1))) & "|" & ws.{name}(Array(1)).Count' for name in ADDS],
    *[f'Set h = {ADDS[name]}: s2 = "": For Each o In ws.{name}: s2 = s2 & TypeName(o) & o.Name: Next' for name in ADDS],
    *[f'Set h = {ADDS[name]}: s2 = TypeName(ws.{name}.Parent) & "|" & ws.{name}.Application.Name & "|" & ws.{name}.Creator' for name in ADDS],
    # A collection reading and setting across its items.
    *[f'Set h = {ADDS[name]}: Set o = {ADDS[name]}: ws.{name}.Left = 200: s2 = h.Left & "|" & o.Left & "|" & ws.{name}.Left' for name in ADDS],
    *[f'Set h = {ADDS[name]}: Set o = {ADDS[name]}: ws.{name}.Delete: s2 = ws.{name}.Count' for name in ADDS], *[f'Set h = {ADDS[name]}: ws.{name}.Select: s2 = TypeName(Selection)' for name in ADDS],
    # Each member set on an item of each collection.
    *[f'Set h = {ADDS[name]}: {one}' for name in ADDS for one in SETS],
    # Text and the shape's own view of it.
    'Set h = ws.TextBoxes.Add(10, 10, 100, 40): h.Text = "Hello World": h.Characters(1, 5).Font.Bold = True: s2 = h.Characters(7, 5).Text & "|" & h.ShapeRange.TextFrame2.TextRange.Runs.Count',
    'Set h = ws.TextBoxes.Add(10, 10, 100, 40): h.Text = "ab": h.Font.Bold = True: h.Text = "cd"', 'Set h = ws.TextBoxes.Add(10, 10, 100, 40): ws.Shapes(1).TextFrame.Characters.Text = "ab": s2 = h.Text', 'Set h = ws.Rectangles.Add(10, 10, 100, 40): h.Text = "ab": s2 = ws.Shapes(1).TextFrame2.TextRange.Text',
    'Set h = ws.TextBoxes.Add(10, 10, 100, 40): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name & "|" & ws.Shapes(1).AutoShapeType', 'Set h = ws.Rectangles.Add(10, 10, 100, 40): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name & "|" & ws.Shapes(1).AutoShapeType',
    'Set h = ws.Ovals.Add(10, 10, 100, 40): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name & "|" & ws.Shapes(1).AutoShapeType', 'Set h = ws.Lines.Add(10, 10, 110, 50): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name & "|" & ws.Shapes(1).Connector',
    'Set h = ws.Arcs.Add(10, 10, 110, 50): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name & "|" & ws.Shapes(1).AutoShapeType', 'Set h = ws.Pictures.Insert(folder & "\\p.png"): s2 = ws.Shapes(1).Type & "|" & ws.Shapes(1).Name',
    # Held after the shape has gone.
    *[f'Set h = {ADDS[name]}: h.Delete: s2 = TypeName(h): v = h.Name: s2 = s2 & "|" & Err.Number: Err.Clear: h.Left = 5: s2 = s2 & "|" & Err.Number: Err.Clear' for name in ADDS],
    *[f'Set h = {ADDS[name]}: ws.Shapes(1).Delete: v = h.Name: s2 = Err.Number: Err.Clear' for name in ADDS],
    # DrawingObjects and the shape range's reroute.
    'Set h = ws.Rectangles.Add(10, 10, 100, 40): s2 = TypeName(ws.DrawingObjects(1)) & "|" & TypeName(ws.DrawingObjects) & "|" & ws.DrawingObjects.Count', 'Set h = ws.Rectangles.Add(10, 10, 100, 40): Set o = ws.Ovals.Add(1, 1, 5, 5): ws.DrawingObjects.Delete: s2 = ws.Shapes.Count',
    'Set h = ws.Rectangles.Add(10, 10, 100, 40): Set o = ws.Ovals.Add(1, 1, 5, 5): ws.DrawingObjects.Left = 7: s2 = h.Left & "|" & o.Left', 'Set h = ws.Rectangles.Add(10, 10, 100, 40): Set o = ws.Ovals.Add(1, 1, 5, 5): ws.DrawingObjects.Select: s2 = TypeName(Selection)',
    'Set h = ws.Rectangles.Add(10, 10, 100, 40): h.Select: s2 = TypeName(Selection) & "|" & Selection.Name', 'Set h = ws.TextBoxes.Add(10, 10, 100, 40): h.Select: s2 = TypeName(Selection) & "|" & Selection.Name',
    'Set h = ws.Rectangles.Add(10, 10, 100, 40): s2 = TypeName(ws.Shapes(1).DrawingObject) & "|" & ws.Shapes(1).DrawingObject.Name', 'Set h = ws.TextBoxes.Add(10, 10, 100, 40): s2 = TypeName(ws.Shapes(1).DrawingObject)', 'Set h = ws.Pictures.Insert(folder & "\\p.png"): s2 = TypeName(ws.Shapes(1).DrawingObject)',
    'Set h = ws.Lines.Add(10, 10, 110, 50): s2 = TypeName(ws.Shapes(1).DrawingObject)', 'Set h = ws.Arcs.Add(10, 10, 110, 50): s2 = TypeName(ws.Shapes(1).DrawingObject)', 'Set h = ws.Ovals.Add(10, 10, 100, 40): s2 = TypeName(ws.Shapes(1).DrawingObject)',
    'Set h = ws.Shapes.AddShape(1, 1, 1, 9, 9): s2 = TypeName(h.TextFrame.Characters.Parent) & "|" & h.TextFrame.Characters.Parent.Name', 'Set h = ws.Shapes.AddTextbox(1, 1, 1, 9, 9): s2 = TypeName(h.TextFrame.Characters.Parent)', 'Set h = ws.Shapes.AddShape(9, 1, 1, 9, 9): s2 = TypeName(h.TextFrame.Characters.Parent)',
    'Set h = ws.Shapes.AddShape(7, 1, 1, 9, 9): s2 = TypeName(h.TextFrame.Characters.Parent) & "|" & TypeName(h.DrawingObject)', 'Set h = ws.Shapes.AddShape(5, 1, 1, 9, 9): s2 = TypeName(h.TextFrame.Characters.Parent) & "|" & TypeName(h.DrawingObject)',
]


def _probe() -> str:
    lines: list[str] = []
    size = 40
    chunks = [range(at, min(at + size, len(CASES))) for at in range(0, len(CASES), size)]
    # Each in a procedure of its own, as one holding them all is too large to compile, taking its own error: passed up through a
    # call, Excel's -2147024809 turns into VBA's 5.
    for number, chunk in enumerate(chunks):
        lines += [f'Private Sub Act{number}(ByVal which As Long, ByVal ws As Worksheet, ByVal folder As String, h As Object, o As Object, s2 As Variant, failure As Long)', '    Dim v As Variant',
                  '    On Error GoTo Failed', '    Select Case which']
        for index in chunk:
            statements = CASES[index]
            lines += [f'    Case {index}']
            if 'Err.Clear' in statements:
                lines += ['        On Error Resume Next']
            lines += [f'        {statements}']
        lines += ['    End Select', '    Exit Sub', 'Failed:', '    failure = Err.Number', 'End Sub', '']
    # The reads split likewise.
    reads = [READS[at:at + 40] for at in range(0, len(READS), 40)]
    for number, part in enumerate(reads):
        lines += [f'Private Function Read{number}(ByVal ws As Worksheet, h As Object, s2 As Variant) As String', '    Dim s As String, v As String', '    On Error Resume Next']
        for read in part:
            lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "~" & v']
        lines += [f'    Read{number} = s', 'End Function', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, h As Object, o As Object, s As String, s2 As Variant, failure As Long', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)', f'    Select Case which \\ {size}']
    for number in range(len(chunks)):
        lines += [f'    Case {number}: Act{number} which, ws, folder, h, o, s2, failure']
    lines += ['    End Select', '    s = IIf(failure = 0, "ok", "Err " & failure)']
    lines += [f'    s = s & Read{number}(ws, h, s2)' for number in range(len(reads))]
    lines += ['    Tried = s', '    Application.CutCopyMode = False', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\0\xff\0\0\xff\0\0' * 2)) + chunk(b'IEND', b'')


def main() -> None:
    source = _probe()
    # The whole probe outlasts one sitting, so a run carries on from the cases an earlier one left beside the fixture, stopping after a while.
    partial = OUT.with_suffix('.part')
    done: list[list[object]] = json.loads(partial.read_text(encoding='utf-8')) if partial.exists() else []
    started = time.monotonic()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        (Path(folder) / 'p.png').write_bytes(png())
        excel.new_document()
        while len(done) < len(CASES) and time.monotonic() - started < 400:
            first, last = len(done), min(len(done) + 9, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=150.0)
            assert result.ok and not result.dialogs, (first, CASES[first:last + 1], result)
            for index, answer in zip(range(first, last + 1), str(result.value).split('@@@'), strict=True):
                with zipfile.ZipFile(Path(folder) / f'c{index}.xlsx') as package:
                    names = package.namelist()
                    xml = package.read('xl/drawings/drawing1.xml').decode('utf-8') if 'xl/drawings/drawing1.xml' in names else ''
                    if CASES[index] == f'Set h = {ADDS["Pictures"]}':
                        # As bytes: text mode would turn the part's line break into two.
                        kept = package.read('xl/drawings/_rels/drawing1.xml.rels').decode('utf-8').replace(folder, '<folder>')
                        (partial.with_suffix('.rels')).write_bytes(kept.encode('utf-8'))
                shapes = re.findall(r'<xdr:(?:sp|cxnSp|pic)\b.*?</xdr:(?:sp|cxnSp|pic)>', xml, re.DOTALL)
                done.append([CASES[index], answer.replace(folder, '<folder>'), [one.replace(folder, '<folder>').replace(folder.replace('\\', '/'), '<folder>') for one in shapes]])
            partial.write_text(json.dumps(done), encoding='utf-8')
    print(len(done), 'of', len(CASES))
    if len(done) == len(CASES):
        relationships = partial.with_suffix('.rels')
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source, 'picture_relationships': relationships.read_bytes().decode('utf-8'), 'cases': done}, indent=1) + '\n', encoding='utf-8')
        partial.unlink()
        relationships.unlink()


if __name__ == '__main__':
    main()
