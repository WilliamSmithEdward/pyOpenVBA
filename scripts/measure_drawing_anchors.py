"""Where a shape in a file is: the anchor's cells or its own transform. Rival rules: Excel takes the transform when it is not empty
(what the model did), or the anchor whatever the transform says. Every shape's transform here is changed to one far from its anchor,
and the file's other anchor kinds are written by hand. The prediction of the anchor rule, with 15 point rows and 48 point columns:
  plain, one cell, absolute edit: left 10, top 25, width 106, height 45, none of them 777 or 888.
  a oneCellAnchor: from its cell, sized by its ext. An absoluteAnchor: its pos and ext.
A rotated shape's anchor: rival rules are that it is the box as it stands, or, from 45 to 135 degrees and 225 to 315, the box turned about its
centre, its width and height swapped, so reads give 100 by 50 at every angle.

    python scripts/measure_drawing_anchors.py

writes tests/fixtures/drawing_anchors.json, with the file in it, which tests/test_excel_shape_load.py replays.
"""
from __future__ import annotations

import base64
import io
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fixture_workbook import without_save_path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/drawing_anchors.json'
FAR = '<a:xfrm><a:off x="9867900" y="11277600"/><a:ext cx="9867900" cy="11277600"/></a:xfrm>'
MAKE = '''Public Sub Make(ByVal path As String)
    Dim wb As Workbook, ws As Worksheet, g As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Shapes.AddShape(1, 10, 25, 106, 45).Name = "Plain"
    ws.Shapes.AddShape(1, 10, 125, 106, 45).Name = "OneEdit"
    ws.Shapes("OneEdit").Placement = xlMove
    ws.Shapes.AddShape(1, 10, 225, 106, 45).Name = "AbsEdit"
    ws.Shapes("AbsEdit").Placement = xlFreeFloating
    ws.Shapes.AddShape(1, 10, 325, 106, 45).Name = "OneAnchor"
    ws.Shapes.AddShape(1, 10, 425, 106, 45).Name = "AbsAnchor"
    ws.Shapes.AddShape(1, 300, 25, 50, 50).Name = "M1"
    ws.Shapes.AddShape(1, 400, 125, 50, 50).Name = "M2"
    Dim angle As Variant
    For Each angle In Array(30, 44, 45, 90, 134, 135, 200, 224, 225, 270, 314, 315)
        With ws.Shapes.AddShape(1, 600, 25 + angle * 2, 100, 50)
            .Name = "R" & angle
            .Rotation = angle
        End With
    Next
    Set g = ws.Shapes.Range(Array("M1", "M2")).Group
    g.Name = "Grp"
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Sub
'''
READ = '''Public Function Tried(ByVal path As String) As String
    Dim wb As Workbook, one As Object, s As String, m As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    For Each one In wb.Worksheets(1).Shapes
        s = s & "|" & one.Name & " " & one.Left & " " & one.Top & " " & one.Width & " " & one.Height & " " & one.Placement & " " & one.Rotation
        If one.Type = msoGroup Then
            For Each m In one.GroupItems
                s = s & "|" & m.Name & " " & m.Left & " " & m.Top & " " & m.Width & " " & m.Height
            Next
        End If
    Next
    wb.Close False
    Tried = Mid$(s, 2)
End Function
'''


def _changed(data: bytes) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as saved, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as kept:
        for item in saved.infolist():
            body = saved.read(item.filename)
            if item.filename == 'xl/drawings/drawing1.xml':
                body = _drawing(body.decode('utf-8')).encode('utf-8')
            kept.writestr(item.filename, body)
    return out.getvalue()


def _drawing(xml: str) -> str:
    pieces = re.split(r'(?=<xdr:twoCellAnchor)', xml)
    result = []
    for piece in pieces:
        if 'name="Grp"' in piece:
            result.append(re.sub(r'<a:xfrm><a:off[^>]*/><a:ext[^>]*/><a:chOff', '<a:xfrm><a:off x="9867900" y="11277600"/><a:ext cx="9867900" cy="11277600"/><a:chOff', piece, count=1))
            continue
        rotated = re.search(r'name="R\d+"', piece)
        if rotated:
            piece = re.sub(r'<a:xfrm( [^>]*)?>.*?</a:xfrm>', lambda found: FAR.replace('<a:xfrm>', f'<a:xfrm{found.group(1) or ""}>'), piece, count=1)
        for name in ('Plain', 'OneEdit', 'AbsEdit', 'OneAnchor', 'AbsAnchor'):
            if f'name="{name}"' not in piece:
                continue
            piece = re.sub(r'<a:xfrm>.*?</a:xfrm>', FAR, piece, count=1)
            if name == 'OneAnchor':
                piece = re.sub(r'<xdr:twoCellAnchor[^>]*>', '<xdr:oneCellAnchor>', piece).replace('</xdr:twoCellAnchor>', '</xdr:oneCellAnchor>')
                piece = re.sub(r'<xdr:to>.*?</xdr:to>', '<xdr:ext cx="1270000" cy="508000"/>', piece)
            if name == 'AbsAnchor':
                piece = re.sub(r'<xdr:twoCellAnchor[^>]*>', '<xdr:absoluteAnchor>', piece).replace('</xdr:twoCellAnchor>', '</xdr:absoluteAnchor>')
                piece = re.sub(r'<xdr:from>.*?</xdr:from><xdr:to>.*?</xdr:to>', '<xdr:pos x="254000" y="508000"/><xdr:ext cx="1270000" cy="508000"/>', piece)
        result.append(piece)
    return ''.join(result)


def main() -> None:
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        path = Path(folder) / 'made.xlsx'
        excel.new_document()
        assert excel.run_vba(MAKE, 'Make', (str(path),), timeout=60.0).ok
        data = without_save_path(_changed(path.read_bytes()))
        path.write_bytes(data)
        result = excel.run_vba(READ, 'Tried', (str(path),), timeout=60.0)
        print(result.outcome, result.value, [one.message[:120] for one in result.dialogs], flush=True)
        assert result.ok, result
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'reader': READ, 'answer': str(result.value), 'file': base64.b64encode(data).decode('ascii')}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
