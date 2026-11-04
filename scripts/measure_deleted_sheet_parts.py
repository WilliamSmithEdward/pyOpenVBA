"""What a deleted sheet leaves in the file.

    python scripts/measure_deleted_sheet_parts.py

Excel makes a workbook whose sheet Gone holds a marker in a cell of each kind and in every object a sheet can have (tests/fixtures/deleted_sheet_parts.xlsx). Excel
and the model each delete the sheet and save. The record has the parts of each file, and which parts of each still hold a marker.

Writes tests/fixtures/deleted_sheet_parts.json. tests/test_excel_deleted_sheet_parts.py replays it.
"""
from __future__ import annotations

import json
import re
import struct
import sys
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
sys.path.insert(0, str(ROOT / 'scripts'))
from fixture_workbook import copy_saved  # noqa: E402
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402

OUT = ROOT / 'tests/fixtures/deleted_sheet_parts.json'
BOOK = ROOT / 'tests/fixtures/deleted_sheet_parts.xlsx'
MARKERS = ('GONETEXT', 'GONENOTE', 'GONESHAPE', 'GONETABLE', 'GONELINK', 'GONEFORMULA', 'GONEPIVOT', 'GONERECORD', 'GONESHEET', '271828')


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(b'\x00\x00\x00\x00')) + chunk(b'IEND', b'')


MAKE = r'''Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, g As Worksheet, k As Worksheet, d As Worksheet, c As ChartObject
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add
    Set k = wb.Worksheets(1): k.Name = "Keep"
    Set g = wb.Worksheets.Add(After:=k): g.Name = "GONESHEET"
    Set d = wb.Worksheets.Add(After:=g): d.Name = "Data"
    k.Range("A1").Value = 1: k.Range("A2").Value = 2: k.Range("A3").Value = 3: k.Range("B1").Value = "kept text"
    Set c = k.ChartObjects.Add(10, 100, 200, 120)
    c.Chart.SetSourceData k.Range("A1:A3")
    g.Range("A1").Value = 271828: g.Range("A2").Value = 2: g.Range("A3").Value = 3: g.Range("B1").Value = "GONETEXT"
    g.Range("B2").Formula = "=""GONEFORMULA"""
    Set c = g.ChartObjects.Add(10, 100, 200, 120)
    c.Chart.SetSourceData g.Range("A1:A3")
    g.Shapes.AddPicture folder & "\p.png", False, True, 300, 10, 20, 20
    g.Shapes.AddShape(1, 300, 50, 40, 40).Name = "GONESHAPE"
    g.Range("C1").AddComment "GONENOTE"
    g.Range("E1").Value = "x": g.Range("E2").Value = 1
    g.ListObjects.Add(xlSrcRange, g.Range("E1:E2"), , xlYes).Name = "GONETABLE"
    g.Hyperlinks.Add g.Range("F1"), "http://example.com/GONELINK"
    g.Buttons.Add 400, 10, 40, 20
    d.Range("A1").Value = "h": d.Range("A2").Value = "GONERECORD": d.Range("B1").Value = "i": d.Range("B2").Value = 2
    wb.PivotCaches.Create(xlDatabase, d.Range("A1:B2")).CreatePivotTable g.Range("H3"), "GONEPIVOT"
    wb.SaveAs folder & "\book.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
Public Function Edited(ByVal path As String, ByVal out As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path)
    wb.Worksheets("GONESHEET").Delete
    wb.Worksheets("Data").Range("A2").Value = 7
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
Public Function Reads(ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wb = Workbooks.Open(path)
    Reads = Err.Number & " " & TypeName(wb)
    If Not wb Is Nothing Then
        Reads = Reads & " sheets " & wb.Worksheets.Count & " charts " & wb.Worksheets("Keep").ChartObjects.Count & " caches " & wb.PivotCaches.Count & " b1 " & wb.Worksheets("Keep").Range("B1").Value
        wb.Close False
    End If
    Application.DisplayAlerts = True
End Function
'''
EDIT = 'Public Sub Edit()\n    Application.DisplayAlerts = False\n    ActiveWorkbook.Worksheets("GONESHEET").Delete\n    ActiveWorkbook.Worksheets("Data").Range("A2").Value = 7\nEnd Sub\n'


def markers(path: Path) -> dict[str, list[str]]:
    """Each marker against the parts still holding it."""
    with zipfile.ZipFile(path) as package:
        found: dict[str, list[str]] = {}
        for name in package.namelist():
            data = package.read(name)
            for marker in MARKERS:
                if marker.encode() in data:
                    found.setdefault(marker, []).append(name)
        return found


def parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return sorted(package.namelist())


def workbook_tags(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        relationships = package.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        workbook = package.read('xl/workbook.xml').decode('utf-8')
    return {'relationships': sorted(re.findall(r'\bTarget="([^"]*)"', relationships)), 'pivot_caches': re.findall(r'<pivotCache\b[^>]*/>', workbook)}


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        (base / 'p.png').write_bytes(png())
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (folder,), timeout=60.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(base / 'book.xlsx', BOOK)
            done = excel.run_vba(MAKE, 'Edited', (str(BOOK), str(base / 'own.xlsx')), timeout=60.0, line_numbers=False)
            assert done.ok and done.value == 'ok', done
            app = ExcelApplication.open(BOOK, with_vba=False)
            app.add_module(EDIT, name='E')
            app.run('E.Edit')
            app.save(base / 'model.xlsx')
            record = {
                'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'edit': EDIT, 'markers': MARKERS,
                'start': {'parts': parts(BOOK), 'markers': markers(BOOK)},
                'excel': {'parts': parts(base / 'own.xlsx'), 'markers': markers(base / 'own.xlsx'), 'workbook': workbook_tags(base / 'own.xlsx'),
                          'reads': str(excel.run_vba(MAKE, 'Reads', (str(base / 'own.xlsx'),), timeout=60.0, line_numbers=False).value)},
                'model': {'parts': parts(base / 'model.xlsx'), 'markers': markers(base / 'model.xlsx'), 'workbook': workbook_tags(base / 'model.xlsx'),
                          'reads': str(excel.run_vba(MAKE, 'Reads', (str(base / 'model.xlsx'),), timeout=60.0, line_numbers=False).value)},
            }
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    for who in ('excel', 'model'):
        print(who, record[who]['reads'], len(record[who]['parts']))
        print('   markers', record[who]['markers'])
        print('   workbook', record[who]['workbook'])
    print('only model', sorted(set(record['model']['parts']) - set(record['excel']['parts'])))
    print('only excel', sorted(set(record['excel']['parts']) - set(record['model']['parts'])))


if __name__ == '__main__':
    main()
