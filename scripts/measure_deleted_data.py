"""Whether what a macro deletes leaves the file: every part of the package searched for a marker the deleted thing held.

    python scripts/measure_deleted_data.py

Excel makes a workbook with a marker in each thing a macro can delete (tests/fixtures/deleted_data.xlsx, and deleted_data.xlsm for a sheet's VBA). For each deletion
Excel and the model each make it and save. The record has which parts of each file still hold each marker.

Writes tests/fixtures/deleted_data.json. tests/test_excel_deleted_data.py replays it.
"""
from __future__ import annotations

import json
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
from pyopenvba.excel import ExcelFile  # noqa: E402
from pyopenvba.exceptions import PyOpenVBAError, VBARuntimeError, VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/deleted_data.json'
BOOK = ROOT / 'tests/fixtures/deleted_data.xlsx'
MACRO_BOOK = ROOT / 'tests/fixtures/deleted_data.xlsm'


def png() -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 1, 1, 8, 2, 0, 0, 0)) + chunk(b'tEXt', b'Comment\x00MKIMAGE') + chunk(b'IDAT', zlib.compress(b'\x00\x00\x00\x00')) + chunk(b'IEND', b'')


MAKE = r'''Public Function Make(ByVal folder As String) As String
    Dim wb As Workbook, m As Worksheet, o As Worksheet, d As Worksheet, g As Worksheet, ext As Workbook, c As ChartObject
    Application.DisplayAlerts = False
    Set ext = Workbooks.Add
    ext.Worksheets(1).Name = "Sheet1": ext.Worksheets(1).Range("A1").Value = "MKEXT"
    ext.SaveAs folder & "\ext.xlsx", xlOpenXMLWorkbook
    Set wb = Workbooks.Add
    Do While wb.Worksheets.Count < 4: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Loop
    Set m = wb.Worksheets(1): m.Name = "Main"
    Set o = wb.Worksheets(2): o.Name = "Other"
    Set d = wb.Worksheets(3): d.Name = "Src"
    Set g = wb.Worksheets(4): g.Name = "MKSHEET"
    m.Range("A1").Value = "MKCLEAR"
    m.Range("A2").Value = "MKOVER"
    m.Range("A3").Value = "MKCELLCLEAR"
    m.Range("A5").Value = "MKROW"
    m.Range("A6").Value = 314159
    m.Range("A7").Formula = "=""MKFORMULA"""
    m.Range("A8").Value = "MKKEPT"
    m.Range("A9").Value = "MKTWICE"
    m.Range("B9").Value = "MKTWICE"
    m.Range("Z1").Value = "MKCOLUMN"
    m.Shapes.AddShape(1, 300, 10, 40, 40).Name = "MKSHAPE"
    m.Shapes("MKSHAPE").TextFrame.Characters.Text = "MKSHAPETEXT"
    m.Range("C1").AddComment "MKNOTE"
    m.Hyperlinks.Add m.Range("C2"), "http://example.com/MKLINK", , , "shown"
    m.Range("E1").Value = "MKTABLEHEAD": m.Range("E2").Value = 1
    m.ListObjects.Add(xlSrcRange, m.Range("E1:E2"), , xlYes).Name = "MKTABLE"
    m.Range("G1").Value = 1: m.Range("G2").Value = 2
    Set c = m.ChartObjects.Add(300, 100, 200, 120)
    c.Name = "MKCHART"
    c.Chart.SetSourceData m.Range("G1:G2")
    c.Chart.HasTitle = True: c.Chart.ChartTitle.Text = "MKCHARTTITLE"
    m.Shapes.AddPicture(folder & "\p.png", False, True, 300, 250, 20, 20).Name = "MKPIC"
    wb.Names.Add "MKNAME", "=""MKNAMEVALUE"""
    m.Range("K1").Formula = "=[ext.xlsx]Sheet1!$A$1"
    d.Range("A1").Value = "h": d.Range("B1").Value = "i": d.Range("A2").Value = "MKRECORD": d.Range("B2").Value = 2
    wb.PivotCaches.Create(xlDatabase, d.Range("A1:B2")).CreatePivotTable o.Range("B3"), "MKPIVOT"
    g.Range("A1").Value = "MKSHEETCELL"
    wb.BuiltinDocumentProperties("Title").Value = "MKTITLE"
    ext.Close False
    wb.SaveAs folder & "\book.xlsx", xlOpenXMLWorkbook
    wb.SaveAs folder & "\macro.xlsm", xlOpenXMLWorkbookMacroEnabled
    Dim one As Object
    For Each one In wb.VBProject.VBComponents
        If one.Type = 100 Then
            If one.Properties("Name").Value = "MKSHEET" Then one.CodeModule.AddFromString "Public Sub MKVBASUB()" & vbCrLf & "    Debug.Print ""MKVBATEXT""" & vbCrLf & "End Sub"
        End If
    Next
    wb.VBProject.VBComponents.Add(1).Name = "MKMODULE"
    wb.VBProject.VBComponents("MKMODULE").CodeModule.AddFromString "Public Sub MKMODSUB()" & vbCrLf & "End Sub"
    wb.Save
    wb.Close False
    Application.DisplayAlerts = True
    Make = "ok"
End Function
'''
EDITOR = r'''Public Function Edited(ByVal path As String, ByVal out As String, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Worksheet
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0): Set ws = wb.Worksheets("Main")
    {edit}
    Application.DisplayAlerts = False
    wb.SaveAs out, kind
    wb.Close False
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Edited = "ok"
End Function
'''
#: Each deletion and the markers it should take out of the file.
EDITS: dict[str, tuple[str, tuple[str, ...]]] = {
    'nothing, an unrelated value': ('ws.Range("AA1").Value = 1', ()),
    'ClearContents': ('ws.Range("A1").ClearContents', ('MKCLEAR',)),
    'Clear': ('ws.Range("A3").Clear', ('MKCELLCLEAR',)),
    'text overwritten': ('ws.Range("A2").Value = "x"', ('MKOVER',)),
    'text used twice, one cleared': ('ws.Range("A9").ClearContents', ()),
    'text used twice, both cleared': ('ws.Range("A9:B9").ClearContents', ('MKTWICE',)),
    'a number cleared': ('ws.Range("A6").ClearContents', ('314159',)),
    'a formula cleared': ('ws.Range("A7").ClearContents', ('MKFORMULA',)),
    'Rows.Delete': ('ws.Rows(5).Delete', ('MKROW',)),
    'Columns.Delete': ('ws.Columns(26).Delete', ('MKCOLUMN',)),
    'Shape.Delete': ('ws.Shapes("MKSHAPE").Delete', ('MKSHAPE', 'MKSHAPETEXT')),
    'Comment.Delete': ('ws.Range("C1").Comment.Delete', ('MKNOTE',)),
    'Hyperlink.Delete': ('ws.Range("C2").Hyperlinks(1).Delete', ('MKLINK',)),
    'ListObject.Delete': ('ws.ListObjects("MKTABLE").Delete', ('MKTABLE', 'MKTABLEHEAD')),
    'ListObject.Unlist': ('ws.ListObjects("MKTABLE").Unlist', ('MKTABLE',)),
    'ChartObject.Delete': ('ws.ChartObjects("MKCHART").Delete', ('MKCHART', 'MKCHARTTITLE')),
    'Picture Delete': ('ws.Shapes("MKPIC").Delete', ('MKPIC', 'MKIMAGE')),
    'PivotTable contents cleared': ('wb.Worksheets("Other").PivotTables("MKPIVOT").TableRange2.ClearContents', ('MKPIVOT',)),
    'Name.Delete': ('wb.Names("MKNAME").Delete', ('MKNAME', 'MKNAMEVALUE')),
    'PivotTable cleared': ('wb.Worksheets("Other").PivotTables("MKPIVOT").TableRange2.Clear', ('MKPIVOT', 'MKRECORD')),
    'BreakLink': ('wb.BreakLink wb.LinkSources(1)(1), xlLinkTypeExcelLinks', ('MKEXT',)),
    'linked formula cleared': ('ws.Range("K1").ClearContents', ()),
    'a sheet deleted': ('wb.Worksheets(4).Delete', ('MKSHEET', 'MKSHEETCELL')),
    'title emptied': ('wb.BuiltinDocumentProperties("Title").Value = ""', ('MKTITLE',)),
}
MACRO_EDITS: dict[str, tuple[str, tuple[str, ...]]] = {
    # By its number, as a macro of the model's holds the text it runs.
    'a sheet with VBA deleted': ('wb.Worksheets(4).Delete', ('MKSHEET', 'MKSHEETCELL', 'MKVBASUB', 'MKVBATEXT')),
    'nothing, a macro book': ('ws.Range("AA1").Value = 1', ()),
}
MARKERS = sorted({marker for edits in (EDITS, MACRO_EDITS) for _, markers in edits.values() for marker in markers} | {'MKMODSUB'})


def held(path: Path) -> dict[str, list[str]]:
    """Each marker against the parts holding it, every part of the package decompressed; a VBA project by its modules' source too."""
    found: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            data = package.read(name)
            for marker in MARKERS:
                if marker.encode() in data or marker.encode('utf-16-le') in data:
                    found.setdefault(marker, []).append(name)
        project = 'xl/vbaProject.bin' in package.namelist()
    if project:
        try:
            with ExcelFile(path) as host:
                sources = {module.name: module.source for module in host.vba_project().modules}
        except PyOpenVBAError as failure:
            found['!unread'] = [str(failure)[:120]]
        else:
            for module, source in sources.items():
                for marker in MARKERS:
                    if marker in source:
                        found.setdefault(marker, []).append(f'vba:{"sheet" if marker != "MKMODSUB" else module}')
    return found


def in_model(start: Path, out: Path, edit: str) -> str:
    app = ExcelApplication.open(start, with_vba=start.suffix == '.xlsm', sandbox=False)
    app.add_module(f'Public Sub Edit()\n    Dim wb As Object, ws As Object\n    Application.DisplayAlerts = False\n    Set wb = ActiveWorkbook: Set ws = wb.Worksheets("Main")\n    {edit}\nEnd Sub\n', name='MeasureEdit')
    try:
        app.run('MeasureEdit.Edit')
        app.save(out)
    except (VBAUnsupportedError, VBARuntimeError) as refused:
        return f'{type(refused).__name__} {str(refused)[:150]}'
    return ''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder:
        base = Path(folder)
        (base / 'p.png').write_bytes(png())
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            made = excel.run_vba(MAKE, 'Make', (folder,), timeout=90.0, line_numbers=False)
            assert made.ok and made.value == 'ok', made
            copy_saved(base / 'book.xlsx', BOOK)
            copy_saved(base / 'macro.xlsm', MACRO_BOOK)
            print('start', held(BOOK))
            rows: dict[str, dict[str, object]] = {}
            for number, (name, (edit, gone)) in enumerate({**EDITS, **MACRO_EDITS}.items()):
                macro = name in MACRO_EDITS
                suffix = '.xlsm' if macro else '.xlsx'
                start = MACRO_BOOK if macro else BOOK
                own = base / f'own{number}{suffix}'
                done = excel.run_vba(EDITOR.format(edit=edit), 'Edited', (str(start), str(own), 52 if macro else 51), timeout=60.0, line_numbers=False)
                row: dict[str, object] = {'edit': edit, 'macro': macro, 'should_go': list(gone)}
                if done.ok and done.value == 'ok':
                    row['excel'] = held(own)
                else:
                    row['excel_refuses'] = f'{done.outcome} {str(done.error)[:160]}'
                saved = base / f'model{number}{suffix}'
                refused = in_model(start, saved, edit)
                if refused:
                    row['model_refuses'] = refused
                else:
                    row['model'] = held(saved)
                rows[name] = row
                excel_left = {marker: row['excel'][marker] for marker in gone if marker in row['excel']} if 'excel' in row else row.get('excel_refuses')
                model_left = {marker: row['model'][marker] for marker in gone if marker in row['model']} if 'model' in row else refused
                print(name.ljust(32), 'EXCEL', str(excel_left)[:90].ljust(40), 'MODEL', str(model_left)[:150], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'make': MAKE, 'markers': MARKERS, 'start': held(BOOK), 'start_macro': held(MACRO_BOOK), 'edits': rows}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
