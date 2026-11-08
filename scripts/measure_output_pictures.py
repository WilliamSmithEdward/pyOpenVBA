"""Chart.Export and CopyPicture as far as the model answers them itself: argument errors, path rules, what each returns and what
lands on the clipboard, typed and late.

Safe to run: every file goes under a temporary folder, which is also the current folder and Application.DefaultFilePath for the
row, so an omitted or a relative name lands there; Interactive:=True runs last, as it may put up a form. The clipboard is emptied
before each row, so a paste after a failed CopyPicture finds nothing of an earlier row. A row whose call would not compile typed
(an argument left out, one Excel does not have) runs late only, as a typed compile error stops the harness.

Rival rules: a bad argument is refused before anything is written, or after; a name with no folder goes to the current folder, or
to the default file path; the filter follows the name's extension, or FilterName; Export returns True or nothing; CopyPicture puts a
picture on the clipboard, or an object a paste makes a shape of; a chart in a frame exports as a chart sheet does, or not at all.

    python scripts/measure_output_pictures.py

writes tests/fixtures/output_pictures.json and the workbook the rows open, output_pictures.xlsx. It keeps its place in a file under TEMP, so a run that stops can go on.
"""
from __future__ import annotations

import ctypes
import json
import stat
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/output_pictures.json'
#: The workbook every row opens, as Excel saved it, for the model's replay.
BOOK = ROOT / 'tests/fixtures/output_pictures.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'output_pictures_partial.json'
#: Stop short of the tool's limit; a second run carries on.
BUDGET = 520.0
#: Each owner: how the row gets it, and the class it is declared as typed.
OWNERS = {
    'Chart': ('wb.Charts("Plot")', 'Chart'),
    'Frame': ('wb.Worksheets("One").ChartObjects(1).Chart', 'Chart'),
    'Range': ('wb.Worksheets("One").Range("A1:B3")', 'Range'),
    'Shape': ('wb.Worksheets("One").Shapes("Box")', 'Shape'),
    'ChartObject': ('wb.Worksheets("One").ChartObjects(1)', 'ChartObject'),
}
EXPORTERS = ('Chart', 'Frame')
PICTURES = ('Chart', 'Frame', 'Range', 'Shape', 'ChartObject')
A = 'folder & "\\a.png"'
EXPORT_ROWS: dict[str, str] = {
    'base': f'r = o.Export({A})', 'gif': 'r = o.Export(folder & "\\a.gif")', 'jpg': 'r = o.Export(folder & "\\a.jpg")',
    'jpeg': 'r = o.Export(folder & "\\a.jpeg")', 'bmp': 'r = o.Export(folder & "\\a.bmp")', 'tif': 'r = o.Export(folder & "\\a.tif")',
    'no_extension': 'r = o.Export(folder & "\\a")', 'pdf': 'r = o.Export(folder & "\\a.pdf")', 'upper': 'r = o.Export(folder & "\\A.PNG")',
    'statement': f'o.Export {A}',
    'filter_png': f'r = o.Export({A}, "PNG")', 'filter_gif': f'r = o.Export({A}, "GIF")', 'filter_jpg': f'r = o.Export({A}, "JPG")',
    'filter_jpeg': f'r = o.Export({A}, "JPEG")', 'filter_bmp': f'r = o.Export({A}, "BMP")', 'filter_lower': f'r = o.Export({A}, "png")',
    'filter_bogus': f'r = o.Export({A}, "bogus")', 'filter_empty': f'r = o.Export({A}, "")', 'filter_null': f'r = o.Export({A}, Null)',
    'filter_number': f'r = o.Export({A}, 5)', 'filter_named': f'r = o.Export(Filename:={A}, FilterName:="GIF")',
    'interactive_false': f'r = o.Export({A}, "PNG", False)',
    'name_empty': 'r = o.Export("")', 'name_null': 'r = o.Export(Null)', 'name_number': 'r = o.Export(5)',
    'name_relative': 'r = o.Export("rel.png")', 'name_relative_folder': 'r = o.Export("sub\\rel.png")',
    'name_dotdot': 'r = o.Export("sub\\..\\up.png")', 'name_no_folder': 'r = o.Export(folder & "\\nothere\\a.png")',
    'name_is_folder': 'r = o.Export(folder)', 'name_read_only': 'r = o.Export(folder & "\\ro.png")',
    'name_bad_character': 'r = o.Export(folder & "\\a?b.png")', 'name_long': 'r = o.Export(folder & "\\" & String(300, "a") & ".png")',
    'name_existing': 'r = o.Export(folder & "\\there.png")',
}
#: Late only: typed, these do not compile.
EXPORT_LATE = {'name_omitted': 'r = o.Export()', 'bogus': f'r = o.Export({A}, Bogus:=1)'}
#: Last, as it may put up a form.
EXPORT_LAST = {'interactive_true': f'r = o.Export({A}, "PNG", True)'}
PICTURE_ROWS: dict[str, str] = {
    'none': 'o.CopyPicture', 'screen_picture': 'o.CopyPicture xlScreen, xlPicture', 'printer_picture': 'o.CopyPicture xlPrinter, xlPicture',
    'screen_bitmap': 'o.CopyPicture xlScreen, xlBitmap', 'printer_bitmap': 'o.CopyPicture xlPrinter, xlBitmap',
    'named': 'o.CopyPicture Appearance:=xlScreen, Format:=xlBitmap',
    'appearance_3': 'o.CopyPicture 3, xlPicture', 'appearance_0': 'o.CopyPicture 0, xlPicture', 'appearance_negative': 'o.CopyPicture -1, xlPicture',
    'appearance_text': 'o.CopyPicture "x", xlPicture', 'appearance_digits': 'o.CopyPicture "2", xlPicture', 'appearance_null': 'o.CopyPicture Null, xlPicture',
    'appearance_true': 'o.CopyPicture True, xlPicture',
    'format_3': 'o.CopyPicture xlScreen, 3', 'format_0': 'o.CopyPicture xlScreen, 0', 'format_negative': 'o.CopyPicture xlScreen, -1',
    'format_text': 'o.CopyPicture xlScreen, "x"', 'format_null': 'o.CopyPicture xlScreen, Null', 'format_true': 'o.CopyPicture xlScreen, True',
}
#: A chart's own third argument, Size; typed on the others it does not compile.
CHART_SIZE = {'size_screen': 'o.CopyPicture xlScreen, xlPicture, xlScreen', 'size_printer': 'o.CopyPicture xlScreen, xlPicture, xlPrinter',
              'size_3': 'o.CopyPicture xlScreen, xlPicture, 3', 'size_text': 'o.CopyPicture xlScreen, xlPicture, "x"'}
PICTURE_LATE = {'returns': 'r = o.CopyPicture()', 'bogus': 'o.CopyPicture Bogus:=1'}
#: Rows outside the owners' table, each typed through its own declaration.
STATES = {
    'export_empty_chart': ('Chart', 'wb.Charts("Empty")', 'r = o.Export(folder & "\\a.png")'),
    'export_hidden_chart': ('Chart', 'wb.Charts("Plot")', 'o.Visible = xlSheetHidden\n    r = o.Export(folder & "\\a.png")'),
    'export_frame_hidden_sheet': ('Chart', 'wb.Worksheets("One").ChartObjects(1).Chart', 'wb.Worksheets("One").Visible = xlSheetHidden\n    r = o.Export(folder & "\\a.png")'),
    'export_saved_flag': ('Chart', 'wb.Charts("Plot")', 'wb.Saved = True\n    r = o.Export(folder & "\\a.png")\n    s = "saved=" & wb.Saved'),
    'export_not_active': ('Chart', 'wb.Charts("Plot")', 'wb.Worksheets("Two").Activate\n    r = o.Export(folder & "\\a.png")\n    s = "active=" & wb.ActiveSheet.Name'),
    'picture_range_hidden_rows': ('Range', 'wb.Worksheets("One").Range("A1:B3")', 'wb.Worksheets("One").Rows("1:3").Hidden = True\n    o.CopyPicture'),
    'picture_range_empty': ('Range', 'wb.Worksheets("One").Range("D10:E12")', 'o.CopyPicture'),
    'picture_range_column': ('Range', 'wb.Worksheets("One").Range("A:A")', 'o.CopyPicture'),
    'picture_range_areas': ('Range', 'wb.Worksheets("One").Range("A1,B3")', 'o.CopyPicture'),
    'picture_range_inactive_sheet': ('Range', 'wb.Worksheets("Three").Range("A1:B2")', 'o.CopyPicture'),
    'picture_shape_hidden': ('Shape', 'wb.Worksheets("One").Shapes("Box")', 'o.Visible = msoFalse\n    o.CopyPicture'),
    'picture_hidden_chart': ('Chart', 'wb.Charts("Plot")', 'o.Visible = xlSheetHidden\n    o.CopyPicture'),
    'picture_saved_flag': ('Range', 'wb.Worksheets("One").Range("A1:B3")', 'wb.Saved = True\n    o.CopyPicture\n    s = "saved=" & wb.Saved'),
}
#: What a paste of the clipboard makes, read after the row.
PASTE = '''    wb.Worksheets("Two").Visible = xlSheetVisible
    wb.Worksheets("Two").Activate
    n = wb.Worksheets("Two").Shapes.Count
    wb.Worksheets("Two").Paste
    p = Err.Number
    Err.Clear
    kind = TypeName(Selection)
    s = s & "|paste=" & p & ":" & (wb.Worksheets("Two").Shapes.Count - n) & ":" & kind
    If wb.Worksheets("Two").Shapes.Count > n Then s = s & ":" & wb.Worksheets("Two").Shapes(n + 1).Type
'''
MAKER = r'''Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, one As Worksheet, r As Long, c As Chart
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set one = wb.Worksheets(1)
    one.Name = "One"
    wb.Worksheets.Add(After:=one).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets("Two")).Name = "Three"
    wb.Worksheets("Three").Range("A1:B2").Value = 7
    For r = 1 To 5
        one.Cells(r, 1).Value = r
        one.Cells(r, 2).Value = r * r
    Next
    one.ChartObjects.Add(150, 10, 240, 160).Chart.SetSourceData one.Range("A1:B5")
    one.Shapes.AddShape(msoShapeRectangle, 20, 120, 60, 30).Name = "Box"
    wb.Worksheets("Two").Activate
    Set c = wb.Charts.Add(After:=wb.Worksheets("Three"))
    c.Name = "Empty"
    Set c = wb.Charts.Add(After:=wb.Worksheets("Three"))
    c.SetSourceData one.Range("A1:B5")
    c.Name = "Plot"
    one.Activate
    wb.SaveAs path, xlOpenXMLWorkbook
    Make = "StandardHeight=" & one.StandardHeight
    wb.Close False
End Function
'''


def row_source(call: str, declared: str, get: str, *, live: bool = True) -> str:
    """One row in a module of its own, so a compile error stops only that row. Not ``live``, for the model, the row reads no paste, as the model pastes no
    picture, and leaves DefaultFilePath and the drive alone, which the model does not hold: the current folder is what a relative name goes by (measured
    apart).
    """
    tail = PASTE if live else ""
    keep = "    there = Application.DefaultFilePath\n" if live else ""
    point = "    Application.DefaultFilePath = folder\n" if live else ""
    back = "    Application.DefaultFilePath = there\n" if live else ""
    drive = "    ChDrive Left(folder, 1)\n" if live else ""
    drive_back = "    ChDrive Left(was, 1)\n" if live else ""
    body = ''.join(f'    {line}\n' for line in call.split('\n'))
    return f'''Public Function Row(ByVal path As String, ByVal folder As String) As String
    Dim wb As Workbook, o As {declared}, s As String, r As Variant, was As String, there As String, e As Long, d As String
    Dim n As Long, p As Long, mode As Long, kind As String
    Application.DisplayAlerts = False
    was = CurDir
{keep}{drive}
    ChDir folder
{point}    Set wb = Workbooks.Open(path, 0)
    Set o = {get}
    On Error Resume Next
{body}    e = Err.Number
    d = Err.Description
    Err.Clear
    s = s & "|r=" & TypeName(r) & ":" & CStr(r)
    mode = Application.CutCopyMode
    s = s & "|mode=" & mode
{tail}    On Error GoTo 0
    wb.Close False
{back}{drive_back}
    ChDir was
    Row = "E" & e & "~" & Left(d, 120) & "~" & s
End Function
'''


def cases() -> list[dict[str, str]]:
    found: list[dict[str, str]] = []

    def add(member: str, owner: str, binding: str, row: str, call: str) -> None:
        found.append({'member': member, 'owner': owner, 'binding': binding, 'row': row, 'call': call})

    for owner in EXPORTERS:
        for binding in ('Typed', 'Late'):
            for row, call in EXPORT_ROWS.items():
                add('Export', owner, binding, row, call)
        for row, call in EXPORT_LATE.items():
            add('Export', owner, 'Late', row, call)
    for owner in PICTURES:
        for binding in ('Typed', 'Late'):
            for row, call in PICTURE_ROWS.items():
                add('CopyPicture', owner, binding, row, call)
            for row, call in CHART_SIZE.items():
                if binding == 'Late' or owner in ('Chart', 'Frame'):
                    add('CopyPicture', owner, binding, row, call)
        for row, call in PICTURE_LATE.items():
            add('CopyPicture', owner, 'Late', row, call)
    for row, (declared, _get, call) in STATES.items():
        for binding in ('Typed', 'Late'):
            add('state', declared, binding, row, call)
    for owner in EXPORTERS:
        for binding in ('Typed', 'Late'):
            for row, call in EXPORT_LAST.items():
                add('Export', owner, binding, row, call)
    return found


def _empty_clipboard() -> None:
    user32 = ctypes.windll.user32  # pyright: ignore[reportAttributeAccessIssue, reportUnknownMemberType, reportUnknownVariableType]
    for _ in range(20):
        if user32.OpenClipboard(None):  # pyright: ignore[reportUnknownMemberType]
            user32.EmptyClipboard()  # pyright: ignore[reportUnknownMemberType]
            user32.CloseClipboard()  # pyright: ignore[reportUnknownMemberType]
            return
        time.sleep(0.1)
    raise SystemExit('the clipboard stayed held')


def _kind(data: bytes) -> str:
    """What a file is by its first bytes, whatever its name says."""
    for magic, kind in ((b'\x89PNG', 'png'), (b'GIF8', 'gif'), (b'\xff\xd8', 'jpeg'), (b'BM', 'bmp'), (b'II*\x00', 'tiff'), (b'MM\x00*', 'tiff'), (b'%PDF', 'pdf')):
        if data.startswith(magic):
            return kind
    return 'before' if data == b'before' else f'other:{data[:4].hex()}'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    budget = float(sys.argv[1]) if len(sys.argv) > 1 else BUDGET
    all_cases = cases()
    done: dict[str, dict[str, object]] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    attempts = 0
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        book = home / 'book.xlsx'
        made = session.run_vba(MAKER, 'Make', (str(book),), timeout=60.0, line_numbers=False)
        assert made.ok, made
        print('made', made.value, flush=True)
        from fixture_workbook import copy_saved

        copy_saved(book, BOOK)
        for case in all_cases:
            key = f'{case["member"]} {case["owner"]} {case["binding"]} {case["row"]}'
            if key in done:
                continue
            if time.monotonic() - started > budget:
                print('stopping for now', len(done), 'of', len(all_cases), flush=True)
                break
            # A new folder for each try: one a stopped Excel held may not be clean.
            attempts += 1
            folder = home / f'run{attempts}'
            (folder / 'sub').mkdir(parents=True)
            (folder / 'there.png').write_bytes(b'before')
            (folder / 'ro.png').write_bytes(b'before')
            (folder / 'ro.png').chmod(stat.S_IREAD)
            owner = case['owner']
            get, declared = OWNERS[owner] if case['member'] != 'state' else (STATES[case['row']][1], owner)
            source = row_source(case['call'], declared if case['binding'] == 'Typed' else 'Object', get)
            _empty_clipboard()
            try:
                ran = session.run_vba(source, 'Row', (str(book), str(folder)), timeout=45.0, line_numbers=False)
            except HarnessError as failed:
                print('restarting', str(failed)[:80], flush=True)
                session.close()
                time.sleep(5)
                session = ExcelSession(HarnessConfig(exclusive=False))
                session.new_document()
                continue
            answer = str(ran.value) if ran.ok else f'failed {ran.outcome} {[one.message[:110] for one in ran.dialogs]}'
            answer = answer.replace(str(folder), '<DIR>').replace(held, '<HOME>')
            files = sorted(f'{str(one.relative_to(folder)).replace(chr(92), "/")} {_kind(one.read_bytes())}' for one in folder.rglob('*')
                           if one.is_file() and one.name != 'ro.png' and not (one.name == 'there.png' and one.read_bytes() == b'before'))
            done[key] = {**case, 'excel': answer, 'files': files}
            print(key, answer[:110], files[:3], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
            (folder / 'ro.png').chmod(stat.S_IWRITE)
    session.close()
    if len(done) == len(all_cases):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': list(done.values())}, indent=0) + '\n', encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(done), 'cases', flush=True)


if __name__ == '__main__':
    main()
