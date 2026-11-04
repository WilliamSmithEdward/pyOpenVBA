"""ExportAsFixedFormat and PrintOut as far as the model answers them itself: argument errors, path rules and the order against BeforePrint, typed and late.

Safe to run: every export passes OpenAfterPublish:=False, so no viewer opens, and every PrintOut passes PrintToFile:=True, so nothing goes to paper. Preview is
never passed. Every file goes under a temporary folder, which is also the current folder and Application.DefaultFilePath for the row, so an omitted or a relative
name lands there. A row that changes the active printer stops the run.

Rival rules: a bad argument is refused before anything is written, or after; a name with no folder goes to the current folder, or to the workbook's, or to the
default file path; an extension is added, or not; BeforePrint runs before the path is looked at, or after; Cancel stops the export, or not.

    python scripts/measure_output_members.py

writes tests/fixtures/output_members.json. It keeps its place in a file under TEMP, so a run that stops can go on.
"""
from __future__ import annotations

import json
import os
import stat
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/output_members.json'
PARTIAL = Path(tempfile.gettempdir()) / 'output_members_partial.json'
#: Stop short of the tool's limit; a second run carries on.
BUDGET = float(sys.argv[1]) if len(sys.argv) > 1 else 520.0
OWNERS = {'Worksheet': 'wb.Worksheets(1)', 'Workbook': 'wb', 'Range': 'wb.Worksheets(1).Range("A1:B3")', 'Chart': 'wb.Charts(1)', 'Sheets': 'wb.Sheets', 'Window': 'wb.Windows(1)'}
EXPORTS = ('Worksheet', 'Workbook', 'Range', 'Chart')
SAFE = 'OpenAfterPublish:=False'
TO = 'Filename:=folder & "\\x.pdf"'
EXPORT_ROWS = {
    'base': f'Type:=xlTypePDF, {TO}, {SAFE}', 'xps': f'Type:=xlTypeXPS, Filename:=folder & "\\x.xps", {SAFE}',
    'type_2': f'Type:=2, {TO}, {SAFE}', 'type_negative': f'Type:=-1, {TO}, {SAFE}', 'type_text': f'Type:="x", {TO}, {SAFE}', 'type_digit_text': f'Type:="0", {TO}, {SAFE}',
    'type_null': f'Type:=Null, {TO}, {SAFE}', 'type_true': f'Type:=True, {TO}, {SAFE}', 'type_fraction': f'Type:=0.5, {TO}, {SAFE}', 'type_empty': f'Type:=Empty, {TO}, {SAFE}',
    'type_omitted': f'{TO}, {SAFE}',
    'name_omitted': f'Type:=xlTypePDF, {SAFE}', 'name_empty': f'Type:=xlTypePDF, Filename:="", {SAFE}', 'name_null': f'Type:=xlTypePDF, Filename:=Null, {SAFE}',
    'name_number': f'Type:=xlTypePDF, Filename:=5, {SAFE}', 'name_true': f'Type:=xlTypePDF, Filename:=True, {SAFE}', 'name_relative': f'Type:=xlTypePDF, Filename:="rel.pdf", {SAFE}',
    'name_relative_folder': f'Type:=xlTypePDF, Filename:="sub\\rel.pdf", {SAFE}', 'name_dotdot': f'Type:=xlTypePDF, Filename:="sub\\..\\up.pdf", {SAFE}',
    'name_no_extension': f'Type:=xlTypePDF, Filename:=folder & "\\noext", {SAFE}', 'name_other_extension': f'Type:=xlTypePDF, Filename:=folder & "\\x.txt", {SAFE}',
    'name_xps_as_pdf': f'Type:=xlTypePDF, Filename:=folder & "\\x.xps", {SAFE}', 'name_no_folder': f'Type:=xlTypePDF, Filename:=folder & "\\nothere\\x.pdf", {SAFE}',
    'name_is_folder': f'Type:=xlTypePDF, Filename:=folder, {SAFE}', 'name_read_only': f'Type:=xlTypePDF, Filename:=folder & "\\ro.pdf", {SAFE}',
    'name_bad_character': f'Type:=xlTypePDF, Filename:=folder & "\\a?b.pdf", {SAFE}', 'name_long': f'Type:=xlTypePDF, Filename:=folder & "\\" & String(300, "a") & ".pdf", {SAFE}',
    'name_existing': f'Type:=xlTypePDF, Filename:=folder & "\\there.pdf", {SAFE}',
    'quality_0': f'Type:=xlTypePDF, {TO}, Quality:=0, {SAFE}', 'quality_1': f'Type:=xlTypePDF, {TO}, Quality:=1, {SAFE}', 'quality_2': f'Type:=xlTypePDF, {TO}, Quality:=2, {SAFE}',
    'quality_text': f'Type:=xlTypePDF, {TO}, Quality:="x", {SAFE}', 'quality_null': f'Type:=xlTypePDF, {TO}, Quality:=Null, {SAFE}',
    'props_text': f'Type:=xlTypePDF, {TO}, IncludeDocProperties:="x", {SAFE}', 'props_null': f'Type:=xlTypePDF, {TO}, IncludeDocProperties:=Null, {SAFE}',
    'areas_text': f'Type:=xlTypePDF, {TO}, IgnorePrintAreas:="x", {SAFE}', 'areas_true': f'Type:=xlTypePDF, {TO}, IgnorePrintAreas:=True, {SAFE}',
    'from_0': f'Type:=xlTypePDF, {TO}, From:=0, {SAFE}', 'from_negative': f'Type:=xlTypePDF, {TO}, From:=-1, {SAFE}', 'from_past': f'Type:=xlTypePDF, {TO}, From:=99, {SAFE}',
    'from_text': f'Type:=xlTypePDF, {TO}, From:="x", {SAFE}', 'from_null': f'Type:=xlTypePDF, {TO}, From:=Null, {SAFE}', 'to_0': f'Type:=xlTypePDF, {TO}, To:=0, {SAFE}',
    'to_past': f'Type:=xlTypePDF, {TO}, To:=99, {SAFE}', 'from_after_to': f'Type:=xlTypePDF, {TO}, From:=3, To:=1, {SAFE}', 'bogus': f'Type:=xlTypePDF, {TO}, Bogus:=1, {SAFE}',
    'class_pointer': f'Type:=xlTypePDF, {TO}, FixedFormatExtClassPtr:=Nothing, {SAFE}',
}
FILE = 'PrintToFile:=True'
NAMED = f'{FILE}, PrToFileName:=folder & "\\p.prn"'
PRINT_ROWS = {
    'base': NAMED, 'name_relative': f'{FILE}, PrToFileName:="rel.prn"', 'name_relative_folder': f'{FILE}, PrToFileName:="sub\\rel.prn"', 'name_empty': f'{FILE}, PrToFileName:=""',
    'name_null': f'{FILE}, PrToFileName:=Null', 'name_number': f'{FILE}, PrToFileName:=5', 'name_no_folder': f'{FILE}, PrToFileName:=folder & "\\nothere\\p.prn"',
    'name_is_folder': f'{FILE}, PrToFileName:=folder', 'name_read_only': f'{FILE}, PrToFileName:=folder & "\\ro.prn"', 'name_bad_character': f'{FILE}, PrToFileName:=folder & "\\a?b.prn"',
    'from_0': f'{NAMED}, From:=0', 'from_negative': f'{NAMED}, From:=-1', 'from_past': f'{NAMED}, From:=99', 'from_text': f'{NAMED}, From:="x"', 'from_null': f'{NAMED}, From:=Null',
    'to_0': f'{NAMED}, To:=0', 'from_after_to': f'{NAMED}, From:=3, To:=1', 'copies_0': f'{NAMED}, Copies:=0', 'copies_negative': f'{NAMED}, Copies:=-1', 'copies_2': f'{NAMED}, Copies:=2',
    'copies_text': f'{NAMED}, Copies:="x"', 'copies_null': f'{NAMED}, Copies:=Null', 'copies_fraction': f'{NAMED}, Copies:=1.5', 'collate_text': f'{NAMED}, Collate:="x"', 'collate_true': f'{NAMED}, Collate:=True',
    'printer_unknown': f'{NAMED}, ActivePrinter:="No Such Printer"', 'printer_empty': f'{NAMED}, ActivePrinter:=""', 'printer_null': f'{NAMED}, ActivePrinter:=Null', 'preview_false': f'{NAMED}, Preview:=False',
    'areas_text': f'{NAMED}, IgnorePrintAreas:="x"', 'areas_true': f'{NAMED}, IgnorePrintAreas:=True', 'bogus': f'{NAMED}, Bogus:=1',
}
# Each uses wb; the event's count is in the workbook's last cell, read after. Rows outside the owners' table.
STATES = {
    'export_empty_sheet': f'wb.Worksheets(3).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}',
    'export_hidden_sheet': f'wb.Worksheets(2).Visible = xlSheetHidden\nwb.Worksheets(2).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}',
    'export_not_active': f'wb.Worksheets(2).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}\ns = wb.ActiveSheet.Name',
    'export_protected': f'wb.Worksheets(1).Protect\nwb.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}',
    'export_returns': f'r = wb.Worksheets(1).ExportAsFixedFormat(Type:=xlTypePDF, {TO}, {SAFE})\ns = TypeName(r)',
    'export_saved_flag': f'wb.Saved = True\nwb.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}\ns = wb.Saved',
    'print_empty_sheet': f'wb.Worksheets(3).PrintOut {NAMED}',
    'print_hidden_sheet': f'wb.Worksheets(2).Visible = xlSheetHidden\nwb.Worksheets(2).PrintOut {NAMED}',
    'print_not_active': f'wb.Worksheets(2).PrintOut {NAMED}\ns = wb.ActiveSheet.Name',
    'print_returns': f'r = wb.Worksheets(1).PrintOut({NAMED})\ns = TypeName(r)',
    'print_saved_flag': f'wb.Saved = True\nwb.Worksheets(1).PrintOut {NAMED}\ns = wb.Saved',
}
# BeforePrint: its code goes in the workbook being printed, as the row runs with events on. It counts itself in a cell, and cancels in the rows that say so.
EVENTS = {
    'event_export': (False, f'wb.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}'),
    'event_export_cancelled': (True, f'wb.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}'),
    'event_export_bad_type': (False, f'wb.Worksheets(1).ExportAsFixedFormat Type:=2, {TO}, {SAFE}'),
    'event_export_no_folder': (False, f'wb.Worksheets(1).ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & "\\nothere\\x.pdf", {SAFE}'),
    'event_export_book': (False, f'wb.ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}'),
    'event_export_range': (False, f'wb.Worksheets(1).Range("A1").ExportAsFixedFormat Type:=xlTypePDF, {TO}, {SAFE}'),
    'event_print': (False, f'wb.Worksheets(1).PrintOut {NAMED}'),
    'event_print_cancelled': (True, f'wb.Worksheets(1).PrintOut {NAMED}'),
    'event_print_no_folder': (False, f'wb.Worksheets(1).PrintOut {FILE}, PrToFileName:=folder & "\\nothere\\p.prn"'),
    'event_print_bad_from': (False, f'wb.Worksheets(1).PrintOut {NAMED}, From:="x"'),
}
MAKER = r'''Public Function Make(ByVal path As String) As String
    Dim wb As Workbook, r As Long
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "One"
    wb.Worksheets.Add(After:=wb.Worksheets(1)).Name = "Two"
    wb.Worksheets.Add(After:=wb.Worksheets(2)).Name = "Blank"
    For r = 1 To 150
        wb.Worksheets(1).Cells(r, 1).Value = r
        wb.Worksheets(2).Cells(r, 1).Value = r
    Next
    wb.Charts.Add.SetSourceData wb.Worksheets(1).Range("A1:A5")
    wb.Worksheets(1).Activate
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Make = Application.ActivePrinter
End Function
'''


def row_source(call: str, declared: str, get: str) -> str:
    """One row in a module of its own, so a compile error stops only that row."""
    body = ''.join(f'    {line}\n' for line in call.split('\n'))
    return f'''Public Function Row(ByVal path As String, ByVal folder As String, ByVal cancel As Long) As String
    Dim wb As Workbook, o As {declared}, s As String, r As Variant, was As String, there As String, n As Long, e As Long, d As String
    Application.DisplayAlerts = False
    was = CurDir
    there = Application.DefaultFilePath
    ChDrive Left(folder, 1)
    ChDir folder
    Application.DefaultFilePath = folder
    Set wb = Workbooks.Open(path, 0)
{"    Set o = " + get + chr(10) if get else ""}    If cancel >= 0 Then
        wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString "Private Sub Workbook_BeforePrint(Cancel As Boolean)" & vbLf & "Application.EnableEvents = False" & vbLf & "Worksheets(3).Range(""Z1"").Value = Worksheets(3).Range(""Z1"").Value + 1" & vbLf & "Application.EnableEvents = True" & vbLf & "Cancel = " & IIf(cancel = 1, "True", "False") & vbLf & "End Sub"
        Application.EnableEvents = True
    End If
    On Error Resume Next
{body}    e = Err.Number
    d = Err.Description
    On Error GoTo 0
    Application.EnableEvents = False
    If cancel >= 0 Then s = s & "events=" & wb.Worksheets(3).Range("Z1").Value
    wb.Close False
    Application.DefaultFilePath = there
    ChDrive Left(was, 1)
    ChDir was
    Row = "E" & e & "~" & Left(d, 120) & "~" & s & "~" & Application.ActivePrinter
End Function
'''


def cases() -> list[dict[str, object]]:
    found: list[dict[str, object]] = []
    for owner in EXPORTS:
        for binding in ('Typed', 'Late'):
            found.extend({'member': 'ExportAsFixedFormat', 'owner': owner, 'binding': binding, 'row': name, 'call': f'o.ExportAsFixedFormat {call}', 'cancel': -1} for name, call in EXPORT_ROWS.items())
    for owner in OWNERS:
        for binding in ('Typed', 'Late'):
            found.extend({'member': 'PrintOut', 'owner': owner, 'binding': binding, 'row': name, 'call': f'o.PrintOut {call}', 'cancel': -1} for name, call in PRINT_ROWS.items())
    found.extend({'member': 'state', 'owner': '', 'binding': 'Typed', 'row': name, 'call': call, 'cancel': -1} for name, call in STATES.items())
    found.extend({'member': 'event', 'owner': '', 'binding': 'Typed', 'row': name, 'call': call, 'cancel': int(cancel)} for name, (cancel, call) in EVENTS.items())
    return found


def documents() -> set[str]:
    home = Path(os.environ.get('USERPROFILE', '.'))
    found: set[str] = set()
    for folder in (home / 'Documents', home / 'OneDrive' / 'Documents'):
        if folder.is_dir():
            found |= {str(one) for one in folder.iterdir() if one.suffix.lower() in ('.pdf', '.xps', '.prn')}
    return found


def main() -> None:
    all_cases = cases()
    done: dict[str, dict[str, object]] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    attempts = 0
    before = documents()
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        book = home / 'book.xlsx'
        made = session.run_vba(MAKER, 'Make', (str(book),), timeout=60.0, line_numbers=False)
        assert made.ok, made
        printer = str(made.value)
        for case in all_cases:
            key = f'{case["member"]} {case["owner"]} {case["binding"]} {case["row"]}'
            if key in done:
                continue
            if time.monotonic() - started > BUDGET:
                print('stopping for now', len(done), 'of', len(all_cases), flush=True)
                break
            # A new folder for each try: one a stopped Excel held may not be clean.
            attempts += 1
            folder = home / f'run{attempts}'
            (folder / 'sub').mkdir(parents=True)
            (folder / 'there.pdf').write_bytes(b'before')
            for name in ('ro.pdf', 'ro.prn'):
                (folder / name).write_bytes(b'before')
                (folder / name).chmod(stat.S_IREAD)
            owner = str(case['owner'])
            declared = owner if case['binding'] == 'Typed' and owner else 'Object'
            source = row_source(str(case['call']), declared, OWNERS.get(owner, ''))
            try:
                ran = session.run_vba(source, 'Row', (str(book), str(folder), case['cancel']), timeout=45.0, line_numbers=False)
            except HarnessError as failed:
                print('restarting', str(failed)[:80], flush=True)
                session.close()
                time.sleep(5)
                session = ExcelSession(HarnessConfig(exclusive=False))
                session.new_document()
                continue
            answer = str(ran.value) if ran.ok else f'failed {ran.outcome} {[one.message[:110] for one in ran.dialogs]}'
            answer = answer.replace(str(folder), '<DIR>').replace(held, '<HOME>')
            files = sorted(str(one.relative_to(folder)).replace('\\', '/') + (f' before' if one.read_bytes() == b'before' else '') for one in folder.rglob('*') if one.is_file() and one.name not in ('ro.pdf', 'ro.prn', 'there.pdf') or one.is_file() and one.name in ('there.pdf',) and one.read_bytes() != b'before')
            done[key] = {**case, 'excel': answer, 'files': files}
            if ran.ok and not answer.endswith('~' + printer):
                print('PRINTER CHANGED, stopping', key, answer[-60:], flush=True)
                PARTIAL.write_text(json.dumps(done), encoding='utf-8')
                session.close()
                return
            print(key, answer[:90], files[:3], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
            for name in ('ro.pdf', 'ro.prn'):
                (folder / name).chmod(stat.S_IWRITE)
    session.close()
    stray = documents() - before
    if stray:
        print('NEW FILES IN DOCUMENTS', sorted(stray), flush=True)
    if len(done) == len(all_cases):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': list(done.values())}, indent=0) + '\n', encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(done), 'cases', flush=True)


if __name__ == '__main__':
    main()
