"""ExportAsFixedFormat and PrintOut on a chart in a frame, typed and late: the rows of scripts/measure_output_members.py, whose chart was a chart sheet. Rival rules: it answers as a chart sheet does, or differs as what it prints is part of a sheet.

A valid export of one answers -2147417848, "The object invoked has disconnected from its clients", bound late, shown or hidden and active or not, and writes no file; bound typed it stops the
macro's run, which takes the harness a restart each time. So a typed row is run only where the late one did not disconnect, and otherwise is recorded as `typed not run`.

Safe as that script is: OpenAfterPublish:=False on every export, PrintToFile:=True on every print, no Preview, everything under a temporary folder, and a changed printer stops the run.

    python scripts/measure_output_embedded_chart.py

writes tests/fixtures/output_embedded_chart.json, which tests/test_excel_output_hooks.py replays.
"""
from __future__ import annotations

import json
import stat
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from measure_output_members import EXPORT_ROWS, PRINT_ROWS, row_source  # noqa: E402

OUT = ROOT / 'tests/fixtures/output_embedded_chart.json'
BOOK = ROOT / 'tests/fixtures/output_members.xlsx'
GET = 'wb.Worksheets(1).ChartObjects(1).Chart'
MAKE = r'''Public Function Make(ByVal book As String, ByVal path As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(book, 0)
    wb.Worksheets(1).ChartObjects.Add(10, 10, 200, 100).Chart.SetSourceData wb.Worksheets(1).Range("A1:A5")
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
    Make = Application.ActivePrinter
End Function
'''
RETURNS = r'''Public Function Back(ByVal path As String, ByVal folder As String, ByVal cancel As Boolean) As String
    Dim wb As Workbook, o As Object, r As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString "Private Sub Workbook_BeforePrint(Cancel As Boolean)" & vbLf & "Cancel = " & cancel & vbLf & "End Sub"
    Application.EnableEvents = True
    Set o = wb.Worksheets(1).ChartObjects(1).Chart
    r = o.PrintOut(PrintToFile:=True, PrToFileName:=folder & "\p.prn")
    Application.EnableEvents = False
    Back = TypeName(r) & " " & CStr(r)
    wb.Close False
End Function
'''


def main() -> None:
    rows = [('ExportAsFixedFormat', name, f'o.ExportAsFixedFormat {call}') for name, call in EXPORT_ROWS.items()] + [('PrintOut', name, f'o.PrintOut {call}') for name, call in PRINT_ROWS.items()]
    cases: list[dict[str, object]] = []
    returns: dict[str, str] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        home = Path(held)
        book = home / 'book.xlsx'
        session = ExcelSession(HarnessConfig(exclusive=False))
        session.new_document()
        made = session.run_vba(MAKE, 'Make', (str(BOOK), str(book)), timeout=60.0, line_numbers=False)
        assert made.ok, made
        printer = str(made.value)
        attempt = 0
        for member, name, call in rows:
            disconnected = False
            for binding, declared in (('Late', 'Object'), ('Typed', 'Chart')):
                attempt += 1
                if binding == 'Typed' and disconnected:
                    cases.append({'member': member, 'owner': 'Chart', 'binding': binding, 'row': name, 'call': call, 'excel': 'failed typed not run', 'files': []})
                    continue
                folder = home / f'run{attempt}'
                (folder / 'sub').mkdir(parents=True)
                (folder / 'there.pdf').write_bytes(b'before')
                for readonly in ('ro.pdf', 'ro.prn'):
                    (folder / readonly).write_bytes(b'before')
                    (folder / readonly).chmod(stat.S_IREAD)
                try:
                    ran = session.run_vba(row_source(call, declared, GET), 'Row', (str(book), str(folder), -1), timeout=45.0, line_numbers=False)
                except HarnessError:
                    session.close()
                    time.sleep(5)
                    session = ExcelSession(HarnessConfig(exclusive=False))
                    session.new_document()
                    ran = session.run_vba(row_source(call, declared, GET), 'Row', (str(book), str(folder), -1), timeout=45.0, line_numbers=False)
                answer = str(ran.value) if ran.ok else f'failed {ran.outcome} {[one.message[:110] for one in ran.dialogs]}'
                answer = answer.replace(str(folder), '<DIR>').replace(held, '<HOME>')
                files = sorted(str(one.relative_to(folder)).replace('\\', '/') for one in folder.rglob('*') if one.is_file() and one.name not in ('ro.pdf', 'ro.prn', 'there.pdf') and one.suffix != '.tmp')
                cases.append({'member': member, 'owner': 'Chart', 'binding': binding, 'row': name, 'call': call, 'excel': answer, 'files': files})
                disconnected = answer.startswith('E-2147417848')
                print(member, binding, name, answer[:70], files[:2], flush=True)
                for readonly in ('ro.pdf', 'ro.prn'):
                    (folder / readonly).chmod(stat.S_IWRITE)
                assert not ran.ok or answer.endswith('~' + printer), f'the printer changed at {name}'
        for cancel in (False, True):
            done = session.run_vba(RETURNS, 'Back', (str(book), held, cancel), timeout=45.0, line_numbers=False)
            returns[str(cancel)] = str(done.value) if done.ok else f'failed {done.outcome}'
            print('returns', cancel, returns[str(cancel)], flush=True)
        session.close()
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'get': GET, 'cases': cases, 'returns': returns}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'cases')


if __name__ == '__main__':
    main()
