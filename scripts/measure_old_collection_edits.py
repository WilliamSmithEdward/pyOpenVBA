"""Changes to the sheet's old collections and their items on a workbook Excel saved: what each answers, and what Excel reads from the save.

Each case opens tests/fixtures/old_collections.xlsx, makes one change through the collection or its first item, reads every item, saves and
reads the file again: in Excel, then the same in the model with Excel reading the model's save. Rival rules for a set through the collection:
it reaches every item, or only the first. Left, Top, Width and Height are left out, as they depend on the display's DPI.

    python scripts/measure_old_collection_edits.py

with --model, keeps what Excel answered and measures only what Excel reads from the model's save. It writes tests/fixtures/old_collection_edits.json, which tests/test_excel_old_collection_edits.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAError, VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/old_collection_edits.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_edits_partial.json'
COLLECTIONS = ['Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures', 'Rectangles', 'TextBoxes']
CHANGES = ['Visible = False', 'Locked = False', 'PrintObject = False', 'Enabled = False', 'Placement = 3', 'Caption = "New"', 'Text = "New"', 'LockedText = False',
           'Value = 1', 'OnAction = "M"', 'Name = "Renamed"', 'Delete', 'Font.Bold = True', 'Interior.ColorIndex = 6', 'Border.ColorIndex = 3', 'Characters.Text = "Ch"']
READS = ['Name', 'Visible', 'Locked', 'PrintObject', 'Placement', 'Enabled', 'Text', 'Caption', 'Value', 'LockedText', 'OnAction', 'Font.Bold', 'Interior.ColorIndex', 'Border.ColorIndex']


def source() -> str:
    pick = ''.join(f'    Case "{name}"\n        Set Held = ws.{name}\n' for name in COLLECTIONS)
    reads = ''.join(f'    Case {index}\n        v = o.{read}\n' for index, read in enumerate(READS))
    changes = ''.join(f'    Case {index}\n        If first Then\n            c(1).{change}\n        Else\n            c.{change}\n        End If\n' for index, change in enumerate(CHANGES))
    return f'''Private Function Held(ws As Object, ByVal name As String) As Object
    Select Case name
{pick}    End Select
End Function

Public Function ReadOne(ws As Object, ByVal name As String, ByVal place As Long, ByVal which As Long) As String
    Dim o As Object, v As Variant
    Set o = Held(ws, name)(place)
    On Error Resume Next
    Select Case which
{reads}    End Select
    If Err.Number Then
        ReadOne = "E" & Err.Number
    Else
        ReadOne = TypeName(v) & " " & CStr(v)
    End If
End Function

Public Function Counted(ws As Object, ByVal name As String) As Long
    Counted = Held(ws, name).Count
End Function

Public Function ReadAll(ws As Object, ByVal name As String) As String
    Dim place As Long, which As Long, s As String
    For place = 1 To Counted(ws, name)
        For which = 0 To {len(READS) - 1}
            s = s & ReadOne(ws, name, place, which) & ";"
        Next
    Next
    ReadAll = Counted(ws, name) & ":" & s
End Function

Public Function Change(ws As Object, ByVal name As String, ByVal first As Boolean, ByVal which As Long) As String
    Dim c As Object
    Set c = Held(ws, name)
    On Error Resume Next
    Select Case which
{changes}    End Select
    Change = "E" & Err.Number
End Function

Public Function ReadSaved(ByVal path As String, ByVal name As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    ReadSaved = ReadAll(wb.Worksheets(1), name)
    wb.Close False
End Function
'''


DRIVER = '''
Public Function Tried(ByVal book As String, ByVal name As String, ByVal first As Boolean, ByVal which As Long, ByVal out As String) As String
    Dim wb As Workbook, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(book, 0)
    s = Change(wb.Worksheets(1), name, first, which)
    s = s & "@" & ReadAll(wb.Worksheets(1), name)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Tried = s & "@" & ReadSaved(out, name)
End Function

Public Function ReadSavedMany(ByVal list As String, ByVal name As String) As String
    Dim one As Variant, s As String
    For Each one In Split(list, "|")
        s = s & ReadSaved(one, name) & "@@"
    Next
    ReadSavedMany = s
End Function
'''
MODEL = '''
Public Function ChangeHere(ByVal name As String, ByVal first As Boolean, ByVal which As Long) As String
    ChangeHere = Change(ActiveWorkbook.Worksheets(1), name, first, which)
End Function

Public Function ReadHere(ByVal name As String) As String
    Dim place As Long, which As Long, s As String
    For place = 1 To Counted(ActiveWorkbook.Worksheets(1), name)
        For which = 0 To ''' + str(len(READS) - 1) + '''
            s = s & ReadOne(ActiveWorkbook.Worksheets(1), name, place, which) & ";"
        Next
    Next
    ReadHere = Counted(ActiveWorkbook.Worksheets(1), name) & ":" & s
End Function
'''


def in_model(module: str, name: str, first: bool, which: int, out: Path) -> str:
    """The change in the model and its save, or why it reports itself."""
    app = ExcelApplication.open(BOOK, with_vba=False)
    app.interpreter.sandbox = False
    app.add_module(module, name='Probe')
    try:
        answer = str(app.run('Probe.ChangeHere', name, first, which))
        app.save(out)
    except VBAUnsupportedError as refused:
        return f'reports {str(refused)[:70]}'
    except VBAError as failed:
        return f'failed {str(failed)[:70]}'
    return answer


def live(session: list[ExcelSession], module: str, routine: str, *arguments: object) -> str | None:
    for _ in range(3):
        try:
            done = session[0].run_vba(module, routine, arguments, timeout=30.0, line_numbers=False)
        except HarnessError as failed:
            print('restarting', str(failed)[:80], flush=True)
            session[0].close()
            time.sleep(5)
            session[0] = ExcelSession(HarnessConfig(exclusive=False))
            session[0].new_document()
            continue
        if not done.ok:
            print('  not passed', routine, arguments[1:4], done.outcome, [one.message[:90] for one in done.dialogs], flush=True)
        return str(done.value) if done.ok else None
    return None


def main() -> None:
    module = source()
    done: dict[str, dict[str, str]] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    keep = {f'{one["collection"]} {"item" if one["first"] else "collection"} {one["change"]}': one['excel']
            for one in json.loads(OUT.read_text(encoding='utf-8'))['cases']} if '--model' in sys.argv else {}
    session = [ExcelSession(HarnessConfig(exclusive=False))]
    session[0].new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder:
        for name in COLLECTIONS:
            for first in (False, True):
                for which, change in enumerate(CHANGES):
                    key = f'{name} {"item" if first else "collection"} {change}'
                    if key in done:
                        continue
                    own, mine = Path(folder) / f'own{len(done)}.xlsx', Path(folder) / f'mine{len(done)}.xlsx'
                    excel = keep[key] if keep else live(session, module + DRIVER, 'Tried', str(BOOK), name, first, which, str(own))
                    model = in_model(module + MODEL, name, first, which, mine)
                    read = live(session, module + DRIVER, 'ReadSaved', str(mine), name) if mine.exists() else None
                    done[key] = {'collection': name, 'first': first, 'change': change, 'excel': excel if excel is not None else 'failed', 'model_save': model,  # type: ignore[dict-item]
                                 'excel_reads_model': read if read is not None else ''}
                    parts = (excel or '').split('@')
                    print('same' if len(parts) == 3 and parts[2] == read else 'DIFF' if read else 'none', key, flush=True)
                    PARTIAL.write_text(json.dumps(done), encoding='utf-8')
    session[0].close()
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': module, 'reads': READS, 'changes': CHANGES, 'cases': list(done.values())}, indent=0) + '\n', encoding='utf-8')
    PARTIAL.unlink()
    print(len(done), 'cases')


if __name__ == '__main__':
    main()
