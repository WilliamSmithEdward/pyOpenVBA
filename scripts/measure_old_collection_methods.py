"""A method called on a whole old collection, Worksheet.DropDowns.AddItem "z", Worksheet.Buttons.BringToFront and the like, typed and
late, on the workbook Excel saved for tests/fixtures/old_collection_members.json: what the call answers, the sheet's drawings and lists
after it, and the same once the workbook is saved and opened again.

Rival rules, written before measuring:

- A list method (AddItem, RemoveItem, RemoveAllItems) on DropDowns or ListBoxes changes every item's list, as a property set on the
  collection does (tests/fixtures/old_collection_lets.json); or only the first item's; or is refused (1004). On DrawingObjects it reaches
  only the drawings that have a list, or fails the whole call once an item has none.
- BringToFront and SendToBack move all the items together and keep their order among themselves; or move them one at a time, which
  reverses them; or are refused.
- Duplicate copies every item and answers the copies as a collection of the same kind; or only the first, answering that item.
- A method the collection's items do not have (Buttons.AddItem) is 438.

Typed, a member the collection's type does not have is a compile error, which stops the harness, so typed rows are only the methods the
checklist lists on that collection. Each row runs in a module of its own. Nothing here reads a size or a place, so the display's DPI does
not enter.

Safety (the Supervisor's rules, 2026-11-07): one Excel through live.lock (an exclusive session); no clipboard, spelling, mail, print,
export, encryption, signing, certificate, IRM, sensitivity-label or identity member; the workbook is a copy under TEMP, opened for
writing and saved under a new name; no window is touched.

    python scripts/measure_old_collection_methods.py [seconds]

writes tests/fixtures/old_collection_methods.json, which tests/test_excel_old_collection_methods.py replays. It keeps its place in a
file under TEMP, so a run that stops can go on.
"""
from __future__ import annotations

import csv
import json
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/old_collection_methods.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_methods_partial.json'
BUDGET = 520.0
COLLECTIONS = ('Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures', 'Rectangles', 'TextBoxes', 'DrawingObjects')
#: Each call: the checklist member it is, and the statement that makes it on ``o``. ``v = ...`` keeps what the call answers.
CALLS: dict[str, tuple[str, str]] = {
    'AddItem': ('AddItem', 'o.AddItem "z"'),
    'AddItem at 1': ('AddItem', 'o.AddItem "z", 1'),
    'RemoveItem': ('RemoveItem', 'o.RemoveItem 1'),
    'RemoveItem two': ('RemoveItem', 'o.RemoveItem 1, 2'),
    'RemoveAllItems': ('RemoveAllItems', 'o.RemoveAllItems'),
    'BringToFront': ('BringToFront', 'o.BringToFront'),
    'SendToBack': ('SendToBack', 'o.SendToBack'),
    'Duplicate': ('Duplicate', 'Set v = o.Duplicate'),
    'List': ('List', 'v = o.List(1)'),
    'Selected': ('Selected', 'v = o.Selected(1)'),
    'LinkCombo': ('LinkCombo', 'o.LinkCombo True'),
}
DRIVER = r'''Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "obj " & TypeName(v)
    ElseIf IsArray(v) Then
        Show = "array"
    ElseIf IsNull(v) Then
        Show = "null"
    ElseIf IsError(v) Then
        Show = "error " & CStr(v)
    Else
        Show = TypeName(v) & " " & CStr(v)
    End If
End Function

Private Function State(ws As Object) As String
    Dim sh As Object, s As String, i As Long, t As String
    For Each sh In ws.Shapes
        s = s & "," & sh.Name
    Next
    s = Mid(s, 2)
    On Error Resume Next
    For i = 1 To ws.DropDowns.Count
        Err.Clear
        t = Join(ws.DropDowns(i).List, ",")
        If Err.Number Then t = "E" & Err.Number
        s = s & "|dd" & i & " " & ws.DropDowns(i).ListCount & ":" & t
    Next
    For i = 1 To ws.ListBoxes.Count
        Err.Clear
        t = Join(ws.ListBoxes(i).List, ",")
        If Err.Number Then t = "E" & Err.Number
        s = s & "|lb" & i & " " & ws.ListBoxes(i).ListCount & ":" & t
    Next
    State = s
End Function

Public Function Row(ByVal path As String, ByVal saved As String) As String
    Dim wb As Workbook, o As {declared}, e As Long, s As String, v As Variant
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set o = wb.Worksheets(1).{c}
    v = Empty
    On Error Resume Next
    {call}
    e = Err.Number
    Err.Clear
    s = "E" & e & "|" & Show(v)
    If Err.Number Then s = "E" & e & "|unshown"
    On Error GoTo 0
    s = s & "|" & State(wb.Worksheets(1))
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = s & "|reopened " & State(wb.Worksheets(1))
    wb.Close False
    Row = s
End Function
'''


def listed() -> set[tuple[str, str]]:
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as file:
        return {(row['object'], row['member']) for row in csv.DictReader(file) if row['kind'] == 'method'}


def cases() -> list[dict[str, str]]:
    known = listed()
    found: list[dict[str, str]] = []
    for collection in COLLECTIONS:
        for call, (member, _) in CALLS.items():
            for binding in ('Typed', 'Late'):
                if binding == 'Typed' and (collection, member) not in known:
                    continue
                found.append({'collection': collection, 'call': call, 'member': member, 'binding': binding})
    return found


def row_source(case: dict[str, str]) -> str:
    declared = f'Excel.{case["collection"]}' if case['binding'] == 'Typed' else 'Object'
    return DRIVER.replace('{declared}', declared).replace('{c}', case['collection']).replace('{call}', CALLS[case['call']][1])


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    budget = float(sys.argv[1]) if len(sys.argv) > 1 else BUDGET
    all_cases = cases()
    done: dict[str, str] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as session:
        session.new_document()
        for number, case in enumerate(all_cases):
            key = f'{case["collection"]} {case["call"]} {case["binding"]}'
            if key in done:
                continue
            if time.monotonic() - started > budget:
                print('stopping for now', len(done), 'of', len(all_cases), flush=True)
                break
            book = Path(held) / f'book{number}.xlsx'
            book.write_bytes(BOOK.read_bytes())
            try:
                ran = session.run_vba(row_source(case), 'Row', (str(book), str(Path(held) / f'saved{number}.xlsx')), timeout=45.0, line_numbers=False)
            except HarnessError as failed:
                # A typed compile error or a hung call stops this session; the next run picks up after it.
                done[key] = f'failed harness {str(failed)[:60]}'
                PARTIAL.write_text(json.dumps(done), encoding='utf-8')
                print(key, done[key], flush=True)
                break
            done[key] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome} {[one.message[:90] for one in ran.dialogs]}'
            print(key, done[key][:200], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
    if len(done) == len(all_cases):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'driver': DRIVER, 'calls': CALLS,
                                   'cases': [{**case, 'excel': done[f'{case["collection"]} {case["call"]} {case["binding"]}']} for case in all_cases]}, indent=0) + '\n',
                       encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(all_cases), 'cases', flush=True)


if __name__ == '__main__':
    main()
