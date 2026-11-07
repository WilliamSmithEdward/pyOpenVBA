"""A property set on a whole old collection, Worksheet.Buttons.Caption = "x" and the like, typed and late, on the workbook Excel saved for
tests/fixtures/old_collection_members.json: what the set answers, what each item reads after, and what each reads once the workbook is
saved and opened again.

Rival rules, written before measuring: a set on the collection sets every item (as a ShapeRange does), or only the first, or is
refused as the read is (1004, which Excel answers for most reads over items that differ); items without the property (DrawingObjects
mixes kinds) are skipped, or fail the whole set. Left, Top, Width and Height are left out, as they depend on the display's DPI.

Typed, a member the collection's type does not have is a compile error, which stops the harness, so typed rows are only those the
checklist lists on that collection. Each row runs in a module of its own.

    python scripts/measure_old_collection_lets.py [seconds]
    python scripts/measure_old_collection_lets.py subsets

writes tests/fixtures/old_collection_lets.json, which tests/test_excel_old_collection_lets.py replays. It keeps its place in a file
under TEMP, so a run that stops can go on.
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
OUT = ROOT / 'tests/fixtures/old_collection_lets.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_lets_partial.json'
BUDGET = 520.0
COLLECTIONS = ('Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures', 'Rectangles', 'TextBoxes', 'DrawingObjects')
#: The value each property is set to: one that differs from what a new item holds.
VALUES = {
    'Accelerator': '"K"', 'AddIndent': 'True', 'ArrowHeadLength': 'xlArrowHeadLengthLong', 'ArrowHeadStyle': 'xlArrowHeadStyleOpen',
    'ArrowHeadWidth': 'xlArrowHeadWidthWide', 'AutoScaleFont': 'True', 'AutoSize': 'True', 'CancelButton': 'True', 'Caption': '"Fresh"',
    'DefaultButton': 'True', 'DismissButton': 'True', 'Display3DShading': 'True', 'DisplayVerticalScrollBar': 'True', 'DropDownLines': '4',
    'Enabled': 'False', 'Formula': '"=$A$21"', 'HelpButton': 'True', 'HorizontalAlignment': 'xlCenter', 'InputType': 'xlNumber', 'LargeChange': '7',
    'LinkedCell': '"$D$20"', 'List': 'Array("x", "y")', 'ListFillRange': '"$A$21:$A$22"', 'ListIndex': '1', 'Locked': 'False', 'LockedText': 'False',
    'Max': '50', 'Min': '2', 'MultiLine': 'True', 'MultiSelect': 'xlSimple', 'Orientation': 'xlUpward', 'PhoneticAccelerator': '"Q"',
    'Placement': 'xlFreeFloating', 'PrintObject': 'False', 'ReadingOrder': 'xlRTL', 'RoundedCorners': 'True', 'Shadow': 'True', 'SmallChange': '3',
    'Text': '"Fresh"', 'Value': 'xlOn', 'VerticalAlignment': 'xlCenter', 'Visible': 'False',
}
#: How a read after shows a List, which is an array.
READ_AS = {'List': 'Join(ws.{c}(i).List, ",")'}
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

Private Function Items(ws As Object) As String
    Dim i As Long, s As String
    On Error Resume Next
    For i = 1 To ws.{c}.Count
        Err.Clear
        s = s & ";" & Show({read})
        If Err.Number Then s = s & "E" & Err.Number
    Next
    Items = ws.{c}.Count & ":" & Mid(s, 2)
End Function

Public Function Row(ByVal path As String, ByVal saved As String) As String
    Dim wb As Workbook, o As {declared}, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set o = wb.Worksheets(1).{c}
    On Error Resume Next
    o.{m} = {v}
    e = Err.Number
    Err.Clear
    On Error GoTo 0
    s = "E" & e & "|" & Items(wb.Worksheets(1))
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = s & "|reopened " & Items(wb.Worksheets(1))
    wb.Close False
    Row = s
End Function
'''


def listed() -> set[tuple[str, str]]:
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as file:
        return {(row['object'], row['member']) for row in csv.DictReader(file) if row['kind'] == 'property'}


def cases() -> list[dict[str, str]]:
    known = listed()
    found: list[dict[str, str]] = []
    for collection in COLLECTIONS:
        for member, value in VALUES.items():
            if (collection, member) not in known:
                continue
            for binding in ('Typed', 'Late'):
                found.append({'collection': collection, 'member': member, 'value': value, 'binding': binding})
    return found


def row_source(case: dict[str, str]) -> str:
    collection, member = case['collection'], case['member']
    read = READ_AS.get(member, 'ws.{c}(i).{m}').format(c=collection, m=member)
    declared = f'Excel.{collection}' if case['binding'] == 'Typed' else 'Object'
    return (DRIVER.replace('{read}', read).replace('{declared}', declared).replace('{c}', collection).replace('{m}', member).replace('{v}', case['value']))


#: Each kind's two items, as DrawingObjects picks them by name: a set on all the sheet's drawings split by kind, to find which kind makes the
#: whole set answer an error that the kind's own collection does not.
KINDS = {'Arcs': 'Arc', 'Buttons': 'Button', 'CheckBoxes': 'CheckBox', 'DropDowns': 'DropDown', 'Lines': 'Line', 'ListBoxes': 'ListBox',
         'OptionButtons': 'OptionButton', 'Ovals': 'Oval', 'Pictures': 'Picture', 'Rectangles': 'Rectangle', 'TextBoxes': 'TextBox'}
SUBSET = r'''Public Function Row(ByVal path As String, ByVal saved As String) As String
    Dim wb As Workbook, ws As Object, o As Object, e As Long, s As String, x As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    Set o = ws.DrawingObjects(Array(ws.{c}(1).Name, ws.{c}(2).Name))
    On Error Resume Next
    o.{m} = {v}
    e = Err.Number
    Err.Clear
    For Each x In o
        Err.Clear
        s = s & ";" & TypeName(x.{m}) & " " & CStr(x.{m})
        If Err.Number Then s = s & "E" & Err.Number
    Next
    On Error GoTo 0
    wb.Close False
    Row = "E" & e & "|" & Mid(s, 2)
End Function
'''


def subsets() -> None:
    """Each property set on the sheet's drawings picked down to one kind, late, its answer and the items' reads after: tests/fixtures/old_collection_subsets.json."""
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    out = ROOT / 'tests/fixtures/old_collection_subsets.json'
    found: list[dict[str, str]] = []
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for member, value in VALUES.items():
            for collection in KINDS:
                book = Path(held) / 'book.xlsx'
                book.write_bytes(BOOK.read_bytes())
                source = SUBSET.replace('{c}', collection).replace('{m}', member).replace('{v}', value)
                try:
                    ran = session.run_vba(source, 'Row', (str(book), ''), timeout=45.0, line_numbers=False)
                    answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                except HarnessError as failed:
                    session.close()
                    time.sleep(5)
                    session = ExcelSession(HarnessConfig(exclusive=False))
                    session.new_document()
                    answer = f'failed harness {str(failed)[:60]}'
                found.append({'member': member, 'value': value, 'kind': collection, 'excel': answer})
                print(member, collection, answer[:120], flush=True)
    session.close()
    out.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SUBSET, 'cases': found}, indent=0) + '\n', encoding='utf-8')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    budget = float(sys.argv[1]) if len(sys.argv) > 1 else BUDGET
    all_cases = cases()
    done: dict[str, str] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        for number, case in enumerate(all_cases):
            key = f'{case["collection"]} {case["member"]} {case["binding"]}'
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
                print('restarting', str(failed)[:80], flush=True)
                session.close()
                time.sleep(5)
                session = ExcelSession(HarnessConfig(exclusive=False))
                session.new_document()
                done[key] = f'failed harness {str(failed)[:60]}'
                continue
            done[key] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome} {[one.message[:90] for one in ran.dialogs]}'
            print(key, done[key][:150], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
    session.close()
    if len(done) == len(all_cases):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'driver': DRIVER, 'values': VALUES,
                                   'cases': [{**case, 'excel': done[f'{case["collection"]} {case["member"]} {case["binding"]}']} for case in all_cases]}, indent=0) + '\n',
                       encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(all_cases), 'cases', flush=True)


if __name__ == '__main__':
    if sys.argv[1:] == ["subsets"]:
        subsets()
    else:
        main()
