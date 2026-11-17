"""The hidden _Default of the old control collections whose type library gives it As Long: CheckBoxes, OptionButtons, DropDowns,
ListBoxes and DrawingObjects, typed and late, on the workbook Excel saved for tests/fixtures/old_collection_members.json: read as
[_Default], read as the collection's own value, and set, with the items as they are, alike and different.

Rival rules, written before measuring:

- Typed, [_Default] is the items' Value, read as a read of Value over them is (their value when alike, 1004 when not), and a set
  reaches every item's Value (rival: it is the item at an index, as on other collections; rival: 438).
- Late, [_Default] is 438, as every hidden member on an As Object variable measured so far (rival: as typed).
- The collection read as a value (v = c) goes through the same _Default typed, and is 438 or a type mismatch late.

Nothing is sized. Safety (the Supervisor's rules): one Excel through live.lock, no mail, print, export, encryption, signing, certificate,
IRM, sensitivity-label or identity member, and no window touched.

    python scripts/measure_old_collection_defaults.py

writes tests/fixtures/old_collection_defaults.json, which tests/test_excel_old_collection_defaults.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/old_collection_defaults.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
COLLECTIONS = ('CheckBoxes', 'OptionButtons', 'DropDowns', 'ListBoxes', 'DrawingObjects')
#: What each collection's items are set to, alike, and what the second alone is set to, different.
VALUES = {'CheckBoxes': ('1', '-4146'), 'OptionButtons': ('1', '-4146'), 'DropDowns': ('2', '3'), 'ListBoxes': ('2', '3'), 'DrawingObjects': ('1', '2')}
OPERATIONS = ('read', 'value', 'set')
STATES = ('as is', 'alike', 'differ')
SOURCE = r'''Private Function Items(ws As Object) As String
    Dim x As Object, s As String
    On Error Resume Next
    For Each x In ws.{c}
        Err.Clear
        s = s & ";" & TypeName(x) & " " & x.Value
        If Err.Number Then s = s & " E" & Err.Number
    Next
    Items = Mid(s, 2)
End Function

Public Function Row(ByVal path As String, ByVal state As String, ByVal operation As String) As String
    Dim wb As Workbook, ws As Object, c As {declared}, v As Variant, s As String, x As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    For Each x In ws.DropDowns
        x.RemoveAllItems: x.AddItem "a": x.AddItem "b": x.AddItem "c"
    Next
    For Each x In ws.ListBoxes
        x.RemoveAllItems: x.AddItem "a": x.AddItem "b": x.AddItem "c"
    Next
    Err.Clear
    If state = "alike" Then
        For Each x In ws.{c}
            x.Value = {alike}
        Next
    ElseIf state = "differ" Then
        ws.{c}(2).Value = {differ}
    End If
    s = "set" & Err.Number & " "
    Err.Clear
    Set c = ws.{c}
    If operation = "read" Then
        v = c.[_Default]
    ElseIf operation = "value" Then
        v = c
    Else
        c.[_Default] = {alike}
    End If
    If Err.Number Then
        s = s & "E" & Err.Number
    ElseIf operation = "set" Then
        s = s & "ok"
    Else
        s = s & TypeName(v) & " " & CStr(v)
    End If
    Err.Clear
    On Error GoTo 0
    s = s & "|" & Items(ws)
    wb.Close False
    Row = s
End Function
'''


def cases() -> list[dict[str, str]]:
    return [{'collection': collection, 'binding': binding, 'state': state, 'operation': operation}
            for collection in COLLECTIONS for binding in ('Typed', 'Late') for state in STATES for operation in OPERATIONS]


def source_of(case: dict[str, str]) -> str:
    alike, differ = VALUES[case['collection']]
    declared = f'Excel.{case["collection"]}' if case['binding'] == 'Typed' else 'Object'
    return (SOURCE.replace('{c}', case['collection']).replace('{declared}', declared).replace('{alike}', alike).replace('{differ}', differ))


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        book = Path(held) / 'in.xlsx'
        for case in cases():
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(source_of(case), 'Row', (str(book), case['state'], case['operation']), timeout=60.0, line_numbers=False)
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            found.append({**case, 'excel': answer})
            print(case, answer, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'values': VALUES, 'cases': found}, indent=1) + '\n',
                   encoding='utf-8')
    print('wrote', len(found), 'cases', flush=True)


if __name__ == '__main__':
    main()
