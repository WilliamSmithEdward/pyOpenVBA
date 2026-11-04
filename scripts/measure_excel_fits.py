"""How a value fits where it goes: an array written to a range of another size, a copy to a destination of another shape, a Date
beside a String, and what Range.Sort leaves on the sheet's Sort. Live Excel.

    python scripts/measure_excel_fits.py

writes tests/fixtures/excel_fits.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_fits.json'
HELPERS = '''Private Function Seen(ws As Object, ByVal address As String) As String
    Dim c As Object, s As String
    For Each c In ws.Range(address).Cells
        s = s & "," & c.Text
    Next
    Seen = Mid(s, 2)
End Function

Private Function Fresh() As Object
    Set Fresh = Worksheets.Add
End Function

Private Sub Gone(ws As Object)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub
'''
#: Rows and columns of a two-dimensional array, against a range of rows and columns.
ARRAYS = [(1, 1), (1, 2), (2, 1), (2, 2), (3, 3)]
RANGES = ['B2:B5', 'B2:E2', 'B2:D4', 'B2', 'B2:C3']
COPIES = [('A1:B1', 'H1:H3'), ('A1:A2', 'H1:J1'), ('A1:B2', 'H1:K4'), ('A1:B2', 'H1:J3'), ('A1:B2', 'H1:H1'), ('A1:B2', 'H1:I1'), ('A1:B1', 'H1:K1'), ('A1:B1', 'H1:J1'), ('A1:A2', 'H1:H4'), ('A1:A2', 'H1:H3'), ('A1:B2', 'H1:K2'), ('A1:B2', 'H1:I4'), ('A1', 'H1:I2')]
COMPARES = ['#1/2/2020# = "1/2/2020"', '#1/2/2020# = "2020-01-02"', '#1/2/2020# < "1/3/2020"', '"1/2/2020" = #1/2/2020#', '#1/2/2020# = "43832"', '#1/2/2020# = "abc"', '#1/2/2020# = ""', 'd = s', 'd = t', 'v = s', 'v = t', 'd = n', 'v = n', 'd > s', 'v = u', 'd = 43832', '#1/2/2020# = 43832', 'd = True', 's = d', 't = d']
SORTS = [
    ('one_key', 'ws.Range("A1:C5").Sort ws.Range("A1"), xlAscending, Header:=xlYes'),
    ('two_keys', 'ws.Range("A1:C5").Sort ws.Range("A1"), xlAscending, ws.Range("B1"), , xlDescending, Header:=xlYes'),
    ('three_keys', 'ws.Range("A1:C5").Sort ws.Range("A1"), xlAscending, ws.Range("B1"), , xlDescending, ws.Range("C1"), xlAscending, Header:=xlYes'),
    ('no_header', 'ws.Range("A1:C5").Sort ws.Range("A1"), Header:=xlNo'),
    ('match_case', 'ws.Range("A1:C5").Sort ws.Range("A1"), Header:=xlYes, MatchCase:=True'),
    ('left_right', 'ws.Range("A1:C5").Sort ws.Range("A1"), Orientation:=xlLeftToRight, Header:=xlNo'),
    ('twice', 'ws.Range("A1:C5").Sort ws.Range("A1"), Header:=xlYes: ws.Range("A1:C5").Sort ws.Range("B1"), Header:=xlYes'),
    ('smaller', 'ws.Range("A1:C5").Sort ws.Range("A1"), Header:=xlYes: ws.Range("A2:B3").Sort ws.Range("A2"), Header:=xlNo'),
]


def array_probe(rows: int, columns: int, target: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Object, a() As Variant, r As Long, c As Long
    Set ws = Fresh()
    ReDim a(1 To {rows}, 1 To {columns})
    For r = 1 To {rows}: For c = 1 To {columns}: a(r, c) = r * 10 + c: Next: Next
    On Error Resume Next
    ws.Range("{target}").Value = a
    Probe = Err.Number & "!" & Seen(ws, "B2:E5")
    On Error GoTo 0
    Gone ws
End Function
'''


def flat_probe(items: int, target: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Object, a() As Variant, c As Long
    Set ws = Fresh()
    ReDim a(1 To {items})
    For c = 1 To {items}: a(c) = c: Next
    On Error Resume Next
    ws.Range("{target}").Value = a
    Probe = Err.Number & "!" & Seen(ws, "B2:E5")
    On Error GoTo 0
    Gone ws
End Function
'''


def copy_probe(source: str, target: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Object
    Set ws = Fresh()
    ws.Range("A1").Value = "a": ws.Range("B1").Value = "b": ws.Range("A2").Value = "c": ws.Range("B2").Value = "d"
    On Error Resume Next
    ws.Range("{source}").Copy ws.Range("{target}")
    Probe = Err.Number & "!" & Seen(ws, "H1:K4")
    On Error GoTo 0
    Application.CutCopyMode = False
    Gone ws
End Function
'''


def paste_probe(source: str, target: str, how: str) -> str:
    action = f'ws.Range("{target}").PasteSpecial' if how == 'special' else f'ws.Paste ws.Range("{target}")'
    return f'''Function Probe() As String
    Dim ws As Object
    Set ws = Fresh()
    ws.Range("A1").Value = "a": ws.Range("B1").Value = "b": ws.Range("A2").Value = "c": ws.Range("B2").Value = "d"
    ws.Range("{source}").Copy
    On Error Resume Next
    {action}
    Probe = Err.Number & "!" & Seen(ws, "H1:K4")
    On Error GoTo 0
    Application.CutCopyMode = False
    Gone ws
End Function
'''


def compare_probe(line: str) -> str:
    return f'''Function Probe() As String
    Dim d As Date, s As String, t As String, v As Variant, n As Double, u As Variant, r As Variant
    d = #1/2/2020#: s = "1/2/2020": t = "abc": v = #1/2/2020#: n = 43832: u = "1/2/2020"
    On Error Resume Next
    r = ({line})
    Probe = Err.Number & "!" & TypeName(r) & "!" & r
End Function
'''


def sort_probe(line: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Object, s As String, f As Object
    Set ws = Fresh()
    ws.Range("A1:C1").Value = Array("x", "y", "z")
    ws.Range("A2:C2").Value = Array(3, 1, 2): ws.Range("A3:C3").Value = Array(1, 2, 3): ws.Range("A4:C4").Value = Array(2, 3, 1): ws.Range("A5:C5").Value = Array(4, 4, 4)
    {line}
    On Error Resume Next
    s = ws.Sort.SortFields.Count
    s = s & "!" & ws.Sort.Rng.Address: If Err.Number Then s = s & "e" & Err.Number: Err.Clear
    s = s & "!" & ws.Sort.Header & "!" & ws.Sort.MatchCase & "!" & ws.Sort.Orientation & "!" & ws.Sort.SortMethod
    For Each f In ws.Sort.SortFields
        s = s & "!" & f.Key.Address & ":" & f.Order & ":" & f.SortOn & ":" & f.DataOption
        If Err.Number Then s = s & "e" & Err.Number: Err.Clear
    Next
    Probe = s
    On Error GoTo 0
    Gone ws
End Function
'''


def sticky_probe(first: str, second: str) -> str:
    return f'''Function Probe() As String
    Dim ws As Object, s As String, c As Object
    Set ws = Fresh()
    ws.Range("A1").Value = "h": ws.Range("A2").Value = "b": ws.Range("A3").Value = "A": ws.Range("A4").Value = "a": ws.Range("A5").Value = "B"
    ws.Range("A1:A5").Sort ws.Range("A1"), {first}
    ws.Range("A1:A5").Sort ws.Range("A1"), {second}
    For Each c In ws.Range("A1:A5").Cells: s = s & c.Value: Next
    s = s & "!" & ws.Sort.MatchCase & "!" & ws.Sort.Orientation & "!" & ws.Sort.Rng.Address
    Probe = s
    Gone ws
End Function
'''


def closed_probe(line: str) -> str:
    return f'''Function Probe() As String
    Dim o As Object, b As Workbook, v As Variant
    Set o = Workbooks.Add
    Set b = o
    o.Close False
    On Error Resume Next
    v = {line}
    Probe = Err.Number & "!" & TypeName(v)
End Function
'''


def asked() -> dict[str, str]:
    probes: dict[str, str] = {}
    for name, first, second in (('match', 'Header:=xlYes, MatchCase:=True', 'Header:=xlYes'), ('match_off', 'Header:=xlYes, MatchCase:=True', 'Header:=xlYes, MatchCase:=False'),
                                ('header', 'Header:=xlYes', 'Header:=xlGuess'), ('nothing', 'Header:=xlNo', 'Header:=xlNo')):
        probes[f'sticky {name}'] = sticky_probe(first, second)
    for line in ('o.Name', 'b.Name', 'o.Worksheets.Count', 'b.Worksheets.Count', 'o.FullName', 'o.Saved', 'TypeName(o)', 'o Is Nothing', 'o.Close', 'o.Application.Name', 'b.Parent.Name'):
        probes[f'closed {line}'] = closed_probe(line)
    for rows, columns in ARRAYS:
        for target in RANGES:
            probes[f'array {rows}x{columns} {target}'] = array_probe(rows, columns, target)
    for items in (1, 2, 3):
        for target in RANGES:
            probes[f'flat {items} {target}'] = flat_probe(items, target)
    for source, target in COPIES:
        probes[f'copy {source} {target}'] = copy_probe(source, target)
    for source, target in COPIES:
        for how in ('special', 'paste'):
            probes[f'{how} {source} {target}'] = paste_probe(source, target, how)
    for line in COMPARES:
        probes[f'compare {line}'] = compare_probe(line)
    for name, line in SORTS:
        probes[f'sort {name}'] = sort_probe(line)
    return probes


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, object] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for key, source in asked().items():
            if key in answers:
                continue
            result = excel.run_vba(HELPERS + source, 'Probe', timeout=30.0, line_numbers=False)
            assert result.ok and not result.dialogs, (key, result.message, result.dialogs)
            answers[key] = result.value
            print(key.ljust(36), result.value, flush=True)
            OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'helpers': HELPERS, 'sources': asked(), 'answers': answers}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
