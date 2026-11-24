"""The last old-collection members not measured, on the workbook Excel saved for tests/fixtures/old_collection_members.json:

1. DrawingObjects.Vertices, Reshape and Ungroup, typed and late, on that book (no group) and on it with two drawings grouped. Rival
   rules, written before measuring: Vertices is 1004 as most reads over the items are (rival: 438); Reshape is 1004 (rival: 438; rival:
   449 for the missing arguments, given none); Ungroup with no group is 1004 (rival: nothing happens), and with a group it ungroups it,
   returning the drawings (rival: 1004 because the other items are not groups).
2. Formula on a button, the file's first one and one Buttons.Add adds, set to "=$B$2" with B2 "hello": what it, its Caption and Text read,
   in the session and after a reopen, and what the save changed against the same book saved with the button left alone. Rival rules: the
   caption shows the cell, the VML gets FmlaTxbx and the drawing a textlink (rival: 1004, as the collection's set; rival: taken and lost).

Nothing is sized. Safety (the Supervisor's rules): one Excel through live.lock, no mail, print, export, encryption, signing,
certificate, IRM, sensitivity-label or identity member, and no window touched.

    python scripts/measure_old_collection_odds.py

writes tests/fixtures/old_collection_odds.json, which tests/test_excel_old_collection_odds.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from measure_loaded_control_edits import BOOK, parts  # noqa: E402
from measure_loaded_drawing_edits import changes  # noqa: E402

OUT = ROOT / 'tests/fixtures/old_collection_odds.json'
METHODS = r'''Private Function Tried(o As {declared}, ByVal which As String) As String
    Dim v As Variant, r As Object
    On Error Resume Next
    If which = "Vertices" Then
        v = o.Vertices
        If Err.Number = 0 Then Tried = "ok " & TypeName(v)
    ElseIf which = "Reshape" Then
        o.Reshape 1, True
        If Err.Number = 0 Then Tried = "ok"
    Else
        Set r = o.Ungroup
        If Err.Number = 0 Then Tried = "ok " & TypeName(r) & " " & r.Count
    End If
    If Err.Number Then Tried = "E" & Err.Number
End Function

Public Function Row(ByVal path As String, ByVal grouped As Boolean, ByVal which As String) As String
    Dim wb As Workbook, ws As Object, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    If grouped Then ws.Shapes.Range(Array(ws.Rectangles(1).Name, ws.Ovals(1).Name)).Group
    s = ws.DrawingObjects.Count & "|" & Tried(ws.DrawingObjects, which) & "|" & ws.DrawingObjects.Count & " " & ws.Shapes.Count
    wb.Close False
    Row = s
End Function
'''
FORMULA = r'''Public Function Row(ByVal path As String, ByVal saved As String, ByVal fresh As Boolean, ByVal change As Boolean) As String
    Dim wb As Workbook, ws As Object, o As Object, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    ws.Range("B2").Value = "hello"
    If fresh Then Set o = ws.Buttons.Add(500, 300, 60, 30) Else Set o = ws.Buttons(1)
    If change Then
        On Error Resume Next
        o.Formula = "=$B$2"
        e = Err.Number: Err.Clear
        s = "E" & e & "|F[" & o.Formula & "] C[" & o.Caption & "] T[" & o.Text & "]"
        If Err.Number Then s = s & " E" & Err.Number
        On Error GoTo 0
    End If
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    Set ws = wb.Worksheets(1)
    If fresh Then Set o = ws.Buttons(ws.Buttons.Count) Else Set o = ws.Buttons(1)
    On Error Resume Next
    s = s & "|reopened F[" & o.Formula & "] C[" & o.Caption & "] T[" & o.Text & "]"
    If Err.Number Then s = s & " E" & Err.Number
    wb.Close False
    Row = s
End Function
'''


def methods_source(binding: str) -> str:
    return METHODS.replace('{declared}', 'Excel.DrawingObjects' if binding == 'Typed' else 'Object')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, object] = {'methods_source': METHODS, 'formula_source': FORMULA}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        folder = Path(held)
        book = folder / 'in.xlsx'
        methods: list[dict[str, object]] = []
        for which in ('Vertices', 'Reshape', 'Ungroup'):
            for grouped in (False, True):
                for binding in ('Typed', 'Late'):
                    book.write_bytes(BOOK.read_bytes())
                    ran = excel.run_vba(methods_source(binding), 'Row', (str(book), grouped, which), timeout=60.0, line_numbers=False)
                    answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                    methods.append({'which': which, 'grouped': grouped, 'binding': binding, 'excel': answer})
                    print(which, grouped, binding, answer, flush=True)
        record['methods'] = methods
        formulas: list[dict[str, object]] = []
        for fresh in (False, True):
            book.write_bytes(BOOK.read_bytes())
            base = folder / f'base_{fresh}.xlsx'
            ran = excel.run_vba(FORMULA, 'Row', (str(book), str(base), fresh, False), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran.outcome
            book.write_bytes(BOOK.read_bytes())
            out = folder / f'out_{fresh}.xlsx'
            ran = excel.run_vba(FORMULA, 'Row', (str(book), str(out), fresh, True), timeout=60.0, line_numbers=False)
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            before, after = parts(base), parts(out) if out.exists() else {}
            changed = {name: changes(before.get(name, ''), text) for name, text in after.items() if before.get(name) != text}
            formulas.append({'fresh': fresh, 'excel': answer, 'parts': changed})
            print('formula', fresh, answer, sorted(changed), flush=True)
        record['formulas'] = formulas
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), **record}, indent=1) + '\n', encoding='utf-8')
    print('wrote', OUT.name, flush=True)


if __name__ == '__main__':
    main()
