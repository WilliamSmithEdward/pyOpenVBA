"""Three old-collection gaps, on the workbook Excel saved for tests/fixtures/old_collection_members.json:

1. AutoScaleFont = True on its first button, loaded from the file: what it reads, and what the save changed against the same book saved
   with nothing set. Rival rules: taken, reads False and writes nothing, as on a new button (rival: it writes an attribute in the VML or
   the drawing).
2. Shadow = True on a picture, the file's first one and one Pictures.Insert adds: what it and the shape's Shadow.Visible read, and what
   the save changed. Rival rules: the effectLst a drawing's Shadow = True writes (outerShdw blurRad 63500, dist 37357, scrgbClr black),
   inside the picture's spPr (rival: the theme's shadow; rival: 438, as Interior and Border are on a picture).
3. DrawingObjects.Visible read, typed and late, with no drawing hidden, with each drawing hidden alone in turn, and with all hidden.
   Rival rules: the first drawing's Visible (rival: the last one's; rival: True unless every one is hidden; rival: the second's).
   A first run gave False only with the sheet's second drawing, its first button, hidden. So a book a macro makes, Rectangle, Oval,
   Line, CheckBox, Button, Button, splits what was left: the second drawing (the Oval), the first form control (the CheckBox) or the first
   button (the fifth).

Nothing is sized. Safety (the Supervisor's rules): one Excel through live.lock, no mail, print, export, encryption, signing,
certificate, IRM, sensitivity-label or identity member, and no window touched.

    python scripts/measure_old_collection_leftovers.py

writes tests/fixtures/old_collection_leftovers.json, which tests/test_excel_old_collection_leftovers.py replays.
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
from measure_old_collection_reads import tiny_png  # noqa: E402

OUT = ROOT / 'tests/fixtures/old_collection_leftovers.json'
SAVED = r'''Public Function Row(ByVal path As String, ByVal saved As String, ByVal picture As String, ByVal which As String) As String
    Dim wb As Workbook, ws As Object, o As Object, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    If which = "new picture" Or which = "new picture base" Then
        Set o = ws.Pictures.Insert(picture)
    ElseIf which = "button" Or which = "button base" Then
        Set o = ws.Buttons(1)
    Else
        Set o = ws.Pictures(1)
    End If
    On Error Resume Next
    If which = "button" Then
        o.AutoScaleFont = True
        e = Err.Number: Err.Clear
        s = "E" & e & "|" & TypeName(o.AutoScaleFont) & " " & o.AutoScaleFont
    ElseIf which = "new picture" Or which = "loaded picture" Then
        o.Shadow = True
        e = Err.Number: Err.Clear
        s = "E" & e & "|" & TypeName(o.Shadow) & " " & o.Shadow & " " & o.ShapeRange.Shadow.Visible
    End If
    If Err.Number Then s = s & " E" & Err.Number
    On Error GoTo 0
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    If which = "button" Then
        s = s & "|reopened " & ws.Buttons(1).AutoScaleFont
    ElseIf which = "new picture" Then
        s = s & "|reopened " & ws.Pictures(ws.Pictures.Count).Shadow
    ElseIf which = "loaded picture" Then
        s = s & "|reopened " & ws.Pictures(1).Shadow
    End If
    If Err.Number Then s = s & " E" & Err.Number
    wb.Close False
    Row = s
End Function

Private Function Whole(o As {declared}) As String
    Dim v As Variant
    On Error Resume Next
    v = o.Visible
    If Err.Number Then Whole = "E" & Err.Number Else Whole = TypeName(v) & " " & v
End Function

Public Function Visibles(ByVal path As String) As String
    Dim wb As Workbook, ws As Object, i As Long, j As Long, n As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    n = ws.DrawingObjects.Count
    For i = 1 To n
        ws.DrawingObjects(i).Visible = True
    Next
    s = n & "|none " & Whole(ws.DrawingObjects)
    For i = 1 To n
        ws.DrawingObjects(i).Visible = False
        s = s & "|" & i & " " & TypeName(ws.DrawingObjects(i)) & " " & Whole(ws.DrawingObjects)
        ws.DrawingObjects(i).Visible = True
    Next
    For i = 1 To n
        ws.DrawingObjects(i).Visible = False
    Next
    s = s & "|all " & Whole(ws.DrawingObjects)
    wb.Close False
    Visibles = s
End Function

Public Function Made() As String
    Dim wb As Workbook, ws As Object, i As Long, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Rectangles.Add 10, 10, 40, 20: ws.Ovals.Add 60, 10, 40, 20: ws.Lines.Add 110, 10, 150, 30
    ws.CheckBoxes.Add 160, 10, 40, 20: ws.Buttons.Add 210, 10, 40, 20: ws.Buttons.Add 260, 10, 40, 20
    s = ws.DrawingObjects.Count & "|none " & Whole(ws.DrawingObjects)
    For i = 1 To ws.DrawingObjects.Count
        ws.DrawingObjects(i).Visible = False
        s = s & "|" & i & " " & TypeName(ws.DrawingObjects(i)) & " " & Whole(ws.DrawingObjects)
        ws.DrawingObjects(i).Visible = True
    Next
    wb.Close False
    Made = s
End Function
'''
WHICH = ('button', 'new picture', 'loaded picture')


def source(binding: str) -> str:
    return SAVED.replace('{declared}', 'Excel.DrawingObjects' if binding == 'Typed' else 'Object')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: dict[str, object] = {}
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        folder = Path(held)
        picture = folder / 'dot.png'
        picture.write_bytes(tiny_png())
        book = folder / 'in.xlsx'
        cases: list[dict[str, object]] = []
        for which in WHICH:
            base_which = 'new picture base' if which == 'new picture' else 'button base' if which == 'button' else 'loaded base'
            book.write_bytes(BOOK.read_bytes())
            base = folder / f'base_{len(cases)}.xlsx'
            ran = excel.run_vba(source('Late'), 'Row', (str(book), str(base), str(picture), base_which), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran.outcome
            book.write_bytes(BOOK.read_bytes())
            out = folder / f'out_{len(cases)}.xlsx'
            ran = excel.run_vba(source('Late'), 'Row', (str(book), str(out), str(picture), which), timeout=60.0, line_numbers=False)
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            before, after = parts(base), parts(out) if out.exists() else {}
            changed = {name: changes(before.get(name, ''), text) for name, text in after.items() if before.get(name) != text}
            cases.append({'which': which, 'excel': answer, 'parts': changed})
            print(which, answer, sorted(changed), flush=True)
        found['saves'] = cases
        visibles: dict[str, str] = {}
        for binding in ('Typed', 'Late'):
            book.write_bytes(BOOK.read_bytes())
            ran = excel.run_vba(source(binding), 'Visibles', (str(book),), timeout=60.0, line_numbers=False)
            visibles[binding] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            print(binding, visibles[binding], flush=True)
        found['visibles'] = visibles
        made: dict[str, str] = {}
        for binding in ('Typed', 'Late'):
            ran = excel.run_vba(source(binding), 'Made', (), timeout=60.0, line_numbers=False)
            made[binding] = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            print('made', binding, made[binding], flush=True)
        found['made'] = made
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SAVED, **found}, indent=1) + '\n', encoding='utf-8')
    print('wrote', OUT.name, flush=True)


if __name__ == '__main__':
    main()
