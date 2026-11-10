"""A drawing's text linked to a cell through its old object's Formula, and AutoScaleFont, on the first rectangle, oval, arc and text box of
the workbook Excel saved for tests/fixtures/old_collection_members.json and on one a macro adds: what the item reads in the session and
after a reopen, and what the save changed in the drawing part against the same workbook saved with the link left out, tag by tag.

Cells: A1 empty, B2 "hello", B3 1234.5 shown as #,##0.00, B4 "bold" in bold red.

Rival rules, written before measuring:

- Formula reads back the reference as given, without its "=", relative or absolute as written, with a space after it in the session as
  after a reopen (rival: the space only after a reopen; rival: made absolute).
- The text becomes the cell's shown text, " " for an empty cell (rival: its value unformatted).
- The linked run is an a:fld of type TxLink whose rPr carries the cell's font: b="1" and FF0000 for B4 (rival: b="0" and 000000 always,
  as the empty A1 gave).
- Changing the cell after the link changes the text in the session and in the file (rival: the text stays as linked).
- Formula = "" after a link takes it off: textlink="" and a plain run with the text it showed (rival: the text goes too).
- A sheet-qualified reference reads back with the sheet's name (rival: without it).
- AutoScaleFont = True is taken, reads False and writes nothing, as AddIndent (rival: it writes an attribute).

Text and links only, nothing sized, so the display's DPI does not enter: each part is kept as its change from a save of the same book at
the same scale. StandardHeight is recorded. Safety (the Supervisor's rules): one Excel through live.lock, no mail, print, export,
encryption, signing, certificate, IRM, sensitivity-label or identity member, and no window touched.

    python scripts/measure_drawing_text_links.py

writes tests/fixtures/drawing_text_links.json, which tests/test_excel_drawing_text_links.py replays.
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

OUT = ROOT / 'tests/fixtures/drawing_text_links.json'
KINDS = ('Rectangles', 'Ovals', 'Arcs', 'TextBoxes')
#: What each case sets: the Formula, then what follows it ('' for nothing).
LINKS: dict[str, tuple[str, str]] = {
    'empty': ('=$A$1', ''),
    'text': ('=$B$2', ''),
    'number': ('=$B$3', ''),
    'font': ('=$B$4', ''),
    'relative': ('=B2', ''),
    'sheet': ("='{sheet}'!$B$2", ''),
    'changed': ('=$B$2', 'ws.Range("B2").Value = "changed"'),
    'cleared': ('=$B$2', 'o.Formula = ""'),
}
SOURCE = r'''Private Function Item(ws As Object, ByVal kind As String, ByVal fresh As Boolean, ByVal add As Boolean) As Object
    If fresh And add Then
        If kind = "Arcs" Then
            Set Item = CallByName(ws, kind, VbMethod).Add(500, 300, 560, 340)
        Else
            Set Item = CallByName(ws, kind, VbMethod).Add(500, 300, 60, 40)
        End If
    ElseIf fresh Then
        Set Item = CallByName(ws, kind, VbMethod, CallByName(ws, kind, VbMethod).Count)
    Else
        Set Item = CallByName(ws, kind, VbMethod, 1)
    End If
End Function

Private Function Reads(o As Object) As String
    Dim s As String
    On Error Resume Next
    s = "F[" & o.Formula & "]"
    If Err.Number Then s = "FE" & Err.Number
    Err.Clear
    s = s & " T[" & o.Text & "]"
    If Err.Number Then s = s & " TE" & Err.Number
    Err.Clear
    s = s & " C[" & o.Characters.Text & "]"
    If Err.Number Then s = s & " CE" & Err.Number
    Err.Clear
    s = s & " A[" & o.AutoScaleFont & "]"
    If Err.Number Then s = s & " AE" & Err.Number
    Reads = s
End Function

Public Function Row(ByVal path As String, ByVal saved As String, ByVal kind As String, ByVal fresh As Boolean, ByVal formula As String, _
                    ByVal after As String, ByVal scaled As Boolean, ByVal change As Boolean) As String
    Dim wb As Workbook, ws As Object, o As Object, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    ws.Range("B2").Value = "hello"
    ws.Range("B3").Value = 1234.5
    ws.Range("B3").NumberFormat = "#,##0.00"
    ws.Range("B4").Value = "bold"
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Font.Color = RGB(255, 0, 0)
    Set o = Item(ws, kind, fresh, True)
    If change Then
        On Error Resume Next
        If scaled Then
            o.AutoScaleFont = True
        Else
            o.Formula = Replace(formula, "{sheet}", ws.Name)
        End If
        e = Err.Number
        Err.Clear
        If after = "ws.Range(""B2"").Value = ""changed""" Then ws.Range("B2").Value = "changed"
        If after = "o.Formula = """"" Then o.Formula = ""
        If Err.Number Then e = 1000000 + Err.Number
        On Error GoTo 0
        s = Reads(o)
    End If
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = "E" & e & "|" & s & "|reopened " & Reads(Item(wb.Worksheets(1), kind, fresh, False)) & "|" & wb.Worksheets(1).StandardHeight
    wb.Close False
    Row = s
End Function
'''


def cases() -> list[dict[str, object]]:
    found: list[dict[str, object]] = []
    for kind in KINDS:
        for fresh in (False, True):
            for name, (formula, after) in LINKS.items():
                found.append({'kind': kind, 'fresh': fresh, 'case': name, 'formula': formula, 'after': after, 'scaled': False})
            found.append({'kind': kind, 'fresh': fresh, 'case': 'AutoScaleFont', 'formula': '', 'after': '', 'scaled': True})
    return found


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        folder = Path(held)
        bases: dict[tuple[str, bool], dict[str, str]] = {}
        for number, case in enumerate(cases()):
            kind, fresh = str(case['kind']), bool(case['fresh'])
            book = folder / 'in.xlsx'
            if (kind, fresh) not in bases:
                book.write_bytes(BOOK.read_bytes())
                base = folder / f'base_{kind}_{fresh}.xlsx'
                ran = excel.run_vba(SOURCE, 'Row', (str(book), str(base), kind, fresh, '', '', False, False), timeout=60.0, line_numbers=False)
                assert ran.ok and not ran.dialogs, ran.outcome
                bases[(kind, fresh)] = parts(base)
            book.write_bytes(BOOK.read_bytes())
            out = folder / f'out_{number}.xlsx'
            ran = excel.run_vba(SOURCE, 'Row', (str(book), str(out), kind, fresh, case['formula'], case['after'], case['scaled'], True), timeout=60.0,
                                line_numbers=False)
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
            after = parts(out) if out.exists() else {}
            base_parts = bases[(kind, fresh)]
            changed = {name: changes(base_parts.get(name, ''), part) for name, part in after.items() if base_parts.get(name) != part}
            found.append({**case, 'excel': answer, 'parts': changed})
            print(kind, fresh, case['case'], answer, sorted(changed), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': SOURCE, 'cases': found}, indent=1) + '\n', encoding='utf-8')
    print('wrote', len(found), 'cases', flush=True)


if __name__ == '__main__':
    main()
