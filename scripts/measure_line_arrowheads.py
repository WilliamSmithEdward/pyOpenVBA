"""A line's old ArrowHeadStyle, ArrowHeadLength and ArrowHeadWidth, on the first line of the workbook Excel saved for
tests/fixtures/old_collection_members.json and on one a macro adds: what the old object and the shape's LineFormat read in the session
and after a reopen, and what the save changed in the drawing part against the same workbook saved with nothing set, tag by tag.

Rival rules, written before measuring:

- ArrowHeadStyle sets the end the line was drawn to (tailEnd) only, as the one measure showed (Open wrote tailEnd type="arrow"); rival:
  both ends. The styles map to LineFormat's: Open to msoArrowheadOpen ("arrow"), Closed to msoArrowheadTriangle, DoubleOpen to
  msoArrowheadStealth, DoubleClosed to msoArrowheadDiamond, None to none (rival: DoubleClosed to oval).
- ArrowHeadLength and ArrowHeadWidth set both ends, as measured for Long and Wide: Short "sm", Medium "med", Long "lg" (rival: Medium
  writes nothing, being the default).
- The old reads are the end's: ArrowHeadStyle -4142 while the end has none, whatever the length.
- A loaded line takes them as a new one does, patching its a:ln.

Nothing is sized, so the display's DPI does not enter. Safety (the Supervisor's rules): one Excel through live.lock, no mail, print,
export, encryption, signing, certificate, IRM, sensitivity-label or identity member, and no window touched.

    python scripts/measure_line_arrowheads.py

writes tests/fixtures/line_arrowheads.json, which tests/test_excel_line_arrowheads.py replays.
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

OUT = ROOT / 'tests/fixtures/line_arrowheads.json'
#: Each case's statements on the line ``o``.
SETS: dict[str, str] = {
    **{f'style {name}': f'o.ArrowHeadStyle = {value}' for name, value in (('None', -4142), ('Open', 2), ('Closed', 3), ('DoubleOpen', 4), ('DoubleClosed', 5))},
    **{f'length {name}': f'o.ArrowHeadLength = {value}' for name, value in (('Short', 1), ('Medium', -4138), ('Long', 3))},
    **{f'width {name}': f'o.ArrowHeadWidth = {value}' for name, value in (('Narrow', 1), ('Medium', -4138), ('Wide', 3))},
    'style then length': 'o.ArrowHeadStyle = 3: o.ArrowHeadLength = 1',
    'style then none': 'o.ArrowHeadStyle = 2: o.ArrowHeadStyle = -4142',
    'all three': 'o.ArrowHeadStyle = 5: o.ArrowHeadLength = 3: o.ArrowHeadWidth = 1',
}
SOURCE = r'''Private Function Item(ws As Object, ByVal fresh As Boolean, ByVal add As Boolean) As Object
    If fresh And add Then
        Set Item = ws.Lines.Add(500, 300, 560, 340)
    ElseIf fresh Then
        Set Item = ws.Lines(ws.Lines.Count)
    Else
        Set Item = ws.Lines(1)
    End If
End Function

Private Function Reads(o As Object) As String
    Dim s As String, l As Object
    On Error Resume Next
    s = o.ArrowHeadStyle & "," & o.ArrowHeadLength & "," & o.ArrowHeadWidth
    If Err.Number Then s = "E" & Err.Number
    Err.Clear
    Set l = o.ShapeRange.Line
    s = s & " begin " & l.BeginArrowheadStyle & "," & l.BeginArrowheadLength & "," & l.BeginArrowheadWidth
    s = s & " end " & l.EndArrowheadStyle & "," & l.EndArrowheadLength & "," & l.EndArrowheadWidth
    If Err.Number Then s = s & " E" & Err.Number
    Reads = s
End Function

Public Function Row(ByVal path As String, ByVal saved As String, ByVal fresh As Boolean, ByVal which As Long) As String
    Dim wb As Workbook, ws As Object, o As Object, e As Long, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set ws = wb.Worksheets(1)
    Set o = Item(ws, fresh, True)
    If which >= 0 Then
        On Error Resume Next
        Select Case which
%CASES%
        End Select
        e = Err.Number
        On Error GoTo 0
        s = Reads(o)
    End If
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = "E" & e & "|" & s & "|reopened " & Reads(Item(wb.Worksheets(1), fresh, False))
    wb.Close False
    Row = s
End Function
'''


def source() -> str:
    return SOURCE.replace('%CASES%', '\n'.join(f'        Case {index}\n            {statements}' for index, statements in enumerate(SETS.values())))


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found: list[dict[str, object]] = []
    code = source()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        folder = Path(held)
        for fresh in (False, True):
            book = folder / 'in.xlsx'
            book.write_bytes(BOOK.read_bytes())
            base = folder / f'base_{fresh}.xlsx'
            ran = excel.run_vba(code, 'Row', (str(book), str(base), fresh, -1), timeout=60.0, line_numbers=False)
            assert ran.ok and not ran.dialogs, ran.outcome
            base_parts = parts(base)
            for index, (name, statements) in enumerate(SETS.items()):
                book.write_bytes(BOOK.read_bytes())
                out = folder / f'out_{fresh}_{index}.xlsx'
                ran = excel.run_vba(code, 'Row', (str(book), str(out), fresh, index), timeout=60.0, line_numbers=False)
                answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome}'
                after = parts(out) if out.exists() else {}
                changed = {part: changes(base_parts.get(part, ''), text) for part, text in after.items() if base_parts.get(part) != text}
                found.append({'fresh': fresh, 'case': name, 'statements': statements, 'excel': answer, 'parts': changed})
                print(fresh, name, answer, sorted(changed), flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': code, 'cases': found}, indent=1) + '\n', encoding='utf-8')
    print('wrote', len(found), 'cases', flush=True)


if __name__ == '__main__':
    main()
