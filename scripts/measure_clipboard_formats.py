"""Application.ClipboardFormats: the formats Excel's clipboard holds after each kind of copy, read typed and late bound.

Each case empties the clipboard, makes a workbook with a few cells and a shape and chart where it asks for them, runs its
statements, and records ClipboardFormats as TypeName, bounds and items, read through Application and through an Object.

The rivals, written before measuring: (F1) one list for each kind of copy (cells, a picture, a shape, a chart) whatever the
cells hold and whether cut or copied, or (F2) a list that follows what is copied; on an empty clipboard (E1) an array of one
-1, (E2) an error, or (E3) an empty array; once CutCopyMode is set False (C1) the list stays, or (C2) the clipboard is
empty.

    python scripts/measure_clipboard_formats.py

writes tests/fixtures/clipboard_formats.json, which tests/test_excel_clipboard_formats.py replays. The clipboard is shared
between Excel instances, so this runs under live.lock.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/clipboard_formats.json'
CASES: list[str] = [
    'r = 1',
    'ws.Range("A1").Copy',
    'ws.Range("B1").Copy',
    'ws.Range("A2").Copy',
    'ws.Range("Z99").Copy',
    'ws.Range("A1:B2").Copy',
    'ws.Range("A:A").Copy',
    'ws.Range("A1:A2,C1:C2").Copy',
    'ws.Range("A1").Cut',
    'ws.Range("A1").Copy: Application.CutCopyMode = False',
    'ws.Range("A1").Cut: Application.CutCopyMode = False',
    'ws.Range("A1").Copy: ws.Range("C5").Value = 1',
    'ws.Range("A1").Copy: ws.Paste ws.Range("D1")',
    'ws.Range("A1").Cut: ws.Paste ws.Range("D1")',
    'ws.Range("A1").Cut ws.Range("F1")',
    'ws.Range("A1").Copy ws.Range("F1")',
    'ws.Range("A1").CopyPicture',
    'ws.Range("A1").CopyPicture xlScreen, xlBitmap',
    'ws.Range("A1").CopyPicture xlPrinter, xlPicture',
    'ws.Shapes.AddShape(1, 10, 10, 50, 50).Copy',
    'ws.Shapes.AddShape(1, 10, 10, 50, 50).CopyPicture',
    'Set c = ws.ChartObjects.Add(100, 100, 200, 120): c.Chart.SetSourceData ws.Range("A1:B2"): c.Copy',
    'Set c = ws.ChartObjects.Add(100, 100, 200, 120): c.Chart.SetSourceData ws.Range("A1:B2"): c.Chart.ChartArea.Copy',
    'Set c = ws.ChartObjects.Add(100, 100, 200, 120): c.Chart.SetSourceData ws.Range("A1:B2"): c.Chart.CopyPicture',
    'ws.Range("A1").Copy: Set w2 = Workbooks.Add: w2.Close False',
    'ws.Range("A1").Copy: ws.Range("A1").ClearContents',
    'ws.Range("A1").Copy: ws.Rows(1).Insert',
    'ws.Copy: ActiveWorkbook.Close False',
    'ws.Range("A1").Copy: ws.Range("B1").Copy',
]
# Held out: the rule from the cases above, one list for each kind of copy lasting until cells are copied or the clipboard
# ends, on cases written after it.
CHART = 'Set c = ws.ChartObjects.Add(100, 100, 200, 120): c.Chart.SetSourceData ws.Range("A1:B2"): '
CASES += ['ws.Range("B2").Cut: ws.Range("B2").Copy', CHART + 'c.Copy: ws.Range("A1").Copy', 'ws.Range("A1").Copy: ' + CHART + 'c.Copy',
          'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy', 'ws.Range("A1").Copy: ws.Range("A1").CopyPicture',
          CHART + 'c.Chart.ChartArea.Copy: Application.CutCopyMode = False', CHART + 'c.Copy: Application.CutCopyMode = False',
          'ws.Range("A1:C2").Copy: ws.Range("A5").PasteSpecial xlPasteValues', 'ws.Range("A1").Copy: ws.Range("A1").Value = 9',
          'ws.Range("A1").Copy: ws.Calculate', 'ws.Range("A1").Copy: ws.Range("D4").Select', CHART + 'c.Chart.CopyPicture xlScreen, xlBitmap',
          CHART + 'c.Chart.ChartArea.Copy: ws.Range("A1").Copy']
# After CopyPicture as a bitmap or a chart area's Copy, a copy of cells left ClipboardFormats as it was (cases 32, 41). The
# rivals: (P1) that holds after any picture or chart area, whatever cells are copied or cut after, or (P2) it was timing, and
# the same steps give the cells' list. Each of those two again, twice, and the same with other steps.
CASES += ['ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy', 'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy',
          CHART + 'c.Chart.ChartArea.Copy: ws.Range("A1").Copy', CHART + 'c.Chart.ChartArea.Copy: ws.Range("A1").Copy',
          'ws.Range("A1").CopyPicture: ws.Range("A1").Copy', 'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("B1").Copy',
          CHART + 'c.Chart.ChartArea.Copy: ws.Range("A1").Cut', 'ws.Shapes.AddShape(1, 10, 10, 50, 50).CopyPicture: ws.Range("A1").Copy',
          'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy: ws.Paste ws.Range("E5"): r = ws.Range("E5").Value',
          'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy: Application.CutCopyMode = False',
          'ws.Range("A1").CopyPicture xlScreen, xlBitmap: ws.Range("A1").Copy: ws.Range("B1").Copy']
READS = '''Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long

Private Sub Emptied()
    If OpenClipboard(0) <> 0 Then
        EmptyClipboard
        CloseClipboard
    End If
End Sub

Private Function Listed(v As Variant) As String
    Dim s As String, i As Long
    On Error GoTo Bad
    s = TypeName(v)
    If IsArray(v) Then
        s = s & " " & LBound(v) & " " & UBound(v) & ":"
        For i = LBound(v) To UBound(v)
            s = s & " " & TypeName(v(i)) & " " & v(i)
        Next
    Else
        s = s & " " & CStr(v)
    End If
    Listed = s
    Exit Function
Bad:
    Listed = s & " E" & Err.Number
End Function

Private Function Typed() As String
    Dim v As Variant
    On Error GoTo Bad
    v = Application.ClipboardFormats
    Typed = Listed(v)
    Exit Function
Bad:
    Typed = "E" & Err.Number
End Function

Private Function Late() As String
    Dim o As Object, v As Variant
    On Error GoTo Bad
    Set o = Application
    v = o.ClipboardFormats
    Late = Listed(v)
    Exit Function
Bad:
    Late = "E" & Err.Number
End Function

Private Function First() As String
    On Error GoTo Bad
    First = CStr(Application.ClipboardFormats(1))
    Exit Function
Bad:
    First = "E" & Err.Number
End Function
'''


def probe() -> str:
    lines = [READS, 'Private Sub Act(ByVal which As Long, ws As Worksheet, r As Variant)', '    Dim c As Object, w2 As Object', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '',
              'Public Function Probe(ByVal which As Long) As String',
              '    Dim wb As Workbook, ws As Worksheet, r As Variant, s As String',
              '    Application.DisplayAlerts = False',
              '    Emptied',
              '    Set wb = Workbooks.Add(xlWBATWorksheet)',
              '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1").Value = 5',
              '    ws.Range("B1").Value = "text"',
              '    ws.Range("A2").Formula = "=A1*2"',
              '    ws.Range("B2").Value = 7',
              '    ws.Range("C1:C2").Value = 3',
              '    On Error GoTo Failed',
              '    Act which, ws, r',
              '    s = "ok"',
              'Answered:',
              '    On Error GoTo 0',
              '    s = s & "|" & Application.CutCopyMode & "|" & Typed() & "|" & Late() & "|" & First()',
              '    Application.CutCopyMode = False',
              '    wb.Close False',
              '    Emptied',
              '    Application.DisplayAlerts = True',
              '    Probe = s',
              '    Exit Function',
              'Failed:',
              '    s = "Err " & Err.Number',
              '    Resume Answered',
              'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = probe()
    answers: list[str] = []
    for index in range(len(CASES)):
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                done = excel.run_vba(source, 'Probe', (index,), timeout=120.0, line_numbers=False)
                answers.append(str(done.value) if done.ok and not done.dialogs else f'STOPPED {done.outcome} {done.dialogs}')
        except HarnessError as failure:
            answers.append(f'STOPPED {failure}')
        print(index, CASES[index], '=>', answers[-1])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statement, answer] for statement, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
