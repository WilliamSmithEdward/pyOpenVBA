"""The columns ListObjects.Add widens: a table made over cells whose header row Excel fills with Column1, Column2 and so
on widens their columns. First look, held at the current display DPI.

Each case makes a workbook with one sheet, runs its statements, and records r, then for each of columns A to M its Width
in points before and after, the header cell's text in the table, and the Width Excel's own AutoFit gives a scratch column
holding that text, in the plain Normal font and in bold. StandardHeight is recorded at both ends, since every width here
depends on the display's DPI.

Predictions, written before measuring (the dropdown is 16 pixels at 96 DPI in a pivot's header, _pivot_fit.DROPDOWN):
- W1: a header cell Excel fills itself widens its column to the bold AutoFit width of its text plus the dropdown, in
  pixels; a column wider already keeps its width. Rival W2: every header cell is fitted, text a macro typed included.
  Rival W3: the plain AutoFit width plus the dropdown. Predict W1.
- D1: the dropdown is 16 pixels times DPI/96, 24 at 144 DPI. Rival D2: 16 at any DPI. Predict D1.
- N1: a column is never narrowed. Rival N2: fitted exactly, narrower too. Predict N1.
- H1: a hidden column stays hidden at width 0. Predict H1.
- Resize and ListColumns.Add widen the new columns as Add does (R1); rival R2: nothing widens. Predict R1.

What the first look showed: every header cell widens its column, text a macro typed included (W2); never narrower (N1);
a hidden column stays hidden (H1); Resize and ListColumns.Add widen nothing (R2). At 144 DPI the widened column is the
plain AutoFit width plus 26, 26 and 27 pixels, and the bold one plus 22, 22 and 20; SM_CXVSCROLL is 17 at 96 DPI and 26
at 144. Predictions for look 2 (cases 20 and 21), written before it ran:
- A1: the widened width is the plain AutoFit width plus SM_CXVSCROLL, 26 here, for every text. Rival A2: the button is
  added to the text before AutoFit's formula, so the difference is 26 or 27 by text. Rival A3: bold text plus a
  constant. Predict A2.
- F1: AutoFit of a table's column afterwards gives back the widened width, button included. Rival F2: the plain
  AutoFit width, narrower than the widened one. Predict F1.

What look 2 showed: A1, A2 and A3 all wrong as written. The widened column is the bold AutoFit width plus 22 pixels for 27
of the 30 texts and plus 21 for the three widest (MMMMMMMMMM, Postal code 12, Total amount due); look 1's longest header
took 20. AutoFit afterwards makes each column 3 or 4 pixels wider again, so neither F1 nor F2. Predictions for look 3
(cases 22 to 29), written before it ran:
- P1: a pivot's Row Labels column at this DPI is its bold AutoFit width plus the same 22, the dropdown the model already
  has as 16 pixels at 96 DPI. Rival P2: plus 24, 16 scaled by 1.5. Predict P1.
- S1: a style whose header row is not bold widens to the plain AutoFit width plus 22. Rival S2: bold whatever the
  style. Predict S1. A table made with no style (case 24) is measured plain under S1.
- D1: only the header cell counts; a long value under it widens nothing. Rival D2: the data cells count. Predict D1.
- B1: a header cell bold or italic itself measures as its own font does, bold kept; italic adds its reach. Predict B1.
- Z1: a bigger font in the header cell widens to its own bold AutoFit plus 22 again. Predict Z1.

What look 3 showed:
- P1 and P2 both wrong. The pivot's Row Labels column is its bold AutoFit width plus 25 pixels, so a table's header is
  not a pivot's.
- S unresolved. TableStyleLight8 and TableStyleMedium9 both draw their header bold, and so widen as Medium2 does. An
  empty TableStyleName fails with error 5 and makes no table.
- D1 held: the data cells count for nothing.
- B1 held for bold. An italic header widens to its bold upright width plus 23, not its italic AutoFit width.
- Z1 wrong. At 16pt the column is the bold AutoFit width plus 20, and at 20pt plus 17. AutoFit's pad for a big font
  goes into its width, not into the widening.
So the widening is the header text's bold width plus a dropdown, without AutoFit's italic reach or big-font pad. Its
pixels at 96 DPI are not known from these numbers.

    python scripts/measure_table_widen.py

writes tests/fixtures/table_widen.json.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig
from pyvbaharness.results import HarnessError

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/table_widen.json'


def added(source: str, headers: str = 'xlYes', style: str = '') -> str:
    tail = f', , "{style}"' if style else ''
    return f'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("{source}"), , {headers}{tail})'


CASES = [
    added('A1:E2'),
    added('A1:L2'),
    'ws.Range("A1").Value = "x": ws.Range("B1").Value = "y": ' + added('A1:B3'),
    'ws.Range("A1").Value = "a much longer header": ws.Range("B1").Value = "y": ' + added('A1:B3'),
    'ws.Range("A1").Value = "Column1": ws.Range("B1").Value = "Column2": ' + added('A1:B3'),
    'ws.Columns(1).ColumnWidth = 30: ' + added('A1:B2'),
    'ws.Columns(1).ColumnWidth = 3: ' + added('A1:B2'),
    'ws.Range("A1:B2").Value = 5: ' + added('A1:B2', 'xlNo'),
    'ws.Range("A1").Value = "name": ws.Range("A2:B3").Value = 1: ' + added('A1:B3'),
    added('A1:B2', style='TableStyleLight1'),
    'ws.Range("A1:B1").Font.Size = 20: ' + added('A1:B2'),
    'ws.Columns(2).Hidden = True: ' + added('A1:C2'),
    added('A1:B2', 'xlGuess'),
    'ws.Range("A1:C1").Value = 5: ' + added('A1:C3', 'xlNo'),
    added('A1:B2') + ': lo.Resize ws.Range("A1:D2")',
    added('A1:B2') + ': lo.ListColumns.Add',
    added('A1:B2') + ': lo.ShowAutoFilter = False',
    'ws.Range("A1:B2").Value = 5: ' + added('A1:B3', 'xlNo'),
    'ws.Range("A1").Value = "x": ' + added('A1:B3'),
    'ws.Columns(1).ColumnWidth = 11: ' + added('A1:B2'),
    # Look 2: thirty typed headers over columns made narrow first, so each widened width shows whole.
    'ws.Columns("A:AD").ColumnWidth = 0.5: ws.Range("A1:AD1").Value = HEADERS: ' + added('A1:AD2'),
    'ws.Columns("A:AD").ColumnWidth = 0.5: ws.Range("A1:AD1").Value = HEADERS: ' + added('A1:AD2')
    + ': ws.Columns("A:AD").AutoFit',
    # Look 3: a pivot's dropdown at this DPI, header rows that are not bold, data cells, a header bold or italic itself.
    'ws.Range("A1:B3").Value = ws.Evaluate("{""k"",""v"";""a"",1;""b"",2}"): '
    'Set pt = ws.Parent.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("D1"), "P"): '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: '
    'r = ws.Range("D1").Text & "/" & ws.Range("D4").Text & "/" & ws.Range("E1").Text',
    added('A1:B2', style='TableStyleLight8') + ': r = lo.TableStyle.TableStyleElements(xlHeaderRow).Font.Bold',
    'Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes, , ""): r = TypeName(lo.TableStyle)',
    'ws.Range("A1").Value = "x": ws.Range("A2").Value = "a long value in the data": ' + added('A1:A2'),
    'ws.Range("A1:B1").Font.Bold = True: ' + added('A1:B2'),
    'ws.Range("A1:B1").Font.Italic = True: ' + added('A1:B2'),
    added('A1:B2', style='TableStyleMedium9') + ': r = lo.TableStyle.TableStyleElements(xlHeaderRow).Font.Bold',
    'ws.Range("A1:B1").Font.Size = 16: ' + added('A1:B2'),
]
#: Thirty texts of many widths, each its own header.
TEXTS = ['i', 'W', 'ab', 'Id', 'Qty', 'Name', 'mmmm', 'Price', 'iiiiiiii', 'Total', 'Amount', 'WWWWW', 'Region',
         'Column1', 'Customer', 'Quantity', 'Description', 'abcdefghij', 'Unit price', 'Order date', 'llllllllllll',
         'Ship country', 'Sales person', 'MMMMMMMMMM', 'Product name', 'Discount rate', 'Postal code 12', 'xyz',
         'Total amount due', 'Notes']

COLUMNS = 30

PROBE = '''Private Function Widths(ws As Worksheet) As String
    Dim s As String, c As Long
    For c = 1 To %(columns)d
        s = s & IIf(c > 1, ",", "") & ws.Columns(c).Width
    Next
    Widths = s
End Function

Private Function Fits(ws As Worksheet, scratch As Worksheet) As String
    Dim s As String, c As Long, t As String
    For c = 1 To %(columns)d
        t = ws.Cells(1, c).Text
        s = s & " " & c & "[" & t & "]"
        If t <> "" Then
            scratch.Cells(1, 1).Value = "'" & t
            scratch.Cells(1, 1).Font.Size = ws.Cells(1, c).Font.Size
            scratch.Cells(1, 1).Font.Italic = ws.Cells(1, c).Font.Italic
            scratch.Cells(1, 1).Font.Bold = False
            scratch.Columns(1).AutoFit
            s = s & scratch.Columns(1).Width
            scratch.Cells(1, 1).Font.Bold = True
            scratch.Columns(1).AutoFit
            s = s & "/" & scratch.Columns(1).Width
        End If
    Next
    Fits = s
End Function

Private Sub Act(ByVal which As Long, ws As Worksheet, r As Variant)
    Dim lo As ListObject, pt As PivotTable
    Select Case which
%(cases)s
    End Select
End Sub

Public Function Probe(ByVal which As Long) As String
    Dim wb As Workbook, ws As Worksheet, scratch As Worksheet, r As Variant, s As String, before As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    s = "sh " & ws.StandardHeight
    before = Widths(ws)
    On Error GoTo Failed
    Act which, ws, r
    s = s & "|ok"
Answered:
    On Error Resume Next
    s = s & "|" & CStr(r) & "|before " & before & "|after " & Widths(ws)
    Set scratch = wb.Worksheets.Add(After:=ws)
    s = s & "|fits" & Fits(ws, scratch)
    If ws.ListObjects.Count Then s = s & "|" & ws.ListObjects(1).Range.Address & " " & ws.ListObjects(1).ShowHeaders & " " & ws.ListObjects(1).TableStyle
    s = s & "|hidden " & ws.Columns(2).Hidden & "|sh " & ws.StandardHeight
    Err.Clear
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
    Exit Function
Failed:
    s = s & "|Err " & Err.Number
    Resume Answered
End Function
'''


def probe() -> str:
    headers = 'Array(' + ', '.join(f'"{text}"' for text in TEXTS) + ')'
    cases = '\n'.join(f'    Case {index}\n        {statement.replace("HEADERS", headers)}'
                      for index, statement in enumerate(CASES))
    return PROBE % {'columns': COLUMNS, 'cases': cases}


def main() -> None:
    source = probe()
    answers: list[str] = []
    chosen = [int(one) for one in sys.argv[1:]] or list(range(len(CASES)))
    for index in chosen:
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                done = excel.run_vba(source, 'Probe', (index,), timeout=120.0, line_numbers=False)
                answers.append(str(done.value) if done.ok and not done.dialogs
                               else f'STOPPED {done.outcome} {done.error} {done.message} {done.dialogs}')
        except HarnessError as failure:
            answers.append(f'STOPPED {failure}')
        print(index, CASES[index][:90], '=>', answers[-1], flush=True)
    if len(chosen) < len(CASES):
        return
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[statement, answer] for statement, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
