"""Application.DoubleClick under a macro: what it does to the active cell. First look.

Each case makes a workbook with one sheet, runs its statements with a sink class holding the workbook WithEvents, and
records the events the sink saw, r, whether Excel is Ready afterwards, its Visible, and a digest of the workbook: its
sheets, the active sheet and cell, and A1:F8's values. Nothing is saved, and no cell holds a hyperlink.

Predictions, written before measuring:
- E1: DoubleClick raises SheetBeforeDoubleClick once, its Target the active cell and Cancel False, and does nothing
  else to a cell holding a value, a formula or nothing: Excel stays Ready. Rival E2: it raises nothing in an Excel
  that is not shown. Rival E3: it puts the cell in edit mode, so Ready turns False. Predict E1.
- C1: Cancel set in the handler changes nothing that can be seen, as nothing else happens anyway. Predict C1.
- P1: on a pivot table's value, DoubleClick shows the value's detail on a new sheet before the active one, as
  ShowDetail does, raising NewSheet after BeforeDoubleClick; Cancel stops it. Rival P2: nothing. Predict P1.
- V1: with EnableEvents off, no event, and a pivot value's detail is shown all the same. Predict V1.

What the first run showed (cases 0 to 6; case 7 then stopped the run):
- E1 was wrong, and so were its rivals. Over a value, a formula or nothing, DoubleClick does nothing at all, raising
  no BeforeDoubleClick, and Excel stays Ready.
- P1 held, without the event. A pivot value's detail goes on a new sheet before the active one. NewSheet, Activate
  and SelectionChange fire, and the detail's records are selected. Cancel cannot stop it, since no
  BeforeDoubleClick comes.
- V1 held.
- A row label (D2, case 7 as first written) opened Excel's Show Detail dialog, which the harness cannot answer. The
  case now double-clicks the grand total's value instead, unmeasured until this run.

    python scripts/measure_double_click.py

writes tests/fixtures/double_click.json.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/double_click.json'
SINK = '''Public WithEvents Book As Workbook

Private Sub Book_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Log = Log & "[BeforeDoubleClick " & Sh.Name & " " & Target.Address(False, False) & " " & Cancel & "]"
    If WantCancel Then Cancel = True
End Sub

Private Sub Book_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "[Change " & Sh.Name & " " & Target.Address(False, False) & "]"
End Sub

Private Sub Book_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "[SelectionChange " & Sh.Name & " " & Target.Address(False, False) & "]"
End Sub

Private Sub Book_NewSheet(ByVal Sh As Object)
    Log = Log & "[NewSheet " & Sh.Name & "]"
End Sub

Private Sub Book_SheetActivate(ByVal Sh As Object)
    Log = Log & "[Activate " & Sh.Name & "]"
End Sub
'''
LOG = 'Public Log As String\nPublic WantCancel As Boolean\n'
PIVOT = ('ws.Range("A1:B5").Value = ws.Evaluate("{""k"",""v"";""a"",1;""b"",2;""a"",3;""b"",4}"): '
         'Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B5")).CreatePivotTable(ws.Range("D1"), "P"): '
         'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ')
#: DBL turns events on, after a case has set its sheet up with them off, and double-clicks.
DBL = 'Application.EnableEvents = True: Application.DoubleClick'
CASES: list[tuple[bool, str]] = [
    (False, 'ws.Range("A1").Select: ' + DBL),
    (False, 'ws.Range("A1").Value = 5: ws.Range("A1").Select: ' + DBL),
    (False, 'ws.Range("A1").Formula = "=1+1": ws.Range("A1").Select: ' + DBL),
    (True, 'ws.Range("A1").Value = 5: ws.Range("A1").Select: ' + DBL),
    (False, PIVOT + 'ws.Range("E2").Select: ' + DBL),
    (True, PIVOT + 'ws.Range("E2").Select: ' + DBL),
    (False, PIVOT + 'ws.Range("E2").Select: Application.DoubleClick'),
    # A row label's double-click opened Excel's Show Detail dialog, which a probe cannot answer; the grand total's
    # value in its place.
    (False, PIVOT + 'ws.Range("E4").Select: ' + DBL),
    (False, 'ws.Range("A1:B2").Select: ws.Range("B2").Activate: ' + DBL),
    (False, 'ws.Range("A1").Value = 5: ws.Range("A1").Select: ' + DBL + ': Application.DoubleClick'),
    (False, 'ws.ListObjects.Add xlSrcRange, ws.Range("A1:B3"), , xlYes: ws.Range("A2").Select: ' + DBL),
]
PROBE = '''Private Function Digest(wb As Workbook) As String
    Dim s As String, sh As Worksheet, cell As Range
    On Error Resume Next
    s = "ready " & Application.Ready & " visible " & Application.Visible & " sheets"
    For Each sh In wb.Worksheets
        s = s & " " & sh.Name
    Next
    s = s & " active " & ActiveSheet.Name & "!" & ActiveCell.Address(False, False) & " sel " & Selection.Address(False, False)
    For Each cell In wb.Worksheets(wb.Worksheets.Count).Range("A1:F8")
        If Not IsEmpty(cell.Value) Then s = s & " " & cell.Address(False, False) & "=" & cell.Text
    Next
    Digest = s
End Function

Function Probe(ByVal which As Long, ByVal cancels As Boolean) As String
    Dim sink As New ExcelSink, wb As Workbook, ws As Worksheet, pt As PivotTable, r As Variant, failure As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Activate
    Set sink.Book = wb
    Log = ""
    WantCancel = cancels
    On Error GoTo Failed
    Select Case which
{cases}
    End Select
Answered:
    On Error GoTo 0
    Application.EnableEvents = False
    Probe = Log & "|" & CStr(r) & failure & "|" & Digest(wb)
    Application.EnableEvents = True
    wb.Close False
    Exit Function
Failed:
    failure = "!E" & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def probe() -> str:
    cases = '\n'.join(f'    Case {index}\n        {statement}' for index, (_, statement) in enumerate(CASES))
    return PROBE.replace('{cases}', cases)


def install() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object',
             'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in SINK.splitlines()]
    lines += ['part.CodeModule.AddFromString text', 'Install = "ok"', 'End Function']
    return '\n'.join(lines)


def main() -> None:
    chosen = [int(one) for one in sys.argv[1:]] or list(range(len(CASES)))
    source = probe()
    rows: list[list[object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        for index in chosen:
            excel.new_document()
            excel.add_module('EventLog', LOG)
            installed = excel.run_vba(install(), 'Install', timeout=60.0, line_numbers=False)
            assert installed.ok, installed
            done = excel.run_vba(source, 'Probe', (index, CASES[index][0]), timeout=60.0, line_numbers=False)
            answer = (str(done.value) if done.ok and not done.dialogs
                      else f'STOPPED {done.outcome} {done.error} {done.message} {done.dialogs}')
            rows.append([CASES[index][1], CASES[index][0], answer])
            print(index, CASES[index][1][-60:], '=>', answer, flush=True)
    if len(chosen) == len(CASES):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'log': LOG, 'sink': SINK,
                                   'source': source, 'cases': rows}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
