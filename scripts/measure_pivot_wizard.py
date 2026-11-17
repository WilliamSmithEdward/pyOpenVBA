"""Worksheet.PivotTableWizard, and whether PivotTableAfterValueChange fires on a pivot table over a range when
EnableDataValueEditing lets a macro write into its values. First look.

Each case makes a workbook with one sheet holding a small list in A1:C5, runs its statements with a sink class holding
the workbook WithEvents, and records the events the sink saw, r, and a digest of every sheet: its name, each pivot table's
name, range, versions, cache, grand totals, layout flags and style, and the text of the cells the table covers. Each
case is saved in a folder of TEMP, and its pivot and sheet parts kept.

Predictions, written before measuring:
- N1: the wizard names its table PivotTable1, makes a cache of its own, and gives back the PivotTable. Predict N1.
- V1: the table and cache take the default version CreatePivotTable gives here, the compact layout and
  TableStyleMedium2... Rival V2: the wizard makes a classic table, version 1 or 2, InGridDropZones True, with the old
  PivotTable1 style. Predict V2.
- D1: with no TableDestination, the wizard puts the table on a new sheet, before the active one, at A3. Rival D2: at
  the active cell. Predict D1.
- S1: with no SourceData, the source is the current region of the active cell. Predict S1.
- G1: RowGrand, ColumnGrand, HasAutoFormat and SaveData are the table's own flags, read back as given. Predict G1.
- A1: AfterValueChange fires on a range-based pivot table with EnableDataValueEditing on, once per write, its
  TargetRange the written cells. Rival A2: it fires only for OLAP tables, and the write lands in the cell without it.
  Predict A2. Without EnableDataValueEditing the write fails with 1004 (E1).

What look 1 showed:
- N1 held, and the default name counts on through the Excel session: PivotTable1 to PivotTable12 across twelve workbooks.
  CreatePivotTable takes from the same counter.
- V2 held. The wizard's table is Version 0 over a cache of version 8, with InGridDropZones on, no TableStyle2,
  LayoutRowDefault 1 and the fields' own names as headers. An empty one covers 2 by 2 cells where CreatePivotTable's
  covers 3 by 18.
- D1 half held. The new sheet goes before the active one, but the table sits at A1.
- S1 and G1 held. SaveData:=False reads back False.
- A2 and E1 held. AfterValueChange never fired. With EnableDataValueEditing a write lands in the cell, the grand total
  keeps its value, and RefreshTable puts the computed value back. Without it the write fails with 1004, and Change
  fires all the same.
- A destination over a cell that holds something is written over, with no error, when DisplayAlerts is off.
- ReadData:=False fails with 1004.
- Called with no arguments while the active cell is in a table, the wizard gives back that table unchanged.
- A second call with the same source moved the first call's table to E20 instead of making another.
Predictions for look 2, which saves each case and records the active cell, written before it ran:
- C1: the wizard leaves the active cell at the new table's first cell, which is why a second call works on that table.
  Rival C2: the active cell stays where it was. Predict C1.

What look 2 showed: C1 held. TableStyle2 of the wizard's table came back as a space, the harness's spelling of the
NUL a table in no style answers. A second call with no arguments sets asteriskTotals on the table. With
EnableDataValueEditing a write raises Change and no PivotTableAfterValueChange.
Predictions for look 3 (cases 20 to 23, the classic table with fields and no data field), written before it ran:
- L1: a row field alone fills the field's caption and its items down one column, with no grand total, beside the
  empty data area's column: E1:F3 for k. Rival L2: a grand total row as well. Predict L1.
- L2': a column field alone is the same turned: its caption and items across one row, over the data area's row.
- L3: one data field alone shows "Sum of v" over its total, E1:E2.

    python scripts/measure_pivot_wizard.py

writes tests/fixtures/pivot_wizard.json.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_wizard.json'
SINK = '''Public WithEvents Book As Workbook

Private Sub Book_SheetPivotTableAfterValueChange(ByVal Sh As Object, ByVal TargetPivotTable As PivotTable, ByVal TargetRange As Range)
    Log = Log & "[AfterValueChange " & Sh.Name & " " & TargetPivotTable.Name & " " & TargetRange.Address(False, False) & "]"
End Sub

Private Sub Book_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)
    Log = Log & "[Update " & Sh.Name & " " & Target.Name & "]"
End Sub

Private Sub Book_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "[Change " & Sh.Name & " " & Target.Address(False, False) & "]"
End Sub

Private Sub Book_NewSheet(ByVal Sh As Object)
    Log = Log & "[NewSheet " & Sh.Name & "]"
End Sub
'''
LOG = 'Public Log As String\n'
WIZARD = 'Set pt = ws.PivotTableWizard('
FIELDS = ('pt.PivotFields("k").Orientation = xlRowField: pt.PivotFields("g").Orientation = xlColumnField: '
          'pt.AddDataField pt.PivotFields("v"): ')
EDITABLE = ('Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:C5")).CreatePivotTable(ws.Range("E1"), "P"): '
            'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum: ')
CASES = [
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"))',
    'ws.Range("B2").Select: ' + WIZARD + ')',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"))',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), "Mine")',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), RowGrand:=False, ColumnGrand:=False)',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), HasAutoFormat:=False, SaveData:=False)',
    WIZARD + 'xlDatabase, "Sheet1!R1C1:R5C3", ws.Range("E1"))',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): ' + FIELDS + 'r = TypeName(pt)',
    'Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:C5")).CreatePivotTable(ws.Range("E1")): ' + FIELDS
    + 'r = TypeName(pt)',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): ' + WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E20"))',
    'ws.Range("E1").Value = "x": ' + WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"))',
    WIZARD + 'xlConsolidation, Array(ws.Range("A1:C5")), ws.Range("E1"))',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): ' + FIELDS + 'ws.Range("F3").Select: '
    + WIZARD + '): r = pt.Name',
    'wb.Worksheets.Add After:=ws: ' + WIZARD + 'xlDatabase, ws.Range("A1:C5"), wb.Worksheets(2).Range("B2"))',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), TableName:="Mine", ReadData:=False)',
    # PivotTableAfterValueChange on a pivot table over a range.
    EDITABLE + 'pt.EnableDataValueEditing = True: Log = "": ws.Range("F2").Value = 99: r = ws.Range("F2").Value',
    EDITABLE + 'Log = "": ws.Range("F2").Value = 99: r = ws.Range("F2").Value',
    EDITABLE + 'pt.EnableDataValueEditing = True: Log = "": ws.Range("F2:F3").Value = 7: r = ws.Range("F4").Value',
    EDITABLE + 'pt.EnableDataValueEditing = True: Log = "": ws.Range("F4").Value = 1: r = ws.Range("F4").Value',
    EDITABLE + 'pt.EnableDataValueEditing = True: ws.Range("F2").Value = 99: Log = "": pt.RefreshTable: '
    'r = ws.Range("F2").Value',
    # Look 3: the classic table with fields and no data field.
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.PivotFields("k").Orientation = xlRowField',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.PivotFields("g").Orientation = xlColumnField',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlColumnField',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.AddDataField pt.PivotFields("v")',
    # Held-out round A, predicted by the model before it ran.
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("H3")): pt.PivotFields("g").Orientation = xlRowField: '
    'pt.AddDataField pt.PivotFields("v")',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.PivotFields("k").Orientation = xlColumnField: '
    'pt.AddDataField pt.PivotFields("v")',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), RowGrand:=False): ' + FIELDS + 'r = pt.RowGrand',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.AddDataField pt.PivotFields("v"): '
    'pt.PivotFields("k").Orientation = xlRowField',
    'ws.Range("C3").Select: ' + WIZARD + '): r = pt.SourceData',
    # Held-out round B, predicted by the model before it ran.
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), "T1"): pt.PivotFields("g").Orientation = xlRowField: '
    'pt.PivotFields("k").Orientation = xlColumnField: pt.AddDataField pt.PivotFields("v"), , xlCount',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1"), ColumnGrand:=False): '
    'pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v")',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): pt.PivotFields("k").Orientation = xlRowField: '
    'pt.PivotFields("g").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v")',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): ' + FIELDS + 'pt.EnableDataValueEditing = True: '
    'Log = "": ws.Range("F3").Value = 50: r = ws.Range("H3").Value',
    WIZARD + 'xlDatabase, ws.Range("A1:C5"), ws.Range("E1")): ' + FIELDS + 'ws.Range("C2").Value = 9: Log = "": '
    'pt.RefreshTable: r = pt.TableRange2.Address(False, False)',
]
PROBE = '''Private Function Digest(wb As Workbook) As String
    Dim s As String, sh As Worksheet, pt As PivotTable, cell As Range
    On Error Resume Next
    s = "caches " & wb.PivotCaches.Count & " active " & ActiveSheet.Name
    For Each sh In wb.Worksheets
        s = s & " <" & sh.Name
        For Each pt In sh.PivotTables
            s = s & " {" & pt.Name & " " & pt.TableRange2.Address(False, False) & " v" & pt.Version & " c" & pt.PivotCache.Index
            s = s & " cv" & pt.PivotCache.Version & " " & pt.PivotCache.SourceData & " grand " & pt.RowGrand & pt.ColumnGrand
            s = s & " auto " & pt.HasAutoFormat & " save " & pt.SaveData & " zones " & pt.InGridDropZones
            s = s & " style " & pt.TableStyle2 & " layout " & pt.LayoutRowDefault & " " & pt.CompactLayoutRowHeader
            s = s & " drill " & pt.ShowDrillIndicators & " edit " & pt.EnableDataValueEditing
            For Each cell In pt.TableRange2
                If cell.Text <> "" Then s = s & " " & cell.Address(False, False) & "=" & cell.Text
            Next
            s = s & "}"
            Err.Clear
        Next
        s = s & ">"
    Next
    Digest = s
End Function

Function Probe(ByVal folder As String, ByVal which As Long) As String
    Dim sink As New ExcelSink, wb As Workbook, ws As Worksheet, pt As PivotTable, r As Variant, failure As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:C5").Value = ws.Evaluate("{""k"",""g"",""v"";""a"",""x"",1;""b"",""y"",2;""a"",""y"",3;""b"",""x"",4}")
    ws.Activate
    ws.Range("A1").Select
    Set sink.Book = wb
    Log = ""
    Application.EnableEvents = True
    On Error GoTo Failed
    Select Case which
{cases}
    End Select
Answered:
    On Error GoTo 0
    Application.EnableEvents = False
    If Not pt Is Nothing Then r = r & " ret " & pt.Name
    Probe = Log & "|" & CStr(r) & failure & "|" & Digest(wb) & "|at " & ActiveSheet.Name & "!" & ActiveCell.Address(False, False)
    On Error Resume Next
    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook
    wb.Close False
    Exit Function
Failed:
    failure = "!E" & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def probe() -> str:
    cases = '\n'.join(f'    Case {index}\n        {statement}' for index, statement in enumerate(CASES))
    return PROBE.replace('{cases}', cases)


def install() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object',
             'Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"', 'text = ""']
    lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in SINK.splitlines()]
    lines += ['part.CodeModule.AddFromString text', 'Install = "ok"', 'End Function']
    return '\n'.join(lines)


def parts(path: Path) -> dict[str, str]:
    """The pivot parts a saved case holds, and the sheets', with the user name a cache keeps as refreshedBy taken
    out."""
    if not path.exists():
        return {}
    keep = re.compile(r'xl/pivotTables/[^/]*\.xml|xl/pivotCache/pivotCacheDefinition\d+\.xml|xl/worksheets/sheet\d+\.xml')
    with zipfile.ZipFile(path) as package:
        return {name: re.sub(r'refreshedBy="[^"]*"', 'refreshedBy="&lt;user&gt;"', package.read(name).decode('utf-8'))
                for name in package.namelist() if keep.fullmatch(name)}


def main() -> None:
    chosen = [int(one) for one in sys.argv[1:]] or list(range(len(CASES)))
    source = probe()
    rows: list[list[object]] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        for index in chosen:
            excel.new_document()
            excel.add_module('EventLog', LOG)
            installed = excel.run_vba(install(), 'Install', timeout=60.0, line_numbers=False)
            assert installed.ok, installed
            done = excel.run_vba(source, 'Probe', (folder, index), timeout=120.0, line_numbers=False)
            answer = (str(done.value) if done.ok and not done.dialogs
                      else f'STOPPED {done.outcome} {done.error} {done.message} {done.dialogs}')
            rows.append([CASES[index], answer.replace(folder, '<folder>'), parts(Path(folder) / f'c{index}.xlsx')])
            print(index, CASES[index][-70:], '=>', answer, flush=True)
    if len(chosen) == len(CASES):
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'log': LOG, 'sink': SINK,
                                   'source': source, 'cases': rows}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
