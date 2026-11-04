"""Read-only members held to the checklist's VERIFIED standard (docs/host_completeness.md): each read in the states that
can change it, on a sheet of its own, and again after a save and reopen where the file keeps the state. The error from
assigning to them is measured by scripts/measure_read_only_lets.py.

    python scripts/measure_read_audit.py

writes tests/fixtures/read_audit.json, which tests/test_excel_read_audit.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/read_audit.json'
_TABLE = ('ws.Range("B2:C2").Value = Array("x", "y"): ws.Range("B3:C3").Value = Array(1, 2): '
          'ws.Range("B4:C4").Value = Array(3, 4): ws.ListObjects.Add xlSrcRange, ws.Range("B2:C4"), , xlYes')
_TOTALS = _TABLE + ': ws.ListObjects(1).ShowTotals = True'
_EMPTY_TABLE = ('ws.Range("B2:C2").Value = Array("x", "y"): ws.ListObjects.Add xlSrcRange, ws.Range("B2:C3"), , xlYes')
_LINKS = 'ws.Hyperlinks.Add ws.Range("B2"), "http://example.com/", "s", "t", "shown": ws.Hyperlinks.Add ws.Range("B4"), "", "Sheet1!A1"'
_NOTE = 'ws.Range("B2").AddComment "one"'
_ARRAY = 'ws.Range("B2:C3").FormulaArray = "=ROW(A1:B2)"'
_SPILL = 'ws.Range("B2").Formula2 = "=SEQUENCE(2)"'
_REGION = 'ws.Range("B2:C3").Value = 1: ws.Range("E5").Value = 1'
_LIST = 'ws.Range("B2").Validation.Add xlValidateList, xlValidAlertWarning, xlBetween, "a,b"'
_WHOLE = 'ws.Range("B2").Validation.Add xlValidateWholeNumber, xlValidAlertInformation, xlGreater, "5"'
_BETWEEN = 'ws.Range("B2").Validation.Add xlValidateDecimal, xlValidAlertStop, xlNotBetween, "1", "9"'
_LOOP = 'ws.Range("B2").Formula = "=C2": ws.Range("C2").Formula = "=B2"'
#: (the class, the member, setup for this state, what to read).
READS = [
    *[('Range', 'Borders', '', read) for read in ('TypeName(ws.Range("B2").Borders)', 'ws.Range("B2:C3").Borders.Count')],
    ('Range', 'Columns', '', 'ws.Range("B2:C3").Columns.Count & " " & ws.Range("B2:C3").Columns(2).Address'),
    ('Range', 'Columns', '', 'ws.Range("B2:C3,E1").Columns.Count'),
    ('Range', 'Rows', '', 'ws.Range("B2:C3").Rows.Count & " " & ws.Range("B2:C3").Rows(2).Address'),
    ('Range', 'Rows', '', 'ws.Range("B2:C3,E1").Rows.Count'),
    ('Range', 'Comment', '', 'TypeName(ws.Range("B2").Comment)'),
    ('Range', 'Comment', _NOTE, 'TypeName(ws.Range("B2").Comment) & " " & ws.Range("B2").Comment.Text'),
    ('Range', 'Comment', _NOTE, 'TypeName(ws.Range("B2:C2").Comment)'),
    ('Range', 'CurrentArray', '', 'ws.Range("B2").CurrentArray.Address'),
    ('Range', 'CurrentArray', _ARRAY, 'ws.Range("C3").CurrentArray.Address'),
    ('Range', 'CurrentRegion', '', 'ws.Range("B2").CurrentRegion.Address'),
    ('Range', 'CurrentRegion', _REGION, 'ws.Range("C2").CurrentRegion.Address'),
    ('Range', 'CurrentRegion', _REGION, 'ws.Range("D4").CurrentRegion.Address'),
    ('Range', 'CurrentRegion', _REGION, 'ws.Range("B2:E5").CurrentRegion.Address'),
    ('Range', 'HasArray', '', 'ws.Range("B2").HasArray'),
    ('Range', 'HasArray', _ARRAY, 'ws.Range("C3").HasArray & " " & ws.Range("B2:C3").HasArray'),
    ('Range', 'HasArray', _ARRAY, 'IsNull(ws.Range("B2:D3").HasArray)'),
    ('Range', 'HasSpill', '', 'ws.Range("B2").HasSpill'),
    ('Range', 'HasSpill', _SPILL, 'ws.Range("B2").HasSpill & " " & ws.Range("B3").HasSpill'),
    ('Range', 'HasSpill', _SPILL, 'IsNull(ws.Range("B2:B4").HasSpill)'),
    ('Range', 'SpillingToRange', '', 'TypeName(ws.Range("B2").SpillingToRange)'),
    ('Range', 'SpillingToRange', _SPILL, 'ws.Range("B2").SpillingToRange.Address'),
    ('Range', 'SpillParent', '', 'ws.Range("B2").SpillParent.Address'),
    ('Range', 'SpillParent', _SPILL, 'ws.Range("B3").SpillParent.Address'),
    ('Range', 'Hyperlinks', '', 'ws.Range("B2").Hyperlinks.Count'),
    ('Range', 'Hyperlinks', _LINKS, 'ws.Range("B2:B4").Hyperlinks.Count'),
    ('Range', 'ListObject', '', 'TypeName(ws.Range("B2").ListObject)'),
    ('Range', 'ListObject', _TABLE, 'ws.Range("C3").ListObject.Range.Address'),
    ('Range', 'ListObject', _TABLE, 'TypeName(ws.Range("D3").ListObject)'),
    ('Range', 'LocationInTable', '', 'ws.Range("B2").LocationInTable'),
    ('Range', 'Validation', '', 'TypeName(ws.Range("B2").Validation)'),
    ('Comment', 'Author', _NOTE, 'TypeName(ws.Range("B2").Comment.Author)'),
    ('Hyperlink', 'Name', _LINKS, 'ws.Hyperlinks(1).Name & " " & ws.Hyperlinks(2).Name'),
    ('Hyperlink', 'Range', _LINKS, 'ws.Hyperlinks(2).Range.Address'),
    ('ListColumn', 'DataBodyRange', _TABLE, 'ws.ListObjects(1).ListColumns(2).DataBodyRange.Address'),
    ('ListColumn', 'DataBodyRange', _EMPTY_TABLE, 'ws.ListObjects(1).ListColumns(2).DataBodyRange.Address'),
    ('ListColumn', 'Index', _TABLE, 'ws.ListObjects(1).ListColumns("y").Index'),
    ('ListColumn', 'Range', _TOTALS, 'ws.ListObjects(1).ListColumns(1).Range.Address'),
    ('ListObject', 'DataBodyRange', _TABLE, 'ws.ListObjects(1).DataBodyRange.Address'),
    ('ListObject', 'DataBodyRange', _EMPTY_TABLE, 'ws.ListObjects(1).DataBodyRange.Address'),
    ('ListObject', 'HeaderRowRange', _TABLE, 'ws.ListObjects(1).HeaderRowRange.Address'),
    ('ListObject', 'ListColumns', _TABLE, 'ws.ListObjects(1).ListColumns.Count'),
    ('ListObject', 'ListRows', _TABLE, 'ws.ListObjects(1).ListRows.Count'),
    ('ListObject', 'ListRows', _EMPTY_TABLE, 'ws.ListObjects(1).ListRows.Count'),
    ('ListObject', 'Range', _TOTALS, 'ws.ListObjects(1).Range.Address'),
    ('ListObject', 'TotalsRowRange', _TABLE, 'TypeName(ws.ListObjects(1).TotalsRowRange)'),
    ('ListObject', 'TotalsRowRange', _TOTALS, 'ws.ListObjects(1).TotalsRowRange.Address'),
    ('ListRow', 'Index', _TABLE, 'ws.ListObjects(1).ListRows(2).Index'),
    ('ListRow', 'Range', _TABLE, 'ws.ListObjects(1).ListRows(2).Range.Address'),
    ('Sort', 'SortFields', '', 'ws.Sort.SortFields.Count'),
    ('Validation', 'AlertStyle', _LIST, 'ws.Range("B2").Validation.AlertStyle'),
    ('Validation', 'AlertStyle', _WHOLE, 'ws.Range("B2").Validation.AlertStyle'),
    ('Validation', 'AlertStyle', '', 'ws.Range("B2").Validation.AlertStyle'),
    ('Validation', 'Formula1', _LIST, 'ws.Range("B2").Validation.Formula1'),
    ('Validation', 'Formula1', _WHOLE, 'ws.Range("B2").Validation.Formula1'),
    ('Validation', 'Formula1', _BETWEEN, 'ws.Range("B2").Validation.Formula1'),
    ('Validation', 'Formula1', '', 'ws.Range("B2").Validation.Formula1'),
    ('Validation', 'Formula2', _LIST, 'ws.Range("B2").Validation.Formula2'),
    ('Validation', 'Formula2', _BETWEEN, 'ws.Range("B2").Validation.Formula2'),
    ('Validation', 'Formula2', '', 'ws.Range("B2").Validation.Formula2'),
    ('Validation', 'Operator', _LIST, 'ws.Range("B2").Validation.Operator'),
    ('Validation', 'Operator', _WHOLE, 'ws.Range("B2").Validation.Operator'),
    ('Validation', 'Operator', _BETWEEN, 'ws.Range("B2").Validation.Operator'),
    ('Validation', 'Operator', '', 'ws.Range("B2").Validation.Operator'),
    ('Validation', 'Type', _LIST, 'ws.Range("B2").Validation.Type'),
    ('Validation', 'Type', _WHOLE, 'ws.Range("B2").Validation.Type'),
    ('Validation', 'Type', _BETWEEN, 'ws.Range("B2").Validation.Type'),
    ('Validation', 'Type', '', 'ws.Range("B2").Validation.Type'),
    ('Window', 'ActiveCell', 'ws.Activate: ws.Range("C3:D4").Select', 'ActiveWindow.ActiveCell.Address'),
    ('Window', 'ActiveSheet', 'ws.Activate', 'TypeName(ActiveWindow.ActiveSheet) & " " & (ActiveWindow.ActiveSheet Is ws)'),
    ('Window', 'Index', 'ws.Activate', 'ActiveWindow.Index'),
    ('Window', 'Panes', 'ws.Activate', 'ActiveWindow.Panes.Count'),
    ('Window', 'RangeSelection', 'ws.Activate: ws.Range("C3:D4").Select', 'ActiveWindow.RangeSelection.Address'),
    ('Window', 'Selection', 'ws.Activate: ws.Range("C3:D4").Select', 'TypeName(ActiveWindow.Selection) & " " & ActiveWindow.Selection.Address'),
    ('Window', 'Selection', 'ws.Activate: ws.Shapes.AddShape(msoShapeRectangle, 1, 1, 9, 9).Select', 'TypeName(ActiveWindow.Selection)'),
    ('Application', 'ActiveWindow', 'ws.Activate', 'TypeName(ActiveWindow) & " " & (ActiveWindow Is Application.ActiveWindow)'),
    ('Application', 'CalculationVersion', '', 'TypeName(Application.CalculationVersion)'),
    ('Application', 'CalculationState', '', 'Application.CalculationState'),
    ('Application', 'OperatingSystem', '', 'TypeName(Application.OperatingSystem)'),
    ('Application', 'PathSeparator', '', 'Application.PathSeparator'),
    ('Application', 'ThisCell', '', 'TypeName(Application.ThisCell)'),
    ('Application', 'Value', '', 'Application.Value'),
    ('Application', 'Windows', '', 'TypeName(Application.Windows) & " " & (Application.Windows.Count = Windows.Count)'),
    ('Workbook', 'CalculationVersion', '', 'TypeName(wb.CalculationVersion)'),
    ('Workbook', 'Windows', '', 'wb.Windows.Count & " " & TypeName(wb.Windows(1))'),
    ('Worksheet', 'CircularReference', '', 'TypeName(ws.CircularReference)'),
    ('Worksheet', 'CircularReference', _LOOP, 'ws.CircularReference.Address'),
    ('Worksheet', 'Comments', _NOTE, 'ws.Comments.Count & " " & ws.Comments(1).Text'),
    ('Worksheet', 'Comments', '', 'ws.Comments.Count'),
    ('Worksheet', 'Hyperlinks', _LINKS, 'ws.Hyperlinks.Count & " " & ws.Hyperlinks(1).Address'),
    ('Worksheet', 'Hyperlinks', '', 'ws.Hyperlinks.Count'),
    ('Borders', 'Count', '', 'ws.Range("B2").Borders.Count & " " & ws.Range("B2:C3").Borders.Count'),
    ('Borders', 'Item', '', 'TypeName(ws.Range("B2").Borders(xlEdgeTop)) & " " & TypeName(ws.Range("B2").Borders.Item(xlInsideVertical))'),
]
#: The application's and window's answers do not outlast a sheet's state, and a selection is not saved in a way these read.
NOT_REOPENED = ('Window', 'Application')


def source() -> str:
    body = ('Private wb As Object\n\nPublic Sub Setup()\n    Set wb = Workbooks.Add(xlWBATWorksheet)\nEnd Sub\n'
            '\nPublic Sub Reopen(ByVal folder As String)\n    Dim path As String\n'
            '    path = folder & Application.PathSeparator & "reads.xlsx"\n'
            '    wb.SaveAs path, xlOpenXMLWorkbook\n    wb.Close False\n    Set wb = Workbooks.Open(path)\nEnd Sub\n'
            '\nPublic Sub Finish()\n    wb.Close False\nEnd Sub\n')
    for number, (_, _, setup, read) in enumerate(READS):
        body += (f'\nPrivate Function R{number}(ws As Object) As String\n    Dim v As Variant\n    On Error Resume Next\n'
                 f'    v = {read}\n    If Err.Number <> 0 Then\n        R{number} = "e" & Err.Number\n'
                 f'    ElseIf IsNull(v) Then\n        R{number} = "Null"\n    Else\n'
                 f'        R{number} = TypeName(v) & " " & CStr(v)\n    End If\nEnd Function\n'
                 f'\nPublic Function S{number}() As String\n    Dim ws As Object\n'
                 f'    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))\n'
                 + (f'    {setup}\n' if setup else '')
                 + f'    S{number} = Replace(R{number}(ws), wb.Name, "<book>")\nEnd Function\n'
                 f'\nPublic Function A{number}() As String\n    A{number} = Replace(R{number}(wb.Worksheets({number + 2})), wb.Name, "<book>")\nEnd Function\n')
    return body


def _runner(calls: list[str], name: str) -> str:
    parts = [calls[at:at + 200] for at in range(0, len(calls), 200)]
    body = ''
    for number, part in enumerate(parts):
        lines = ''.join(f'    out = out & "@@" & {call}()\n' for call in part)
        body += (f'\nPrivate Function {name}{number}() As String\n    Dim out As String\n{lines}'
                 f'    {name}{number} = Mid(out, 3)\nEnd Function\n')
    joined = ' & "@@" & '.join(f'{name}{number}()' for number in range(len(parts)))
    return body + f'\nPrivate Function {name}() As String\n    {name} = {joined}\nEnd Function\n'


def main() -> None:
    count = len(READS)
    runner = _runner([f'S{number}' for number in range(count)], 'First') + _runner([f'A{number}' for number in range(count)], 'Second')
    runner += ('\nPublic Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Setup\n'
               '    out = First()\n    Reopen folder\n    out = out & "@@" & Second()\n    Finish\n    Probe = out\nEnd Function\n')
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(source() + runner, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source(),
              'cases': [[READS[number][0], READS[number][1], READS[number][3], answers[number], answers[count + number]]
                        for number in range(count)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(count, 'reads')


if __name__ == '__main__':
    main()
