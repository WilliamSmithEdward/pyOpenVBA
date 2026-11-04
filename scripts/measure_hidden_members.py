"""The hidden members of Excel's collections, _NewEnum and _Default, and small members of objects otherwise in the model.

For each collection: what `[_NewEnum]` and `[_Default](1)` answer on an Object, each error 438, and on a typed variable,
an object of type Unknown and the first item; and how many items For Each walks.

    python scripts/measure_hidden_members.py

writes tests/fixtures/hidden_members.json, which tests/test_excel_hidden_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/hidden_members.json'
COLLECTIONS = {
    'Areas': 'ws.Range("A1:B2,D1").Areas', 'Borders': 'ws.Range("A1").Borders', 'Comments': 'ws.Comments', 'CommentsThreaded': 'ws.CommentsThreaded',
    'Names': 'wb.Names', 'Styles': 'wb.Styles', 'Hyperlinks': 'ws.Hyperlinks', 'ListObjects': 'ws.ListObjects', 'ListColumns': 'lo.ListColumns', 'ListRows': 'lo.ListRows',
    'PivotCaches': 'wb.PivotCaches', 'PivotTables': 'ws.PivotTables', 'DocumentProperties': 'wb.BuiltinDocumentProperties', 'CustomProperties': 'ws.CustomProperties', 'FormatConditions': 'ws.Range("A1").FormatConditions',
    'HPageBreaks': 'ws.HPageBreaks', 'VPageBreaks': 'ws.VPageBreaks', 'SortFields': 'ws.Sort.SortFields', 'Filters': 'ws.AutoFilter.Filters', 'Errors': 'ws.Range("A1").Errors', 'Panes': 'ActiveWindow.Panes', 'SheetViews': 'ActiveWindow.SheetViews',
    'IconSets': 'wb.IconSets', 'IconSet': 'wb.IconSets(1)', 'IconCriteria': 'ic.IconCriteria', 'ColorScaleCriteria': 'cs.ColorScaleCriteria', 'SparklineGroups': 'ws.Range("F1").SparklineGroups', 'SparklineGroup': 'ws.Range("F1").SparklineGroups(1)',
    'TableStyles': 'wb.TableStyles', 'Axes': 'ch.Axes', 'Points': 'ch.SeriesCollection(1).Points', 'Trendlines': 'ch.SeriesCollection(1).Trendlines', 'FullSeriesCollection': 'ch.FullSeriesCollection',
}
TYPED = ['ws.HPageBreaks', 'ws.VPageBreaks', 'ws.Sort.SortFields', 'wb.IconSets(1)', 'ws.Range("F1").SparklineGroups', 'ch.SeriesCollection(1).Trendlines', 'ws.Range("A1").Errors', 'ws.Comments', 'wb.Names', 'ActiveWindow.Panes']
SINGLE = [
    'r = TypeName(ic.IconCriteria(1).Value)', 'Set x = db.MinPoint: r = TypeName(x.Application) & "/" & TypeName(x.Parent)', 'Set x = db.BarBorder: r = TypeName(x.Application) & "/" & TypeName(x.Parent)', 'Set x = db.NegativeBarFormat: r = TypeName(x.Application) & "/" & TypeName(x.Parent)',
    'Set x = ws.AutoFilter.Filters(1): r = TypeName(x.Application) & "/" & TypeName(x.Parent)', 'Set x = ws.Shapes(1).TextFrame: r = TypeName(x.Application) & "/" & TypeName(x.Parent)', 'Set x = wb.IconSets(1)(1): r = TypeName(x.Application)', 'Set x = lo.ListRows(1): r = TypeName(x.Application)', 'Set x = ws.Tab: r = TypeName(x.Application)',
    'r = wb.TableStyles("TableStyleLight1").[_Default]', 'r = wb.TableStyles("TableStyleLight1")', 'r = TypeName(wb.TableStyles("TableStyleLight1").TableStyleElements)', 'r = wb.TableStyles("TableStyleLight1").TableStyleElements.Count', 'r = wb.TableStyles("TableStyleLight1").TableStyleElements(xlWholeTable).HasFormat', 'r = wb.TableStyles("TableStyleLight1").TableStyleElements(xlHeaderRow).HasFormat', 'r = wb.TableStyles("TableStyleLight1").TableStyleElements(xlHeaderRow).Font.Bold',
    'r = wb.TableStyles("TableStyleLight1").Duplicate.Name', 'r = wb.TableStyles("TableStyleLight1").Duplicate("Mine").Name & "/" & wb.TableStyles.Count', 'r = wb.TableStyles.Add("Mine").Name & "/" & wb.TableStyles.Count & "/" & wb.TableStyles("Mine").BuiltIn & "/" & wb.TableStyles("Mine").ShowAsAvailableTableStyle & "/" & wb.TableStyles("Mine").ShowAsAvailablePivotTableStyle', 'r = wb.TableStyles.Add("TableStyleLight1").Name', 'r = wb.TableStyles.Add("").Name', 'r = wb.TableStyles.Add.Name',
    'ws.CustomProperties.Add "K", "V": r = ws.CustomProperties(1).[_Default]', 'ws.CustomProperties.Add "K", "V": r = ws.CustomProperties(1)', 'r = ws.Range("A1").Errors.[_Default](1).Value', 'r = ActiveWindow.Panes.[_Default](1).Index', 'ws.Range("A1").NumberFormat = "0.00": r = ws.Range("A1").FormatConditions.Add(xlExpression, , "=TRUE").NumberFormat',
    'r = Application.CellFormat', 'r = TypeName(Application.FindFormat)', 'Application.FindFormat.NumberFormat = "0.00": r = Application.FindFormat.NumberFormatLocal', 'Application.FindFormat.Clear', 'r = TypeName(ws.PageSetup.LeftHeader)', 'r = TypeName(ws.PageSetup.OddPage.LeftHeader.Picture)', 'r = ws.PageSetup.OddPage.LeftHeader.Picture.Filename', 'r = lo.ListColumns(1).XPath Is Nothing', 'r = TypeName(lo.ListColumns(1).XPath)', 'r = TypeName(ws.Protection.AllowEditRanges)', 'r = ws.Protection.AllowEditRanges.Count', 'r = ws.Protection.AllowEditRanges(1).Title', 'ws.Protection.AllowEditRanges.Add "T", ws.Range("A1"): r = ws.Protection.AllowEditRanges.Count & "/" & ws.Protection.AllowEditRanges(1).Title & "/" & ws.Protection.AllowEditRanges(1).Range.Address',
    'r = ActiveWindow.SheetViews(1).DisplayDataTypeIcons', 'r = ws.Range("A1").DisplayFormat.Characters.Count', 'r = TypeName(ws.Range("A1").DisplayFormat.Characters)', 'r = ws.Range("A1").Characters.PhoneticCharacters', 'r = TypeName(ch.SeriesCollection(1).Trendlines.Add.GetProperty("x"))', 'ch.SeriesCollection(1).Trendlines.Add.SetProperty "x", 1', 'r = ch.SeriesCollection(1).Trendlines.Add.Name', 'r = ch.SeriesCollection(1).Trendlines.[_Default](1).Name', 'r = lo.ExportToVisio', 'r = lo.Publish(Array("x", "y"), False)',
]
CASES = [f'Set o = {expression}: r = TypeName(o.[_NewEnum])' for expression in COLLECTIONS.values()] + [f'Set o = {expression}: r = TypeName(o.[_Default](1))' for expression in COLLECTIONS.values()] \
    + [f'Set o = {expression}: For Each x In o: i = i + 1: Next: r = i' for expression in COLLECTIONS.values()] + SINGLE

MODULE = r'''Option Explicit

Private Function Shown(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then
        Shown = CStr(v)
    ElseIf IsObject(v) Then
        Shown = "object " & TypeName(v)
    ElseIf IsArray(v) Then
        Shown = "array"
    Else
        Shown = TypeName(v) & ":" & CStr(v)
    End If
End Function

Public Function Typed() As String
    Dim wb As Workbook, ws As Worksheet, ch As Chart, s As String, t As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("D1:D3").Value = Application.Transpose(Array(1, 2, 3))
    ws.Rows(10).PageBreak = xlPageBreakManual
    ws.Columns(8).PageBreak = xlPageBreakManual
    ws.Sort.SortFields.Add ws.Range("D1")
    ws.Range("D5").AddComment "n"
    ws.Range("F1").SparklineGroups.Add xlSparkLine, "D1:D3"
    Set ch = ws.Shapes.AddChart(xlLine, 300, 0, 100, 100).Chart
    ch.SetSourceData ws.Range("D1:D3")
    ch.SeriesCollection(1).Trendlines.Add
    On Error Resume Next
@@TYPED@@
    Typed = s
    wb.Close False
End Function

Public Function Tried(ByVal which As Long) As String
    Dim wb As Object, ws As Object, lo As Object, ch As Object, ic As Object, cs As Object, db As Object, r As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("K", "V")
    ws.Range("A2:B2").Value = Array(1, 2)
    ws.Range("A3:B3").Value = Array(3, 4)
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B3"), , xlYes)
    ws.Range("D1:D3").Value = Application.Transpose(Array(1, 2, 3))
    ws.Range("D1:D3").AutoFilter
    ws.Range("D5").AddComment "note"
    ws.Hyperlinks.Add ws.Range("D6"), "http://x.test"
    ws.Rows(10).PageBreak = xlPageBreakManual
    ws.Columns(8).PageBreak = xlPageBreakManual
    ws.Sort.SortFields.Add ws.Range("D1")
    ws.Range("A1").FormatConditions.Add xlExpression, , "=TRUE"
    Set ic = ws.Range("E1").FormatConditions.AddIconSetCondition
    Set cs = ws.Range("E2").FormatConditions.AddColorScale(2)
    Set db = ws.Range("E3").FormatConditions.AddDatabar
    ws.Range("F1").SparklineGroups.Add xlSparkLine, "D1:D3"
    ws.Shapes.AddTextbox 1, 200, 200, 50, 50
    Set ch = ws.Shapes.AddChart(xlLine, 300, 0, 100, 100).Chart
    ch.SetSourceData ws.Range("D1:D3")
    ch.SeriesCollection(1).Trendlines.Add
    On Error GoTo Failed
    Run which, wb, ws, lo, ch, ic, cs, db, r
    s = "ok " & Shown(r)
Answered:
    On Error Resume Next
    Tried = s
    Application.FindFormat.Clear
    wb.Close False
    Exit Function
Failed:
    s = "Err " & Err.Number & " " & Err.Description
    Resume Answered
End Function
'''


def _source() -> str:
    lines = ['', 'Private Sub Run(ByVal which As Long, wb As Object, ws As Object, lo As Object, ch As Object, ic As Object, cs As Object, db As Object, r As Variant)',
             '    Dim o As Object, x As Variant, i As Long', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', 'End Sub', '']
    typed = []
    for expression in TYPED:
        for hidden in ('[_Default](1)', '[_NewEnum]'):
            typed.append(f'    Err.Clear: t = "": t = TypeName({expression}.{hidden}): s = s & "%%" & t & " e" & Err.Number')
    return MODULE.replace('@@TYPED@@', '\n'.join(typed)) + '\n'.join(lines)


def main() -> None:
    source = _source()
    answers: list[str] = []
    index = 0
    while index < len(CASES):
        with ExcelSession(HarnessConfig(exclusive=False)) as excel:
            excel.new_document()
            while index < len(CASES):
                result = excel.run_vba(source, 'Tried', (index,), timeout=25.0)
                index += 1
                if not result.ok or result.dialogs:
                    text = '; '.join(getattr(dialog, 'message', '') for dialog in result.dialogs)
                    answers.append(f'breaks: {result.outcome} {text}')
                    print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
                    break
                answers.append(str(result.value))
                print(index - 1, CASES[index - 1], '=>', answers[-1], flush=True)
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Typed', (), timeout=40.0)
        typed = str(result.value).split('%%')[1:]
        print(typed, flush=True)
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'collections': list(COLLECTIONS), 'typed': typed, 'typed_names': TYPED,
              'cases': [[case, answer] for case, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
