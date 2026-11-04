"""How VBA binds a call, read from live Excel.

Two lines tell an early call from a late one. `E.Nope` does not compile on a class the library marks non-extensible, and
`E.Creator = 1` does not compile on any class that has the read-only Creator. Through an Object or a Variant both compile.
The lines sit under `If False` so that nothing runs. The second half runs what an early binding may refuse, typed and late.

    python scripts/measure_vba_binding.py

writes tests/fixtures/vba_binding.json, keeps what it holds and asks only the rest, so run it again after a time limit.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_binding.json'
TOP = '''Public HeldRange As Range
Public HeldObject As Object
Public HeldVariant As Variant
Public Type Rec
    R As Range
    O As Object
End Type

Function FRange() As Range
End Function
Function FObject() As Object
End Function
Function FVariant() As Variant
End Function
Function FNone()
End Function
Property Get PRange() As Range
End Property
Property Get PObject() As Object
End Property
'''
EXPRESSIONS = [
    'ActiveSheet', 'ActiveWorkbook', 'ActiveCell', 'ActiveChart', 'ActiveWindow', 'Selection', 'ThisWorkbook', 'Application', 'Worksheets', 'Worksheets(1)', 'Sheets', 'Sheets(1)', 'Workbooks', 'Workbooks(1)', 'Charts', 'Charts(1)', 'Names', 'Names(1)',
    'Range("A1")', 'Cells', 'Cells(1, 1)', 'Rows(1)', 'Columns(1)', 'Union(Range("A1"), Range("A2"))', 'Intersect(Range("A1"), Range("A2"))', '[A1]', 'WorksheetFunction', 'Application.WorksheetFunction',
    'Application.ActiveSheet', 'Application.ActiveCell', 'Application.Selection', 'Application.Range("A1")', 'Application.Cells', 'Application.Cells(1, 1)', 'Application.Workbooks', 'Application.Workbooks(1)', 'Application.Worksheets(1)', 'Application.ThisWorkbook', 'Application.Caller', 'Application.Windows(1)', 'Application.Evaluate("A1")',
    'ws', 'ws.Range("A1")', 'ws.Cells', 'ws.Cells(1, 1)', 'ws.Rows(1)', 'ws.Columns(1)', 'ws.UsedRange', 'ws.Parent', 'ws.Application', 'ws.Next', 'ws.Shapes', 'ws.Shapes(1)', 'ws.ChartObjects', 'ws.ChartObjects(1)', 'ws.ChartObjects(1).Chart', 'ws.ListObjects', 'ws.ListObjects(1)', 'ws.ListObjects(1).ListRows', 'ws.ListObjects(1).ListRows(1)', 'ws.ListObjects(1).ListRows(1).Range', 'ws.Names', 'ws.Names(1)', 'ws.Hyperlinks', 'ws.Hyperlinks(1)', 'ws.Comments', 'ws.Comments(1)', 'ws.PivotTables', 'ws.PivotTables(1)', 'ws.Sort', 'ws.AutoFilter', 'ws.PageSetup', 'ws.QueryTables', 'ws.QueryTables(1)', 'ws.Tab',
    'wb', 'wb.Worksheets', 'wb.Worksheets(1)', 'wb.Sheets(1)', 'wb.ActiveSheet', 'wb.Names', 'wb.Names(1)', 'wb.Names(1).RefersToRange', 'wb.Connections', 'wb.Connections(1)', 'wb.Styles', 'wb.Styles(1)', 'wb.Application', 'wb.Parent', 'wb.Charts(1)', 'wb.Windows(1)',
    'r', 'r.Value', 'r.Value2', 'r.Offset(1)', 'r.Resize(1)', 'r.Cells', 'r.Cells(1, 1)', 'r(1, 1)', 'r(1)', 'r.Item(1, 1)', 'r.Rows', 'r.Rows(1)', 'r.Columns(1)', 'r.Areas', 'r.Areas(1)', 'r.EntireRow', 'r.End(xlUp)', 'r.CurrentRegion', 'r.MergeArea', 'r.Find("a")', 'r.Parent', 'r.Worksheet', 'r.Application', 'r.Font', 'r.Font.Color', 'r.Interior', 'r.Borders', 'r.Borders(1)', 'r.ListObject', 'r.Hyperlinks', 'r.Comment', 'r.Validation', 'r.FormatConditions', 'r.FormatConditions(1)', 'r.PivotTable', 'r.QueryTable', 'r.SpecialCells(xlCellTypeBlanks)', 'r.Characters', 'r.Dependents',
    'co', 'co.Chart', 'co.Chart.ChartArea', 'co.Chart.SeriesCollection', 'co.Chart.SeriesCollection(1)', 'co.Chart.SeriesCollection(1).Points(1)', 'co.Chart.Axes(1)', 'co.Parent', 'co.TopLeftCell', 'co.ShapeRange', 'ch', 'ch.SeriesCollection', 'ch.SeriesCollection(1)', 'ch.SeriesCollection(1).Format', 'ch.Axes', 'ch.Axes(1)', 'ch.ChartTitle', 'ch.Legend', 'ch.Parent', 'sh', 'sh.Fill', 'sh.Fill.ForeColor', 'sh.TextFrame2', 'sh.TextFrame2.TextRange', 'sh.Parent', 'sh.TopLeftCell', 'lo', 'lo.Range', 'lo.ListRows', 'lo.ListColumns(1)', 'lo.DataBodyRange', 'lo.Parent',
    'o', 'o.Range("A1")', 'v', 'v.Range("A1")', 'FRange', 'FRange()', 'FObject', 'FObject()', 'FVariant()', 'FNone()', 'PRange', 'PObject', 'HeldRange', 'HeldObject', 'HeldVariant', 'arr(1)', 'arrv(1)', 'rec.R', 'rec.O', 'coll', 'coll(1)', 'coll.Item(1)', 'cobj', 'CreateObject("Scripting.Dictionary")', 'Undeclared',
]
BODY = '''Function Probe(Optional pr As Range, Optional po As Object, Optional pv As Variant) As String
    Dim ws As Worksheet, wb As Workbook, r As Range, co As ChartObject, ch As Chart, sh As Shape, lo As ListObject, o As Object, v As Variant
    Dim arr(1) As Range, arrv(1) As Variant, rec As Rec, coll As Collection, cobj As Object
    If False Then
        @@LINE@@
    End If
    Probe = "compiled"
End Function
'''
#: Binding that comes from a statement around the call rather than from an expression. @ is where the probe goes.
FORMS = {
    'with_range': 'With r\n.@\nEnd With', 'with_object': 'With o\n.@\nEnd With', 'with_expression': 'With ws.Range("A1")\n.@\nEnd With', 'with_active_sheet': 'With ActiveSheet\n.@\nEnd With', 'with_sheet_range': 'With ActiveSheet\n.Range("A1").@\nEnd With', 'with_nested': 'With r\nWith .Font\n.@\nEnd With\nEnd With', 'with_variant': 'With v\n.@\nEnd With',
    'for_each_range': 'Dim c As Range\nFor Each c In r\nc.@\nNext', 'for_each_variant': 'Dim k As Variant\nFor Each k In r\nk.@\nNext', 'for_each_object': 'Dim k As Object\nFor Each k In r\nk.@\nNext',
    'set_keeps_declared': 'Set o = r\no.@', 'typed_from_late': 'Set r = ActiveSheet.Range("A1")\nr.@',
    'param_range': 'pr.@', 'param_object': 'po.@', 'param_variant': 'pv.@',
}
#: What an early binding refuses that a late one does not: run typed and through Object, under On Error so that a run-time answer shows.
KINDS = {
    'read_only_count': ('@.Count = 1', 'r'), 'read_only_address': ('@.Address = "x"', 'r'), 'read_only_parent': ('@.Parent = 1', 'r'), 'read_only_font': ('@.Font = 1', 'r'), 'read_only_rows_default': ('@.Rows = 1', 'r9'), 'read_only_sheet_index': ('@.Index = 1', 'ws'), 'set_read_only': ('Set @.Font = Nothing', 'r'), 'set_non_object_member': ('Set @.Value = Nothing', 'r'), 'set_scalar_member': ('Set @.ColumnWidth = Nothing', 'r'),
    'too_many_args': ('v = @.Offset(1, 2, 3)', 'r'), 'too_few_args': ('v = @.Find()', 'r'), 'args_on_scalar': ('v = @.Row(1)', 'r'), 'args_on_property': ('v = @.ColumnWidth(1)', 'r'), 'bad_named': ('v = @.Offset(Foo:=1)', 'r'), 'named_twice': ('v = @.Offset(RowOffset:=1, RowOffset:=2)', 'r'), 'method_as_value': ('v = @.Select', 'r'), 'sub_as_value': ('v = @.Activate', 'ws'), 'sub_method_parens': ('@.Select (1)', 'r'), 'set_method_result': ('Set o = @.Select', 'r'),
    'wrong_literal_long': ('v = @.Offset("x")', 'r'), 'wrong_literal_comments': ('v = @.Comments("x")', 'ws'), 'wrong_literal_string_to_range': ('v = @.Range(1)', 'ws'), 'object_for_long': ('v = @.Offset(ws)', 'r'), 'set_wrong_type': ('Set @ = ws', 'r'), 'let_object_from_object': ('@ = ws', 'r'), 'default_let': ('@ = 1', 'r'), 'default_read': ('v = @', 'r'), 'default_read_sheet': ('v = @', 'ws'), 'default_let_sheet': ('@ = 1', 'ws'), 'hidden_default': ('v = @.[_Default]', 'r'), 'hidden_default_sheet': ('v = @.[_Default]', 'ws'), 'index_on_sheet': ('v = @(1)', 'ws'), 'index_on_range': ('v = @(1)', 'r'), 'index_on_font': ('v = @.Font(1)', 'r'), 'item_on_range': ('v = @.Item(1)', 'r'), 'is_scalar': ('v = @.Value Is Nothing', 'r'), 'not_a_member': ('v = @.Nope', 'r'), 'bang_on_range': ('v = @!A1', 'r'), 'bang_on_sheet': ('v = @!A1', 'ws'), 'typeof_scalar': ('v = TypeOf @.Value Is Range', 'r'), 'for_each_scalar': ('For Each v In @.Value\nNext', 'r'), 'with_scalar': ('With @.Row\nEnd With', 'r'),
}
KIND_BODY = '''Function Probe() As String
    Dim ws As @@W@@, r As @@R@@, r9 As @@R@@, o As Object, v As Variant, s As String
    Set ws = Worksheets.Add
    Set r = ws.Range("B2"): Set r9 = ws.Range("B2:B3")
    r.Value = 5
    On Error Resume Next
    @@LINE@@
    s = Err.Number
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = False
    Dim d As Object
    Set d = Worksheets(1)
    d.Delete
    Application.DisplayAlerts = True
    Probe = s
End Function
'''


def asked() -> dict[str, str]:
    sources: dict[str, str] = {}
    for expression in EXPRESSIONS:
        sources['nope ' + expression] = TOP + BODY.replace('@@LINE@@', f'v = {expression}.Nope')
        sources['creator ' + expression] = TOP + BODY.replace('@@LINE@@', f'{expression}.Creator = 1')
    for name, body in FORMS.items():
        sources['nope ' + name] = TOP + BODY.replace('@@LINE@@', body.replace('@', 'Nope'))
        sources['creator ' + name] = TOP + BODY.replace('@@LINE@@', body.replace('@', 'Creator = 1'))
    for name, (line, holder) in KINDS.items():
        for binding, sheet, cells in (('typed', 'Worksheet', 'Range'), ('late', 'Object', 'Object')):
            sources[f'kind {name} {binding}'] = KIND_BODY.replace('@@W@@', sheet).replace('@@R@@', cells).replace('@@LINE@@', line.replace('@', holder))
    return sources


def save(answers: dict[str, dict[str, object]]) -> None:
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8')


def main() -> None:
    answers: dict[str, dict[str, object]] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    excel: ExcelSession | None = None
    try:
        for key, source in asked().items():
            if key in answers:
                continue
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                excel.__enter__()
                excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=30.0, line_numbers=False)
            compile_error = [dialog for dialog in result.dialogs if dialog.classification == 'compile-error']
            if compile_error:
                answers[key] = {'compile': compile_error[0].message.replace('Compile error:\n\n', '')[:70]}
            elif result.ok:
                answers[key] = {'value': result.value}
            else:
                found = re.search(r'VBA error (-?\d+)', str(result.error))
                answers[key] = {'error': int(found.group(1)) if found else str(result.message)[:80]}
            if not result.ok:
                excel.__exit__(None, None, None)
                excel = None
            if len(answers) % 10 == 0:
                save(answers)
            print(key, answers[key], flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
        save(answers)


if __name__ == '__main__':
    main()
