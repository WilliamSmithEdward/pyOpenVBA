"""A fourteenth sweep of everyday Excel macros, for correctness in what the newest members touched: assigning to what cannot be assigned to,
a collection's Item with every kind of index, charts, shapes and connections. Each probe keeps going under On Error Resume Next and answers a
list, so one error does not hide the next.

    python scripts/measure_excel_everyday_14.py

writes tests/fixtures/excel_everyday_14.json, which tests/test_excel_everyday_14.py replays.
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_everyday_14.json'
PRE = ('Dim ws As Worksheet, wo As Object, o As Object, c As Object, r As Object, s As String, i As Long, v As Variant, q As Object\n'
       'Set ws = Worksheets.Add\nws.Name = "Sweep"\nSet wo = ws\n')
LIST = ('ws.Range("A1:C1").Value = Array("Item", "Qty", "Cost")\nws.Range("A2:C2").Value = Array("Pen", 10, 1.5)\nws.Range("A3:C3").Value = Array("Ink", 3, 12.25)\n'
        'ws.Range("A4:C4").Value = Array("Pad", 7, 4)\nws.Range("A5:C5").Value = Array("Cap", 20, 2)\n')
CHART = LIST + 'Set c = ws.ChartObjects.Add(10, 10, 300, 200)\nc.Chart.ChartType = xlColumnClustered\nc.Chart.SetSourceData ws.Range("A1:C5")\n'
SHAPE = 'Set q = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 100, 50)\nq.Name = "Box"\n'
TABLE = LIST + 'Set q = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C5"), , xlYes)\n'
NOTE = "\ns = s & T(Err.Number)\nErr.Clear"


def tries(*statements: str) -> str:
    """Each statement, with the error it left, in the answer."""
    return 'On Error Resume Next\n' + '\n'.join(f'{one}{NOTE}' for one in statements) + '\nOn Error GoTo 0'


def reads(*expressions: str) -> str:
    """Each expression's value, or its error."""
    return 'On Error Resume Next\n' + '\n'.join(f'v = Empty\nv = {one}\ns = s & Err.Number & "=" & CStr(v) & ";"\nErr.Clear' for one in expressions) + '\nOn Error GoTo 0'


#: Each index a collection is asked with; a name no collection has, and each kind of value.
INDEXES = ['99', '0', '-1', '1.5', '"nope"', '""', 'Null', 'Empty', 'True', 'Array(1)', 'Nothing', '#1/1/2020#', '"1"']


def indexed(setup: str, collection: str) -> str:
    lines = [setup, f'Set o = {collection}', 'On Error Resume Next']
    # Typed, Comments, Areas, Borders and HPageBreaks take a Long or an enum, and a literal of another type does not compile.
    third = collection.replace('ws.', 'wo.') if any(word in collection for word in ('Comments', 'Areas', 'Borders', 'HPageBreaks')) else collection
    for index in INDEXES:
        for form in (f'TypeName(o({index}))', f'TypeName(o.Item({index}))', f'TypeName({third}({index}))'):
            lines += ['v = Empty', f'v = {form}', 's = s & Err.Number & "=" & CStr(v) & ";"', 'Err.Clear']
    lines += ['On Error GoTo 0']
    return '\n'.join(lines)


COLLECTIONS = {
    'worksheets': ('', 'ActiveWorkbook.Worksheets'),
    'sheets': ('', 'ActiveWorkbook.Sheets'),
    'workbooks': ('', 'Application.Workbooks'),
    'names': ('ActiveWorkbook.Names.Add "SweepN", ws.Range("A1")', 'ActiveWorkbook.Names'),
    'sheet_names': ('ws.Names.Add "SweepL", ws.Range("A1")', 'ws.Names'),
    'list_objects': (TABLE, 'ws.ListObjects'),
    'list_rows': (TABLE, 'q.ListRows'),
    'list_columns': (TABLE, 'q.ListColumns'),
    'shapes': (SHAPE, 'ws.Shapes'),
    'chart_objects': (CHART, 'ws.ChartObjects'),
    'series': (CHART, 'c.Chart.SeriesCollection'),
    'points': (CHART, 'c.Chart.SeriesCollection(1).Points'),
    'hyperlinks': ('ws.Hyperlinks.Add ws.Range("A1"), "https://example.com"', 'ws.Hyperlinks'),
    'comments': ('ws.Range("A1").AddComment "x"', 'ws.Comments'),
    'areas': ('', 'ws.Range("A1,C3").Areas'),
    'format_conditions': ('ws.Range("A1").FormatConditions.Add xlCellValue, xlGreater, "=1"', 'ws.Range("A1").FormatConditions'),
    'borders': ('', 'ws.Range("A1").Borders'),
    'styles': ('', 'ActiveWorkbook.Styles'),
    'connections': ('', 'ActiveWorkbook.Connections'),
    'query_tables': ('', 'ws.QueryTables'),
    'pivot_tables': ('', 'ws.PivotTables'),
    'windows': ('', 'Application.Windows'),
    'horizontal_breaks': ('ws.HPageBreaks.Add ws.Range("A5")', 'ws.HPageBreaks'),
    'sort_fields': ('ws.Sort.SortFields.Add ws.Range("A1:A5")', 'ws.Sort.SortFields'),
    'custom_properties': ('', 'ws.CustomProperties'),
}
#: Worksheet.Cells, Rows and Columns are left out: assigning to one fills the whole sheet through its default member, and never comes back.
READ_ONLY = {
    'range': ('Set o = ws.Range("B2")', ['Address', 'Count', 'Row', 'Column', 'Rows', 'Columns', 'Cells', 'Parent', 'Application', 'Worksheet', 'HasFormula', 'Text', 'Areas', 'CurrentRegion',
                                       'EntireRow', 'EntireColumn', 'End', 'Left', 'Top', 'Width', 'Height', 'MergeArea', 'ListObject', 'Hyperlinks', 'Comment', 'Font', 'Interior', 'Borders']),
    'worksheet': ('Set o = ws', ['Index', 'Parent', 'Application', 'UsedRange', 'Shapes', 'ListObjects', 'ChartObjects', 'Names', 'Hyperlinks', 'Comments',
                               'CodeName', 'Next', 'Previous', 'Type', 'ProtectContents', 'AutoFilter', 'QueryTables', 'Sort', 'PageSetup', 'Outline', 'Tab']),
    'workbook': ('Set o = ActiveWorkbook', ['Name', 'FullName', 'Path', 'Sheets', 'Worksheets', 'Names', 'Styles', 'Connections', 'ActiveSheet', 'Parent', 'Application', 'FileFormat', 'ReadOnly',
                                         'Windows', 'Charts', 'CodeName', 'HasVBProject', 'ProtectStructure', 'BuiltinDocumentProperties']),
    'application': ('Set o = Application', ['Name', 'Version', 'Build', 'Workbooks', 'ActiveSheet', 'ActiveWorkbook', 'ActiveCell', 'Selection', 'Worksheets', 'WorksheetFunction', 'Parent', 'Application',
                                         'OperatingSystem', 'ThisWorkbook', 'Sheets', 'Names', 'Windows']),
    'chart_object': (CHART + 'Set o = c', ['Chart', 'Index', 'Parent', 'Application', 'TopLeftCell', 'BottomRightCell', 'ShapeRange']),
    'chart': (CHART + 'Set o = c.Chart', ['SeriesCollection', 'ChartArea', 'PlotArea', 'Legend', 'ChartTitle', 'Parent', 'Application', 'Shapes', 'ChartGroups', 'Axes']),
    'series': (CHART + 'Set o = c.Chart.SeriesCollection(1)', ['Parent', 'Application', 'Points', 'Format', 'Trendlines', 'DataLabels', 'Border', 'Interior']),
    'shape': (SHAPE + 'Set o = q', ['ID', 'Type', 'Parent', 'Application', 'Fill', 'Line', 'TextFrame', 'TextFrame2', 'TopLeftCell', 'BottomRightCell', 'ZOrderPosition', 'Connector', 'HasChart', 'AutoShapeType']),
    'list_object': (TABLE + 'Set o = q', ['DataBodyRange', 'HeaderRowRange', 'TotalsRowRange', 'ListRows', 'ListColumns', 'Parent', 'Application', 'Sort', 'AutoFilter', 'Active', 'SourceType']),
    'name': ('Set o = ActiveWorkbook.Names.Add("SweepN", ws.Range("A1"))', ['Parent', 'Application', 'Index', 'RefersToRange']),
    'hyperlink': ('Set o = ws.Hyperlinks.Add(ws.Range("A1"), "https://example.com")', ['Parent', 'Application', 'Range', 'Shape', 'Type']),
    'comment': ('Set o = ws.Range("A1").AddComment("x")', ['Parent', 'Application', 'Author', 'Shape', 'Next', 'Previous']),
    'font': ('Set o = ws.Range("A1").Font', ['Parent', 'Application']),
    'page_setup': ('Set o = ws.PageSetup', ['Parent', 'Application', 'Pages']),
}
CASES: dict[str, str] = {
    **{f'item_{name}': indexed(*pair) for name, pair in COLLECTIONS.items()},
    **{f'read_only_{name}': setup + '\n' + 'On Error Resume Next\n' + '\n'.join(
        f'o.{member} = 1\ns = s & "{member}=" & Err.Number & ";"\nErr.Clear' for member in members) + '\nOn Error GoTo 0'
       for name, (setup, members) in READ_ONLY.items()},
    **{f'read_only_set_{name}': setup + '\n' + 'On Error Resume Next\n' + '\n'.join(
        f'Set o.{member} = ws\ns = s & "{member}=" & Err.Number & ";"\nErr.Clear' for member in members) + '\nOn Error GoTo 0'
       for name, (setup, members) in READ_ONLY.items() if name in ('range', 'worksheet', 'shape', 'chart')},
    # What read-only errors look like from typed code and through With.
    'read_only_with': 'On Error Resume Next\nWith wo.Range("B2")\n.Address = "x"\ns = Err.Number & ";"\nErr.Clear\n.Count = 1\ns = s & Err.Number\nErr.Clear\nEnd With',
    'read_only_callbyname': 'On Error Resume Next\nCallByName wo.Range("B2"), "Address", VbLet, "x"\ns = Err.Number\nErr.Clear',
    'read_only_stops': 'Set o = ws.Range("B2")\no.Address = "x"\ns = "went on"',
    # Charts as macros use them.
    'chart_after_source': CHART + reads('c.Chart.ChartType', 'c.Chart.SeriesCollection.Count', 'c.Chart.SeriesCollection(1).Name', 'c.Chart.SeriesCollection(2).Name', 'c.Chart.SeriesCollection(1).Formula',
                                       'c.Chart.HasTitle', 'c.Chart.HasLegend', 'c.Chart.Legend.Position', 'c.Name', 'c.Index', 'c.Parent.Name', 'c.Chart.Parent.Name', 'c.TopLeftCell.Address', 'TypeName(c.Chart.Parent)', 'TypeName(c.Chart.ChartArea)', 'c.Chart.ChartGroups.Count', 'c.Chart.PlotBy', 'c.Chart.DisplayBlanksAs', 'c.Chart.PlotVisibleOnly', 'c.Chart.ChartStyle'),
    'chart_title': CHART + reads('c.Chart.HasTitle') + '\nc.Chart.HasTitle = True\nc.Chart.ChartTitle.Text = "Sales"\n' + reads('c.Chart.ChartTitle.Text', 'c.Chart.ChartTitle.Caption', 'c.Chart.ChartTitle.Characters.Text', 'c.Chart.ChartTitle.Font.Bold', 'TypeName(c.Chart.ChartTitle.Format)') + '\nc.Chart.HasTitle = False\n' + reads('c.Chart.HasTitle', 'c.Chart.ChartTitle.Text'),
    'chart_series_values': CHART + 'Set r = c.Chart.SeriesCollection(1)\nv = r.Values\ns = TypeName(v) & LBound(v) & UBound(v) & v(1) & v(4)\nv = r.XValues\ns = s & TypeName(v) & LBound(v) & UBound(v) & v(1) & TypeName(v(1))\ns = s & r.AxisGroup & r.ChartType & r.PlotOrder & r.Points.Count',
    'chart_series_literal': CHART + 'Set r = c.Chart.SeriesCollection.NewSeries\nr.Name = "New"\nr.Values = Array(1, 2, 3, 4)\nr.XValues = Array("a", "b", "c", "d")\n' + reads('r.Formula', 'r.Name', 'c.Chart.SeriesCollection.Count', 'r.PlotOrder') + '\nr.Delete\n' + reads('c.Chart.SeriesCollection.Count', 'r.Name'),
    'chart_series_formula_set': CHART + 'Set r = c.Chart.SeriesCollection(1)\nr.Formula = "=SERIES(""X"",Sweep!$A$2:$A$4,Sweep!$B$2:$B$4,1)"\n' + reads('r.Name', 'r.Formula', 'r.Points.Count', 'UBound(r.Values)'),
    'chart_follows_cells': CHART + 'ws.Rows(1).Insert\ns = c.Chart.SeriesCollection(1).Formula\nws.Range("B3").Value = 99\nv = c.Chart.SeriesCollection(1).Values\ns = s & v(1)',
    'chart_axes': CHART + reads('c.Chart.HasAxis(xlCategory)', 'c.Chart.HasAxis(xlValue)', 'c.Chart.HasAxis(xlValue, xlSecondary)', 'c.Chart.Axes(xlValue).HasTitle', 'c.Chart.Axes(xlValue).HasMajorGridlines', 'c.Chart.Axes(xlCategory).HasMajorGridlines', 'c.Chart.Axes(xlValue).MinimumScaleIsAuto', 'c.Chart.Axes(xlValue).ReversePlotOrder', 'c.Chart.Axes(xlValue).ScaleType', 'c.Chart.Axes(xlCategory).CategoryType', 'c.Chart.Axes(xlValue).TickLabels.NumberFormat', 'c.Chart.Axes(xlValue, xlSecondary).HasTitle', 'TypeName(c.Chart.Axes)', 'c.Chart.Axes.Count') + '\nc.Chart.Axes(xlValue).HasTitle = True\nc.Chart.Axes(xlValue).AxisTitle.Text = "Y"\nc.Chart.Axes(xlValue).MinimumScale = 2\nc.Chart.Axes(xlValue).MaximumScale = 40\n' + reads('c.Chart.Axes(xlValue).AxisTitle.Text', 'c.Chart.Axes(xlValue).MinimumScale', 'c.Chart.Axes(xlValue).MaximumScale', 'c.Chart.Axes(xlValue).MinimumScaleIsAuto'),
    'chart_labels': CHART + 'Set r = c.Chart.SeriesCollection(1)\n' + reads('r.HasDataLabels') + '\nr.HasDataLabels = True\n' + reads('r.HasDataLabels', 'r.DataLabels.Count', 'r.DataLabels.ShowValue', 'r.DataLabels.ShowSeriesName', 'r.DataLabels.Position', 'r.Points(1).HasDataLabel', 'r.Points(1).DataLabel.Text', 'r.DataLabels(2).Text') + '\nr.Points(2).DataLabel.Text = "Hi"\n' + reads('r.Points(2).DataLabel.Text', 'r.Points(1).DataLabel.Text') + '\nr.HasDataLabels = False\n' + reads('r.HasDataLabels', 'r.Points(1).HasDataLabel'),
    'chart_type_change': CHART + 'c.Chart.ChartType = xlLine\ns = c.Chart.ChartType & "|" & c.Chart.SeriesCollection(1).ChartType\nc.Chart.SeriesCollection(2).ChartType = xlColumnClustered\ns = s & "|" & c.Chart.ChartType & "|" & c.Chart.ChartGroups.Count\nc.Chart.ChartType = xlPie\ns = s & "|" & c.Chart.ChartType & "|" & c.Chart.SeriesCollection.Count & "|" & c.Chart.HasLegend',
    'chart_delete': CHART + 'c.Delete\n' + reads('ws.ChartObjects.Count', 'TypeName(c)', 'c Is Nothing'),
    'chart_sheet': CHART + 'Set o = c.Chart.Location(xlLocationAsNewSheet, "SweepChart")\n' + reads('TypeName(o)', 'o.Name', 'ActiveSheet.Name', 'ws.ChartObjects.Count', 'Sheets("SweepChart").SeriesCollection.Count', 'Charts.Count', 'TypeName(Sheets("SweepChart"))') + '\nApplication.DisplayAlerts = False\nSheets("SweepChart").Delete\nws.Activate',
    'chart_two_loop': CHART + 'Set o = ws.ChartObjects.Add(320, 10, 200, 100)\no.Name = "Second"\nFor Each r In ws.ChartObjects\ns = s & r.Name & r.Index & ","\nNext\ns = s & ws.ChartObjects("Second").Index & ws.Shapes.Count & ws.Shapes(2).Name & ws.Shapes(2).Type & ws.Shapes(2).HasChart & TypeName(ws.Shapes(2).Chart)',
    'chart_points_format': CHART + 'Set r = c.Chart.SeriesCollection(1)\nr.Format.Fill.ForeColor.RGB = 255\nr.Points(2).Format.Fill.ForeColor.RGB = 65280\n' + reads('r.Format.Fill.ForeColor.RGB', 'r.Points(1).Format.Fill.ForeColor.RGB', 'r.Points(2).Format.Fill.ForeColor.RGB', 'r.Points(3).Format.Fill.ForeColor.RGB', 'r.Interior.Color', 'r.Points(2).Interior.Color', 'r.Format.Fill.Visible', 'r.Format.Line.Visible'),
    'chart_empty': 'Set c = ws.ChartObjects.Add(10, 10, 300, 200)\n' + reads('c.Chart.SeriesCollection.Count', 'c.Chart.ChartType', 'c.Chart.HasTitle', 'c.Chart.HasLegend', 'c.Chart.HasAxis(xlValue)', 'c.Chart.SeriesCollection(1).Name', 'c.Chart.Legend.Position', 'c.Chart.ChartTitle.Text'),
    # Shapes as macros use them.
    'shape_basics': SHAPE + reads('q.Name', 'q.Type', 'q.AutoShapeType', 'q.Left', 'q.Top', 'q.Width', 'q.Height', 'q.Rotation', 'q.Visible', 'q.ZOrderPosition', 'q.LockAspectRatio', 'q.Placement', 'q.AlternativeText', 'q.Title', 'q.HasChart', 'q.HasTextFrame', 'q.OnAction', 'q.Connector', 'q.TopLeftCell.Address', 'q.BottomRightCell.Address', 'TypeName(q.Parent)', 'q.ID > 0', 'ws.Shapes.Count', 'ws.Shapes("Box").Name', 'ws.Shapes(1).Name'),
    'shape_fill_line': SHAPE + 'q.Fill.ForeColor.RGB = 255\nq.Line.ForeColor.RGB = 65280\nq.Line.Weight = 2\n' + reads('q.Fill.ForeColor.RGB', 'q.Fill.Visible', 'q.Fill.Type', 'q.Fill.Transparency', 'q.Line.ForeColor.RGB', 'q.Line.Weight', 'q.Line.Visible', 'q.Line.DashStyle', 'q.Line.Style') + '\nq.Fill.Visible = msoFalse\nq.Line.Visible = msoFalse\n' + reads('q.Fill.Visible', 'q.Line.Visible', 'q.Fill.ForeColor.RGB'),
    'shape_text': SHAPE + 'q.TextFrame2.TextRange.Text = "Hello"\n' + reads('q.TextFrame2.TextRange.Text', 'q.TextFrame.Characters.Text', 'q.TextFrame2.HasText', 'q.TextFrame2.TextRange.Length', 'q.TextFrame2.TextRange.Font.Size', 'q.TextFrame2.TextRange.Font.Bold', 'q.TextFrame.Characters.Count', 'q.TextFrame.HorizontalAlignment', 'q.TextFrame.VerticalAlignment', 'q.TextFrame2.WordWrap', 'q.TextFrame2.AutoSize') + '\nq.TextFrame2.TextRange.Font.Bold = msoTrue\nq.TextFrame2.TextRange.Font.Size = 20\nq.TextFrame.Characters(1, 2).Font.Italic = True\n' + reads('q.TextFrame2.TextRange.Font.Bold', 'q.TextFrame2.TextRange.Font.Size', 'q.TextFrame2.TextRange.Characters(1, 1).Font.Italic', 'q.TextFrame2.TextRange.Characters(3, 1).Font.Italic', 'q.TextFrame2.TextRange.Font.Italic'),
    'shape_move_delete': SHAPE + 'q.Left = 40\nq.Top = 30\nq.Width = 60\nq.Height = 20\nq.Rotation = 45\nq.Name = "Moved"\n' + reads('q.Left', 'q.Top', 'q.Width', 'q.Height', 'q.Rotation', 'ws.Shapes("Moved").Left', 'ws.Shapes("Box").Name') + '\nq.Delete\n' + reads('ws.Shapes.Count', 'TypeName(q)'),
    'shape_order_group': SHAPE + 'Set r = ws.Shapes.AddShape(msoShapeOval, 50, 10, 100, 50)\nr.Name = "Ring"\n' + reads('q.ZOrderPosition', 'r.ZOrderPosition') + '\nq.ZOrder msoBringToFront\n' + reads('q.ZOrderPosition', 'r.ZOrderPosition', 'ws.Shapes(1).Name') + '\nSet o = ws.Shapes.Range(Array("Box", "Ring")).Group\n' + reads('o.Type', 'o.GroupItems.Count', 'o.GroupItems(1).Name', 'ws.Shapes.Count', 'q.Name', 'o.Name', 'TypeName(o)', 'q.ParentGroup.Name') + '\nSet o = o.Ungroup\n' + reads('o.Count', 'ws.Shapes.Count', 'TypeName(o)'),
    'shape_duplicate': SHAPE + 'Set r = q.Duplicate\n' + reads('TypeName(r)', 'r.Count', 'r.Name', 'r.Left', 'r.Top', 'ws.Shapes.Count', 'ws.Shapes(2).Name', 'q.Left'),
    'shape_range': SHAPE + 'ws.Shapes.AddShape(msoShapeOval, 50, 10, 100, 50).Name = "Ring"\nSet o = ws.Shapes.Range(Array(1, 2))\n' + reads('o.Count', 'o(1).Name', 'o.Item(2).Name', 'TypeName(o)', 'o.Name', 'o.Left', 'o.Width', 'ws.Shapes.Range("Box").Count', 'ws.Shapes.Range(Array("Box")).Name', 'ws.Shapes.Range(Array("nope")).Count', 'ws.Shapes.Range(9).Count') + '\no.Left = 5\n' + reads('ws.Shapes(1).Left', 'ws.Shapes(2).Left'),
    'shape_cells_follow': SHAPE + 'ws.Range("C5").Value = 1\nws.Rows(1).Insert\nws.Columns(1).Insert\n' + reads('q.Left > 10', 'q.Top > 10') + '\nq.Placement = xlFreeFloating\nv = q.Top\nws.Rows(1).Insert\ns = s & (q.Top = v)',
    'shape_textbox_line': 'Set q = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 1, 80, 30)\nSet r = ws.Shapes.AddLine(1, 1, 50, 50)\n' + reads('q.Type', 'q.TextFrame2.Orientation', 'q.Name', 'r.Type', 'r.Name', 'r.Connector', 'r.Line.Weight', 'r.HorizontalFlip', 'r.VerticalFlip', 'r.Fill.Visible', 'q.Fill.Visible', 'q.Line.Visible'),
    'shape_loop_delete': SHAPE + 'ws.Shapes.AddShape(1, 1, 1, 9, 9)\nws.Shapes.AddShape(1, 1, 1, 9, 9)\nFor Each o In ws.Shapes\ns = s & o.Name & ","\nNext\nFor i = ws.Shapes.Count To 1 Step -1\nws.Shapes(i).Delete\nNext\ns = s & ws.Shapes.Count',
    'shape_blank_name_two': SHAPE + 'ws.Shapes.AddShape(msoShapeOval, 50, 10, 100, 50).Name = "Ring"\ns = ws.Shapes("").Name\nq.ZOrder msoBringToFront\ns = s & ws.Shapes("").Name & ws.Shapes(1).Name',
    'shape_loop_delete_called': SHAPE + 'ws.Shapes.AddShape 1, 1, 1, 9, 9\nCall ws.Shapes.AddShape(1, 1, 1, 9, 9)\nFor Each o In ws.Shapes\ns = s & o.Name & ","\nNext\nFor i = ws.Shapes.Count To 1 Step -1\nws.Shapes(i).Delete\nNext\ns = s & ws.Shapes.Count',
    'call_two_in_parens_own_sub': 'Mine(1, 2)',
    'call_one_in_parens': 'Mine1 (1)\nMine1(1)\ns = "ok"',
    'call_named_in_parens': 'Mine1(a:=1)\ns = "ok"',
    'chart_interior_set': CHART + 'Set r = c.Chart.SeriesCollection(1)\nr.Interior.Color = 255\nr.Points(1).Interior.Color = 65280\n' + reads('r.Format.Fill.ForeColor.RGB', 'r.Points(1).Format.Fill.ForeColor.RGB', 'r.Points(2).Format.Fill.ForeColor.RGB', 'r.Interior.Color', 'r.Points(1).Interior.Color', 'TypeName(r.Interior)'),
    # Connections and query tables with no file behind them.
    'connections_none': reads('ActiveWorkbook.Connections.Count', 'TypeName(ActiveWorkbook.Connections)', 'ws.QueryTables.Count', 'TypeName(ws.QueryTables)', 'ws.ListObjects.Count', 'ActiveWorkbook.Queries.Count', 'ws.Range("A1").QueryTable Is Nothing'),
    'connection_add_text': 'Set o = ActiveWorkbook.Connections.Add2("SweepC", "d", "TEXT;nofile.csv", "", xlCmdDefault)\n' + reads('o.Name', 'o.Description', 'o.Type', 'TypeName(o)', 'ActiveWorkbook.Connections.Count', 'ActiveWorkbook.Connections("SweepC").Name', 'o.InModel', 'o.RefreshWithRefreshAll', 'TypeName(o.TextConnection)', 'TypeName(o.OLEDBConnection)') + '\no.Delete\n' + reads('ActiveWorkbook.Connections.Count', 'o.Name'),
    'query_table_add': 'Set q = ws.QueryTables.Add("TEXT;nofile.csv", ws.Range("A1"))\n' + reads('q.Connection', 'q.Destination.Address', 'q.Name', 'q.Refreshing', 'q.TextFileStartRow', 'q.TextFileParseType', 'q.TextFileCommaDelimiter', 'q.FieldNames', 'q.RowNumbers', 'q.AdjustColumnWidth', 'q.RefreshStyle', 'q.SaveData', 'TypeName(q.ResultRange)', 'q.ResultRange.Address', 'q.QueryType', 'ws.QueryTables.Count', 'ws.QueryTables(1).Name', 'ActiveWorkbook.Connections.Count') + '\nq.Delete\n' + reads('ws.QueryTables.Count', 'ActiveWorkbook.Connections.Count', 'q.Name'),
}
HELPERS = """
Private Sub Mine(ByVal a, ByVal b)
End Sub
Private Sub Mine1(ByVal a)
End Sub
Private Function T(ByVal n As Long) As String
    T = n & ";"
End Function
"""


def macro(name: str) -> str:
    cleanup = ('\nOn Error Resume Next\nws.Unprotect\nApplication.CutCopyMode = False\nApplication.DisplayAlerts = False\nActiveWorkbook.Names("SweepN").Delete\nws.Delete\nApplication.DisplayAlerts = True\nOn Error GoTo 0')
    return HELPERS + '\nFunction Probe() As String\n' + PRE + CASES[name] + cleanup + '\nProbe = s\nEnd Function\n'


def main() -> None:
    probes: list[dict[str, object]] = []
    excel: ExcelSession | None = None
    try:
        for name in CASES:
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                excel.__enter__()
                excel.new_document()
            result = excel.run_vba(macro(name), 'Probe', timeout=60.0, line_numbers=False)
            probe: dict[str, object] = {'name': name, 'macro': macro(name)}
            if any(dialog.classification == 'compile-error' for dialog in result.dialogs):
                probe['compile'] = True
            elif result.ok and not result.dialogs:
                probe['result'] = result.value
            else:
                found = re.search(r'VBA error (-?\d+)', str(result.error))
                assert found, (name, result.message, result.dialogs)
                probe['error'] = int(found.group(1))
            if not result.ok:
                excel.__exit__(None, None, None)
                excel = None
            probes.append(probe)
            print(name, str(probe.get('result', probe.get('error', 'does not compile')))[:100], flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
