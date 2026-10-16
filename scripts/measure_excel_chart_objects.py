"""Native embedded-chart collections, identity, geometry and source-data basics."""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_chart_objects.json'
SETUP = 'Dim ws As Worksheet, item As ChartObject\nSet ws = Worksheets.Add\nws.Name = "ChartData" & Worksheets.Count\n'
ADD = 'Set item = ws.ChartObjects.Add(10, 20, 300, 180)\n'
CASES = {
    'empty_collection': 'Probe = ws.ChartObjects.Count & ":" & TypeName(ws.ChartObjects)',
    'created_identity': ADD + 'Probe = item.Name & ":" & TypeName(item) & ":" & TypeName(item.Chart) & ":" & TypeName(item.Parent) & ":" & ws.ChartObjects.Count & ":" & ws.Shapes.Count',
    'created_defaults': ADD + 'Probe = item.Chart.ChartType & ":" & item.Chart.HasTitle & ":" & item.Chart.HasLegend & ":" & item.Chart.SeriesCollection.Count',
    'geometry_types': ADD + 'Probe = item.Left & ":" & item.Top & ":" & item.Width & ":" & item.Height & ":" & TypeName(item.Left) & ":" & TypeName(item.Width)',
    'geometry_mutation': ADD + 'item.Left = 12.25\nitem.Top = 31.5\nitem.Width = 400.125\nitem.Height = 210.25\nProbe = item.Left & ":" & item.Top & ":" & item.Width & ":" & item.Height',
    'negative_position': 'Set item = ws.ChartObjects.Add(-10, -20, 300, 180)\nProbe = item.Left & ":" & item.Top',
    'zero_size': 'On Error Resume Next\nSet item = ws.ChartObjects.Add(10, 20, 0, 0)\nProbe = CStr(Err.Number)\nIf Err.Number = 0 Then Probe = Probe & ":" & item.Width & ":" & item.Height',
    'negative_size': 'On Error Resume Next\nSet item = ws.ChartObjects.Add(10, 20, -30, -40)\nProbe = CStr(Err.Number)\nIf Err.Number = 0 Then Probe = Probe & ":" & item.Width & ":" & item.Height',
    'missing_index': 'On Error Resume Next\nSet item = ws.ChartObjects(1)\nProbe = Err.Number',
    'zero_index': ADD + 'On Error Resume Next\nSet item = ws.ChartObjects(0)\nProbe = Err.Number',
    'missing_name': ADD + 'On Error Resume Next\nSet item = ws.ChartObjects("Absent")\nProbe = Err.Number',
    'renamed_lookup': ADD + 'item.Name = "Revenue"\nProbe = ws.ChartObjects("revenue").Name & ":" & ws.Shapes("Revenue").Type & ":" & item.Chart.Parent.Name',
    'delete': ADD + 'item.Delete\nProbe = ws.ChartObjects.Count & ":" & ws.Shapes.Count',
    'name_after_delete': ADD + 'item.Delete\nSet item = ws.ChartObjects.Add(10, 20, 300, 180)\nProbe = item.Name & ":" & item.Index',
    'shape_chart_shared_identity': 'ws.Shapes.AddShape msoShapeRectangle, 0, 0, 20, 20\n' + ADD + 'Probe = item.Name & ":" & item.Index & ":" & ws.Shapes(2).Name & ":" & item.ShapeRange.Type',
    'chart_type_change': ADD + 'item.Chart.ChartType = xlLine\nProbe = item.Chart.ChartType',
    'title_creation': ADD + 'item.Chart.HasTitle = True\nProbe = item.Chart.ChartTitle.Text\nitem.Chart.ChartTitle.Text = "Revenue"\nProbe = Probe & ":" & item.Chart.ChartTitle.Text',
    'new_series': ADD + 'Dim series As Series\nSet series = item.Chart.SeriesCollection.NewSeries\nProbe = series.Name & ":" & series.Formula & ":" & item.Chart.SeriesCollection.Count',
    'literal_series': ADD + 'Dim series As Series\nSet series = item.Chart.SeriesCollection.NewSeries\nseries.Name = "Revenue"\nseries.Values = Array(2, 4)\nseries.XValues = Array("a", "b")\nProbe = series.Formula',
    'source_data': 'ws.Range("A1").Value = "Category"\nws.Range("B1").Value = "Revenue"\nws.Range("A2").Value = "a"\nws.Range("A3").Value = "b"\nws.Range("B2").Value = 2\nws.Range("B3").Value = 4\n' + ADD + 'item.Chart.SetSourceData Source:=ws.Range("A1:B3"), PlotBy:=xlColumns\nProbe = item.Chart.SeriesCollection.Count & ":" & item.Chart.SeriesCollection(1).Name & ":" & Replace(item.Chart.SeriesCollection(1).Formula, ws.Name, "Data")',
    'identity': ADD + 'Dim other As ChartObject\nSet other = ws.ChartObjects(1)\nProbe = (item Is other) & ":" & (item.Chart.Parent Is item)',
    'type_with_series': ADD + 'item.Chart.SeriesCollection.NewSeries.Values = Array(1, 2)\nitem.Chart.ChartType = xlLine\nProbe = item.Chart.ChartType',
    'negative_position_window': 'Set item = ws.ChartObjects.Add(-10, -20, 300, 180)\nProbe = item.Left & ":" & item.Top & ":" & ActiveWindow.UsableWidth & ":" & ActiveWindow.UsableHeight',
    'two_empty_series': ADD + 'Dim first As Series, second As Series\nSet first = item.Chart.SeriesCollection.NewSeries\nSet second = item.Chart.SeriesCollection.NewSeries\nProbe = first.Name & ":" & second.Name & ":" & first.Formula & ":" & second.Formula',
    'fractional_creation': 'Set item = ws.ChartObjects.Add(12.25, 31.5, 400.125, 210.25)\nProbe = item.Left & ":" & item.Top & ":" & item.Width & ":" & item.Height',
    'series_value_arrays': ADD + 'Dim series As Series, values As Variant, categories As Variant\nSet series = item.Chart.SeriesCollection.NewSeries\nseries.Values = Array(2, 4)\nvalues = series.Values\ncategories = series.XValues\nProbe = TypeName(values) & ":" & LBound(values) & ":" & UBound(values) & ":" & TypeName(values(1)) & ":" & categories(1) & ":" & categories(2)',
    'empty_series_values': ADD + 'Dim series As Series, values As Variant\nSet series = item.Chart.SeriesCollection.NewSeries\nOn Error Resume Next\nvalues = series.Values\nProbe = Err.Number & ":" & TypeName(values)',
}

for label, grid, plot in (
    ('rows_source', [['', 'Jan', 'Feb'], ['Revenue', 2, 4], ['Costs', 3, 1]], 'xlRows'),
    ('rows_default', [['', 'Jan', 'Feb'], ['Revenue', 2, 4], ['Costs', 3, 1]], ''),
    ('numeric_source', [[1, 2], [3, 4], [5, 6]], 'xlColumns'),
    ('labels_without_headers', [['a', 2], ['b', 4]], 'xlColumns'),
    ('numeric_labels_with_headers', [['Year', 'Revenue'], [2020, 2], [2021, 4]], 'xlColumns'),
):
    body = ''
    for row, values in enumerate(grid, 1):
        for column, value in enumerate(values, 1):
            literal = '"' + value.replace('"', '""') + '"' if isinstance(value, str) else str(value)
            body += f'ws.Cells({row}, {column}).Value = {literal}\n'
    body += ADD + f'item.Chart.SetSourceData ws.Range(ws.Cells(1, 1), ws.Cells({len(grid)}, {len(grid[0])}))' + (f', {plot}' if plot else '')
    body += '\nProbe = item.Chart.PlotBy & ":" & item.Chart.SeriesCollection.Count\nDim i As Long\nFor i = 1 To item.Chart.SeriesCollection.Count\nProbe = Probe & "|" & Replace(item.Chart.SeriesCollection(i).Formula, ws.Name, "Data")\nNext'
    CASES[label] = body

def macro(name: str) -> str:
    return 'Function Probe() As String\n' + SETUP + CASES[name] + '\nEnd Function'

if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
