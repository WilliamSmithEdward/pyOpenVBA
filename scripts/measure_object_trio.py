"""Application, Creator and Parent on one object of every Excel class the model has, reached each way a macro does.

Each object is read late bound, as `Dim o As Object` reads it: what TypeName says of its Application and whether that
is Application itself, its Creator and that value's type, its Parent's type and name (not for FillFormat and LineFormat,
which have none though the model's inventory finds one in another library, nor Chart, whose Name follows the active
sheet), and the error from assigning to each of the three,
which Excel's type library holds read-only.

    python scripts/measure_object_trio.py

writes tests/fixtures/object_trio.json, which tests/test_excel_object_trio.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/object_trio.json'

#: A sheet with data, a pivot table, a table, conditional formats, charts, shapes and the rest, for REACHES to reach.
SETUP = '''
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    Set ws2 = wb.Worksheets.Add(After:=ws)
    ws.Name = "Main"
    ws2.Name = "Second"
    ws.Activate
    ws.Range("A1:B1").Value = Array("k", "v")
    ws.Range("A2:B2").Value = Array("a", 1)
    ws.Range("A3:B3").Value = Array("b", 2)
    ws.Range("A4:B4").Value = Array("a", 3)
    ws.Range("A5:B5").Value = Array("c", 4)
    ws.Range("D1").Value = 1: ws.Range("D2").Value = 2: ws.Range("D3").Value = 3: ws.Range("D4").Value = 4
    ws.Range("D5").Value = 5
    ws.Range("F1").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "9"
    ws.Range("G1").Formula = "=1/0"
    ws.Range("G2").Value = "'12"
    ws.Range("H1").AddComment "note"
    ws.Range("H2").AddCommentThreaded "thread"
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B5")).CreatePivotTable(ws2.Range("A3"), "PT")
    pt.PivotFields("k").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum
    pt.PivotFields("k").PivotFilters.Add2 xlCaptionEquals, , "a"
    pt.PivotFields("k").ClearAllFilters
    ws.Range("L1:M1").Value = Array("x", "y")
    ws.Range("L2:M2").Value = Array(1, 2)
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("L1:M2"), , xlYes)
    Set fc = ws.Range("D1:D5").FormatConditions.Add(xlCellValue, xlGreater, "=2")
    Set t10 = ws.Range("D1:D5").FormatConditions.AddTop10
    Set av = ws.Range("D1:D5").FormatConditions.AddAboveAverage
    Set cs = ws.Range("D1:D5").FormatConditions.AddColorScale(3)
    Set db = ws.Range("D1:D5").FormatConditions.AddDatabar
    Set ic = ws.Range("D1:D5").FormatConditions.AddIconSetCondition
    Set uv = ws.Range("D1:D5").FormatConditions.AddUniqueValues
    ws.Range("O1:Q1").Value = Array("c", "s1", "s2")
    ws.Range("O2:Q2").Value = Array("p", 1, 2)
    ws.Range("O3:Q3").Value = Array("q", 3, 1)
    ws.Range("O4:Q4").Value = Array("r", 2, 5)
    Set co = ws.ChartObjects.Add(10, 200, 300, 200)
    Set ch = co.Chart
    ch.ChartType = xlLineMarkers
    ch.SetSourceData ws.Range("O1:Q4")
    ch.HasTitle = True
    ch.HasLegend = True
    ch.Axes(xlValue).HasTitle = True
    ch.Axes(xlValue).HasMajorGridlines = True
    ch.ChartGroups(1).HasHiLoLines = True
    ch.ChartGroups(1).HasUpDownBars = True
    ch.ChartGroups(1).HasDropLines = True
    Set sr = ch.SeriesCollection(1)
    sr.HasDataLabels = True
    sr.Trendlines.Add
    sr.ErrorBar xlY, xlBoth, xlFixedValue, 1
    Set co2 = ws.ChartObjects.Add(320, 200, 300, 200)
    co2.Chart.ChartType = xlColumnStacked
    co2.Chart.SetSourceData ws.Range("O1:Q4")
    co2.Chart.ChartGroups(1).HasSeriesLines = True
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 10, 10, 50, 20)
    Set sh2 = ws.Shapes.AddShape(msoShapeOval, 70, 10, 50, 20)
    Set sh3 = ws.Shapes.AddShape(msoShapeOval, 130, 10, 50, 20)
    Set gr = ws.Shapes.Range(Array(sh2.Name, sh3.Name)).Group
    Set sg = ws.Range("U1").SparklineGroups.Add(xlSparkLine, "Main!D1:D5")
    ws.Scenarios.Add "Sc", ws.Range("D1"), Array(5)
    ws.Sort.SortFields.Add ws.Range("A1")
    ws2.Range("H1:H2").Value = Application.Transpose(Array("h", 1))
    ws2.Range("H1:H2").AutoFilter
    ws.Range("A12").Value = 1
    ws.HPageBreaks.Add ws.Range("A10")
    ws.VPageBreaks.Add ws.Range("C1")
    ws.Hyperlinks.Add ws.Range("V1"), "http://example.com/"
    wb.Names.Add "Nm", "=1"
    ws.Names.Add "Lc", "=2"
    ws.CustomProperties.Add "cp", "v"
    wb.CustomDocumentProperties.Add "dp", False, msoPropertyTypeString, "v"
    ws.Range("W1").Interior.Pattern = xlPatternLinearGradient
    ws.Range("W2").Interior.Pattern = xlPatternRectangularGradient
    wb.Queries.Add "Qy", "let Source = 1 in Source"
'''

#: (the class as the checklist names it, a VBA expression reaching one). A class a macro reaches several ways has one
#: entry per way, since a Parent follows the way: a chart title's Font answers the title and a cell's the range.
REACHES = [
    ('AboveAverage', 'av'), ('Application', 'Application'), ('Areas', 'ws.Range("A1:B2,D4").Areas'),
    ('Author', 'ws.Range("H2").CommentThreaded.Author'), ('AutoFilter', 'ws2.AutoFilter'),
    ('Axes', 'ch.Axes'), ('Axis', 'ch.Axes(xlValue)'), ('AxisTitle', 'ch.Axes(xlValue).AxisTitle'),
    ('Border', 'ws.Range("A1").Borders(xlEdgeTop)'), ('Border', 'wb.Styles(1).Borders(xlEdgeTop)'),
    ('Border', 'fc.Borders(xlTop)'), ('Border', 'ch.ChartArea.Border'),
    ('Border', 'ch.Axes(xlValue).MajorGridlines.Border'),
    ('Borders', 'ws.Range("A1").Borders'), ('Borders', 'wb.Styles(1).Borders'), ('Borders', 'fc.Borders'),
    ('Borders', 'Application.FindFormat.Borders'),
    ('CalculatedFields', 'pt.CalculatedFields'), ('CalculatedItems', 'pt.PivotFields("k").CalculatedItems'),
    ('CellFormat', 'Application.FindFormat'), ('CellFormat', 'Application.ReplaceFormat'),
    ('Chart', 'ch'), ('ChartArea', 'ch.ChartArea'), ('ChartFormat', 'ch.ChartArea.Format'),
    ('ChartGroup', 'ch.ChartGroups(1)'), ('ChartGroups', 'ch.ChartGroups'), ('ChartObject', 'co'),
    ('ChartObjects', 'ws.ChartObjects'), ('ChartTitle', 'ch.ChartTitle'),
    ('ColorFormat', 'sh.Fill.ForeColor'), ('ColorFormat', 'sh.Line.ForeColor'), ('ColorScale', 'cs'),
    ('ColorScaleCriteria', 'cs.ColorScaleCriteria'), ('ColorScaleCriterion', 'cs.ColorScaleCriteria(1)'),
    ('ColorStop', 'ws.Range("W1").Interior.Gradient.ColorStops(1)'),
    ('ColorStops', 'ws.Range("W1").Interior.Gradient.ColorStops'),
    ('Comment', 'ws.Range("H1").Comment'), ('CommentThreaded', 'ws.Range("H2").CommentThreaded'),
    ('Comments', 'ws.Comments'), ('CommentsThreaded', 'ws.CommentsThreaded'),
    ('CommentsThreaded', 'ws.Range("H2").CommentThreaded.Replies'),
    ('ConditionValue', 'db.MinPoint'), ('Connections', 'wb.Connections'), ('ControlFormat', 'Nothing'),
    ('CustomProperties', 'ws.CustomProperties'), ('CustomProperty', 'ws.CustomProperties(1)'),
    ('DataBarBorder', 'db.BarBorder'), ('DataLabel', 'sr.Points(1).DataLabel'), ('DataLabels', 'sr.DataLabels'),
    ('Databar', 'db'), ('DisplayFormat', 'ws.Range("A1").DisplayFormat'),
    ('DocumentProperties', 'wb.BuiltinDocumentProperties'), ('DocumentProperties', 'wb.CustomDocumentProperties'),
    ('DocumentProperty', 'wb.BuiltinDocumentProperties(1)'), ('DocumentProperty', 'wb.CustomDocumentProperties(1)'),
    ('DownBars', 'ch.ChartGroups(1).DownBars'), ('DropLines', 'ch.ChartGroups(1).DropLines'),
    ('Error', 'ws.Range("G2").Errors(xlNumberAsText)'), ('ErrorBars', 'sr.ErrorBars'),
    ('ErrorCheckingOptions', 'Application.ErrorCheckingOptions'), ('Errors', 'ws.Range("G2").Errors'),
    ('FillFormat', 'sh.Fill'), ('FillFormat', 'ch.ChartArea.Format.Fill'),
    ('Filter', 'ws2.AutoFilter.Filters(1)'), ('Filters', 'ws2.AutoFilter.Filters'),
    ('Font', 'ws.Range("A1").Font'), ('Font', 'wb.Styles(1).Font'), ('Font', 'fc.Font'),
    ('Font', 'ch.ChartTitle.Font'), ('Font', 'ch.Legend.Font'), ('Font', 'sr.DataLabels.Font'),
    ('Font', 'ch.ChartArea.Font'), ('Font', 'Application.FindFormat.Font'),
    ('Font', 'ws.Range("A1").Characters(1, 1).Font'),
    ('FormatColor', 'cs.ColorScaleCriteria(1).FormatColor'), ('FormatColor', 'db.BarColor'), ('FormatColor', 'db.AxisColor'),
    ('FormatCondition', 'fc'), ('FormatConditions', 'ws.Range("D1:D5").FormatConditions'),
    ('FullSeriesCollection', 'ch.FullSeriesCollection'), ('Gridlines', 'ch.Axes(xlValue).MajorGridlines'),
    ('GroupShapes', 'gr.GroupItems'), ('HPageBreak', 'ws.HPageBreaks(1)'), ('HPageBreaks', 'ws.HPageBreaks'),
    ('HiLoLines', 'ch.ChartGroups(1).HiLoLines'),
    ('IconCriteria', 'ic.IconCriteria'), ('IconCriterion', 'ic.IconCriteria(1)'),
    ('IconSet', 'wb.IconSets(1)'), ('IconSetCondition', 'ic'), ('IconSets', 'wb.IconSets'),
    ('Interior', 'ws.Range("A1").Interior'), ('Interior', 'wb.Styles(1).Interior'), ('Interior', 'fc.Interior'),
    ('Interior', 'ch.ChartArea.Interior'), ('Interior', 'Application.FindFormat.Interior'),
    ('Legend', 'ch.Legend'), ('LegendEntries', 'ch.Legend.LegendEntries'), ('LegendEntry', 'ch.Legend.LegendEntries(1)'),
    ('LegendKey', 'ch.Legend.LegendEntries(1).LegendKey'), ('LineFormat', 'sh.Line'),
    ('LinearGradient', 'ws.Range("W1").Interior.Gradient'), ('ListColumn', 'lo.ListColumns(1)'),
    ('ListColumns', 'lo.ListColumns'), ('ListObject', 'lo'), ('ListObjects', 'ws.ListObjects'),
    ('ListRow', 'lo.ListRows(1)'), ('ListRows', 'lo.ListRows'), ('Name', 'wb.Names("Nm")'), ('Name', 'ws.Names("Lc")'),
    ('Names', 'wb.Names'), ('Names', 'ws.Names'), ('Names', 'Application.Names'),
    ('NegativeBarFormat', 'db.NegativeBarFormat'),
    ('Outline', 'ws.Outline'), ('Page', 'ws.PageSetup.Pages(1)'), ('PageSetup', 'ws.PageSetup'),
    ('Panes', 'ActiveWindow.Panes'), ('PivotCache', 'pt.PivotCache'), ('PivotCaches', 'wb.PivotCaches'),
    ('PivotCell', 'pt.TableRange1.Cells(2, 1).PivotCell'), ('PivotField', 'pt.PivotFields("k")'),
    ('PivotFields', 'pt.PivotFields'), ('PivotFilters', 'pt.PivotFields("k").PivotFilters'),
    ('PivotFilters', 'pt.ActiveFilters'),
    ('PivotItem', 'pt.PivotFields("k").PivotItems(1)'), ('PivotItemList', 'pt.TableRange1.Cells(2, 2).PivotCell.RowItems'),
    ('PivotItems', 'pt.PivotFields("k").PivotItems'), ('PivotTable', 'pt'), ('PivotTables', 'ws2.PivotTables'),
    ('PlotArea', 'ch.PlotArea'), ('Point', 'sr.Points(1)'), ('Points', 'sr.Points'), ('Protection', 'ws.Protection'),
    ('Queries', 'wb.Queries'), ('Range', 'ws.Range("A1")'), ('Range', 'ws2.Cells'), ('Range', 'ws.Rows(1)'),
    ('Range', 'ws.Range("A1").EntireColumn'), ('Range', 'lo.Range'), ('Range', 'pt.TableRange1'),
    ('RectangularGradient', 'ws.Range("W2").Interior.Gradient'), ('Scenario', 'ws.Scenarios(1)'),
    ('Scenarios', 'ws.Scenarios'), ('Series', 'sr'), ('SeriesCollection', 'ch.SeriesCollection'),
    ('SeriesLines', 'co2.Chart.ChartGroups(1).SeriesLines'), ('Shape', 'sh'), ('Shape', 'gr'),
    ('Shape', 'gr.GroupItems(1)'), ('Shape', 'ch.Parent.ShapeRange(1)'), ('ShapeRange', 'ws.Shapes.Range(Array(1))'),
    ('Shapes', 'ws.Shapes'), ('Shapes', 'ch.Shapes'), ('ChartObjects', 'ch.ChartObjects'), ('Hyperlinks', 'ch.Hyperlinks'), ('SheetViews', 'ActiveWindow.SheetViews'), ('Sheets', 'wb.Sheets'),
    ('Sheets', 'wb.Worksheets'), ('Sheets', 'wb.Charts'), ('Slicers', 'pt.Slicers'), ('Slicers', 'lo.Slicers'),
    ('Sort', 'ws.Sort'), ('Sort', 'lo.Sort'), ('SortField', 'ws.Sort.SortFields(1)'), ('SortFields', 'ws.Sort.SortFields'),
    ('SparkAxes', 'sg.Axes'), ('SparkColor', 'sg.Points.Markers'), ('SparkColor', 'sg.Axes.Horizontal.Axis'), ('SparkHorizontalAxis', 'sg.Axes.Horizontal'),
    ('SparkPoints', 'sg.Points'), ('SparkVerticalAxis', 'sg.Axes.Vertical'), ('Sparkline', 'sg.Item(1)'),
    ('SparklineGroup', 'sg'), ('SparklineGroups', 'ws.Range("U1").SparklineGroups'), ('Style', 'wb.Styles(1)'),
    ('Style', 'ws.Range("A1").Style'), ('Styles', 'wb.Styles'), ('Tab', 'ws.Tab'), ('TableStyle', 'wb.TableStyles(1)'),
    ('TableStyles', 'wb.TableStyles'), ('TextFrame', 'sh.TextFrame'), ('TextFrame2', 'sh.TextFrame2'),
    ('TickLabels', 'ch.Axes(xlValue).TickLabels'), ('Top10', 't10'), ('Trendline', 'sr.Trendlines(1)'),
    ('Trendlines', 'sr.Trendlines'), ('UniqueValues', 'uv'), ('UpBars', 'ch.ChartGroups(1).UpBars'),
    ('VPageBreak', 'ws.VPageBreaks(1)'), ('VPageBreaks', 'ws.VPageBreaks'), ('Validation', 'ws.Range("F1").Validation'),
    ('Window', 'ActiveWindow'), ('Windows', 'Application.Windows'), ('Windows', 'wb.Windows'), ('Workbook', 'wb'),
    ('WorkbookQuery', 'wb.Queries(1)'), ('Workbooks', 'Application.Workbooks'),
    ('Worksheet', 'ws'), ('WorksheetFunction', 'Application.WorksheetFunction'),
    ('WorksheetView', 'ActiveWindow.SheetViews(1)'),
]
REACHES = [one for one in REACHES if one[1] != 'Nothing']

_TRIO = '''
Private Function Trio(o As Object) As String
    Dim a As String, c As String, p As String, n As String, la As String, lc As String, sp As String
    Dim v As Variant, q As Object
    On Error Resume Next
    Set q = o.Application
    If Err.Number <> 0 Then
        a = "e" & Err.Number
        Err.Clear
    Else
        a = TypeName(q) & " " & (q Is Application)
    End If
    v = o.Creator
    If Err.Number <> 0 Then
        c = "e" & Err.Number
        Err.Clear
    Else
        c = v & " " & TypeName(v)
    End If
    Set q = Nothing
    Set q = o.Parent
    If Err.Number <> 0 Then
        p = "e" & Err.Number
        Err.Clear
    Else
        p = TypeName(q)
        If InStr("FillFormat LineFormat Chart", TypeName(q)) = 0 Then n = q.Name
        If Err.Number <> 0 Then
            n = "e" & Err.Number
            Err.Clear
        End If
    End If
    o.Application = 1
    la = "e" & Err.Number
    Err.Clear
    o.Creator = 1
    lc = "e" & Err.Number
    Err.Clear
    Set o.Parent = Application
    sp = "e" & Err.Number
    Err.Clear
    Trio = a & "~" & c & "~" & p & "~" & n & "~" & la & "~" & lc & "~" & sp
End Function
'''


def source() -> str:
    """The probe: Setup builds the book; Probe reads each object, and answers one line per entry of REACHES."""
    cases = ''.join(f'    Case {index}: Set Reached = {reach}\n' for index, (_, reach) in enumerate(REACHES))
    return (
        'Private wb As Object, ws As Object, ws2 As Object, pt As Object, lo As Object, fc As Object, t10 As Object\n'
        'Private av As Object, cs As Object, db As Object, ic As Object, uv As Object, co As Object, ch As Object\n'
        'Private sr As Object, co2 As Object, sh As Object, sh2 As Object, sh3 As Object, gr As Object, sg As Object\n'
        + _TRIO +
        f'\nPublic Sub Setup()\n{SETUP}\nEnd Sub\n'
        '\nPublic Function Reached(ByVal i As Long) As Object\n'
        f'    Select Case i\n{cases}    End Select\nEnd Function\n'
        '\nPublic Function One(ByVal i As Long) As String\n'
        '    Dim o As Object\n'
        '    On Error Resume Next\n'
        '    Set o = Reached(i)\n'
        '    If Err.Number <> 0 Then\n'
        '        One = "e" & Err.Number\n'
        '    Else\n'
        '        One = Replace(Trio(o), wb.Name, "<book>")\n'
        '    End If\n'
        'End Function\n'
        '\nPublic Function Probe() As String\n'
        '    Dim i As Long, out As String\n'
        '    Setup\n'
        f'    For i = 0 To {len(REACHES) - 1}\n'
        '        out = out & "@@" & One(i)\n'
        '    Next i\n'
        '    Probe = Mid(out, 3)\n'
        'End Function\n')


def main() -> None:
    probe = source()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(probe, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': probe,
              'answers': [[name, reach, answer] for (name, reach), answer in zip(REACHES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, reach, answer in record['answers']:
        print(name, reach, '->', answer)


if __name__ == '__main__':
    main()
