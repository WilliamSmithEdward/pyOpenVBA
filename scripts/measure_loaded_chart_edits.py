"""What Excel's save does to a loaded chart's part for one edit at a time: whether it adds the edit to the original and leaves the rest,
or rewrites more. On three books: a plain chart Excel saved, a dressed one Excel saved (a chart style, manual layouts, data labels, a
series' own colour and a trendline), and one openpyxl wrote, which Excel did not. Needs Excel and openpyxl:

    python scripts/measure_loaded_chart_edits.py

writes tests/fixtures/loaded_chart_edits.json and the books beside it.
"""

from __future__ import annotations

import datetime
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from pyvbaharness import ExcelSession, HarnessConfig

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
ABSOLUTE = re.compile(r'<mc:AlternateContent[^>]*><mc:Choice Requires="x15"><x15ac:absPath.*?</mc:AlternateContent>', re.S)
#: Each: what it is, and the statements, with c the chart and co its object. A read follows in r, for its answer and type.
EDITS = (
    ('nothing', ''),
    ('HasTitle True', 'c.HasTitle = True: r = c.HasTitle'),
    ('HasTitle False', 'c.HasTitle = False: r = c.HasTitle'),
    ('HasTitle True then False', 'c.HasTitle = True: c.HasTitle = False: r = c.HasTitle'),
    ('title text', 'c.HasTitle = True: c.ChartTitle.Text = "T": r = c.ChartTitle.Text'),
    ('title delete', 'c.HasTitle = True: c.ChartTitle.Delete: r = c.HasTitle'),
    ('HasLegend False', 'c.HasLegend = False: r = c.HasLegend'),
    ('HasLegend True', 'c.HasLegend = True: r = c.HasLegend'),
    ('legend position', 'c.HasLegend = True: c.Legend.Position = xlLegendPositionTop: r = c.Legend.Position'),
    ('axis title', 'c.Axes(xlValue).HasTitle = True: c.Axes(xlValue).AxisTitle.Text = "V": r = c.Axes(xlValue).AxisTitle.Text'),
    ('axis maximum', 'c.Axes(xlValue).MaximumScale = 50: r = c.Axes(xlValue).MaximumScale'),
    ('axis gridlines', 'c.Axes(xlValue).HasMajorGridlines = False: r = c.Axes(xlValue).HasMajorGridlines'),
    ('series name', 'c.SeriesCollection(1).Name = "N": r = c.SeriesCollection(1).Name'),
    ('series formula', 'c.SeriesCollection(1).Formula = "=SERIES(One!$C$1,One!$A$2:$A$3,One!$C$2:$C$3,1)": r = c.SeriesCollection(1).Formula'),
    ('area font size', 'c.ChartArea.Font.Size = 14: r = c.ChartArea.Font.Size'),
    ('title font bold', 'c.HasTitle = True: c.ChartTitle.Font.Bold = True: r = c.ChartTitle.Font.Bold'),
    ('series fill', 'c.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255: r = c.SeriesCollection(1).Format.Fill.ForeColor.RGB'),
    ('series labels', 'c.SeriesCollection(1).HasDataLabels = True: r = c.SeriesCollection(1).HasDataLabels'),
    ('point label', 'c.SeriesCollection(1).Points(1).HasDataLabel = True: c.SeriesCollection(1).Points(1).DataLabel.Text = "P": r = c.SeriesCollection(1).Points(1).DataLabel.Text'),
    ('chart type', 'c.ChartType = xlLine: r = c.ChartType'),
    ('new series', 'c.SeriesCollection.NewSeries: r = c.SeriesCollection.Count'),
    ('delete series', 'c.SeriesCollection(c.SeriesCollection.Count).Delete: r = c.SeriesCollection.Count'),
    ('source data', 'c.SetSourceData co.Parent.Range("A1:B3"): r = c.SeriesCollection.Count'),
    ('activate', 'co.Activate: r = c.HasTitle'),
    ('HasTitle False then True', 'c.HasTitle = False: c.HasTitle = True: r = c.HasTitle'),
    ('HasTitle True False True', 'c.HasTitle = True: c.HasTitle = False: c.HasTitle = True: r = c.HasTitle'),
    ('title gone then read', 'c.HasTitle = True: c.HasTitle = False: r = c.HasTitle: Err.Clear: r = r & " " & TypeName(c.ChartTitle) & " " & Err.Number'),
    ('title gone then activate', 'c.HasTitle = True: c.HasTitle = False: co.Activate: r = c.HasTitle'),
    ('new title text read', 'c.HasTitle = True: r = c.ChartTitle.Text'),
    ('title delete then True', 'c.HasTitle = True: c.ChartTitle.Delete: c.HasTitle = True: r = c.HasTitle'),
    ('HasLegend False then True', 'c.HasLegend = False: c.HasLegend = True: r = c.HasLegend'),
    ('legend gone position', 'c.HasLegend = False: r = c.Legend.Position'),
    ('legend back position', 'c.HasLegend = False: c.HasLegend = True: r = c.Legend.Position'),
    ('legend position bottom', 'c.Legend.Position = xlLegendPositionBottom: r = c.Legend.Position'),
    ('legend position same', 'c.Legend.Position = c.Legend.Position: r = c.Legend.Position'),
    ('axis minimum', 'c.Axes(xlValue).MinimumScale = 1: r = c.Axes(xlValue).MinimumScale'),
    ('axis both', 'c.Axes(xlValue).MinimumScale = 1: c.Axes(xlValue).MaximumScale = 50: r = c.Axes(xlValue).MaximumScale'),
    ('axis maximum twice', 'c.Axes(xlValue).MaximumScale = 50: c.Axes(xlValue).MaximumScale = 60: r = c.Axes(xlValue).MaximumScale'),
    ('axis maximum auto', 'c.Axes(xlValue).MaximumScale = 50: c.Axes(xlValue).MaximumScaleIsAuto = True: r = c.Axes(xlValue).MaximumScaleIsAuto'),
    ('axis maximum auto read', 'c.Axes(xlValue).MaximumScale = 50: r = c.Axes(xlValue).MaximumScaleIsAuto'),
    ('gridlines back', 'c.Axes(xlValue).HasMajorGridlines = False: c.Axes(xlValue).HasMajorGridlines = True: r = c.Axes(xlValue).HasMajorGridlines'),
    ('category gridlines', 'c.Axes(xlCategory).HasMajorGridlines = True: r = c.Axes(xlCategory).HasMajorGridlines'),
    ('minor gridlines', 'c.Axes(xlValue).HasMinorGridlines = True: r = c.Axes(xlValue).HasMinorGridlines'),
    ('minor gridlines off', 'c.Axes(xlValue).HasMinorGridlines = True: c.Axes(xlValue).HasMinorGridlines = False: r = c.Axes(xlValue).HasMinorGridlines'),
    ('series name formula', 'c.SeriesCollection(1).Name = "N": r = c.SeriesCollection(1).Formula'),
    ('series name cell', 'c.SeriesCollection(1).Name = "=One!$B$1": r = c.SeriesCollection(1).Name'),
    ('legend position read', 'r = c.Legend.Position'),
    ('title text marks', 'c.HasTitle = True: c.ChartTitle.Text = "a & b <c> \'d\' " & Chr(34) & "e" & Chr(34): r = c.ChartTitle.Text'),
    ('title text two lines', 'c.HasTitle = True: c.ChartTitle.Text = "a" & vbLf & "b": r = c.ChartTitle.Text'),
    ('title text cr', 'c.HasTitle = True: c.ChartTitle.Text = "a" & vbCr & "b": r = c.ChartTitle.Text'),
    ('title text empty', 'c.HasTitle = True: c.ChartTitle.Text = "": r = c.ChartTitle.Text'),
    ('title text spaces', 'c.HasTitle = True: c.ChartTitle.Text = " a  b ": r = c.ChartTitle.Text'),
    ('title text twice', 'c.HasTitle = True: c.ChartTitle.Text = "T": c.ChartTitle.Text = "U": r = c.ChartTitle.Text'),
    ('title caption', 'c.HasTitle = True: c.ChartTitle.Caption = "T": r = c.ChartTitle.Caption'),
    ('title text then gone', 'c.HasTitle = True: c.ChartTitle.Text = "T": c.HasTitle = False: c.HasTitle = True: r = c.ChartTitle.Text'),
    ('title text read after back', 'c.HasTitle = False: c.HasTitle = True: r = c.ChartTitle.Text'),
    ('title formula', 'c.HasTitle = True: c.ChartTitle.Formula = "=One!$B$1": r = c.ChartTitle.Text'),
    ('axis title on', 'c.Axes(xlValue).HasTitle = True: r = c.Axes(xlValue).HasTitle'),
    ('axis title text read', 'c.Axes(xlValue).HasTitle = True: r = c.Axes(xlValue).AxisTitle.Text'),
    ('axis title off', 'c.Axes(xlValue).HasTitle = True: c.Axes(xlValue).HasTitle = False: r = c.Axes(xlValue).HasTitle'),
    ('axis title back', 'c.Axes(xlValue).HasTitle = True: c.Axes(xlValue).AxisTitle.Text = "V": c.Axes(xlValue).HasTitle = False: c.Axes(xlValue).HasTitle = True: r = c.Axes(xlValue).AxisTitle.Text'),
    ('category axis title', 'c.Axes(xlCategory).HasTitle = True: c.Axes(xlCategory).AxisTitle.Text = "K": r = c.Axes(xlCategory).AxisTitle.Text'),
    ('axis title with gridlines', 'c.Axes(xlValue).HasMinorGridlines = True: c.Axes(xlValue).HasTitle = True: r = c.Axes(xlValue).HasTitle'),
    ('legend moved position read', 'c.Legend.Position = xlLegendPositionTop: c.Legend.Position = xlLegendPositionLeft: r = c.Legend.Position'),
)


def source() -> str:
    lines = [r'''Public Function Make(ByVal path As String, ByVal dressed As Boolean) As String
    Dim wb As Workbook, ws As Worksheet, c As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "One"
    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")
    ws.Range("A2:C2").Value = Array("a", 10, 4)
    ws.Range("A3:C3").Value = Array("b", 20, 6)
    If dressed Then
        Set c = ws.Shapes.AddChart2(201, xlColumnClustered).Chart
        c.SetSourceData ws.Range("A1:C3")
        c.HasTitle = True
        c.ChartTitle.Text = "Mine"
        c.Legend.Left = 10
        c.PlotArea.Width = 200
        c.SeriesCollection(1).HasDataLabels = True
        c.SeriesCollection(1).Format.Fill.ForeColor.RGB = 65280
        c.SeriesCollection(2).Trendlines.Add
    Else
        Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart
        c.SetSourceData ws.Range("A1:C3")
    End If
    wb.SaveAs path, xlOpenXMLWorkbook
    wb.Close False
End Function
''', 'Private Sub Edit(ByVal which As Long, c As Object, co As Object, r As Variant)', '    Select Case which']
    lines += [f'    Case {index}: {statements}' for index, (_, statements) in enumerate(EDITS) if statements]
    lines += ['    End Select', 'End Sub', r'''
Public Function Act(ByVal path As String, ByVal out As String, ByVal which As Long) As String
    Dim wb As Workbook, co As Object, c As Object, r As Variant, s As String
    Set wb = Workbooks.Open(path)
    Set co = wb.Worksheets(1).ChartObjects(1)
    Set c = co.Chart
    On Error Resume Next
    Edit which, c, co, r
    s = Err.Number & " " & TypeName(r) & " " & r
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs out, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close False
    Act = s
End Function
''']
    return '\n'.join(lines)


def without_folder(made: Path, out: Path) -> None:
    with zipfile.ZipFile(made) as package, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as target:
        for item in package.infolist():
            data = package.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                data = ABSOLUTE.sub('', data.decode('utf-8')).encode('utf-8')
            target.writestr(item, data)


def by_openpyxl(out: Path) -> None:
    book = Workbook()
    sheet = book.active
    assert sheet is not None
    sheet.title = 'One'
    for row in (('', 'Rev', 'Cost'), ('a', 10, 4), ('b', 20, 6)):
        sheet.append(row)
    chart = BarChart()
    chart.add_data(Reference(sheet, min_col=2, max_col=3, min_row=1, max_row=3), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=3))
    sheet.add_chart(chart, 'E2')
    book.save(out)


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if name.startswith('xl/charts/')and name.endswith(('.xml', '.rels'))}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    code = source()
    books = ('loaded_chart_plain.xlsx', 'loaded_chart_dressed.xlsx', 'loaded_chart_openpyxl.xlsx')
    cases: list[list[object]] = []
    originals: dict[str, dict[str, str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        for name, dressed in ((books[0], False), (books[1], True)):
            made = folder / name
            done = excel.run_vba(code, 'Make', (str(made), dressed), timeout=120.0, line_numbers=False)
            assert done.outcome == 'passed', done.message
            without_folder(made, FIXTURES / name)
    by_openpyxl(FIXTURES / books[2])
    for name in books:
        originals[name] = chart_parts(FIXTURES / name)
        for which, (label, _) in enumerate(EDITS):
            answer, saved = '', {}
            for _turn in range(2):
                copy, out = folder / 'in.xlsx', folder / f'{name}{which}.xlsx'
                shutil.copy(FIXTURES / name, copy)
                # A new session each case: an edit here can kill Excel, as the series formula on the dressed book does.
                with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
                    done = excel.run_vba(code, 'Act', (str(copy), str(out), which), timeout=120.0, line_numbers=False)
                answer = done.value if done.outcome == 'passed' else f'{done.outcome}: {done.message}'
                if done.outcome == 'passed':
                    saved = chart_parts(out)
                    break
            cases.append([name, which, answer, saved])
            print(name, label, '=>', answer, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'edits': EDITS, 'originals': originals, 'cases': cases}
    (FIXTURES / 'loaded_chart_edits.json').write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
