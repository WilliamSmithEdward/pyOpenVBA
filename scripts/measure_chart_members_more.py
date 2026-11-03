"""More of what a macro reaches on a chart: a series' shadow, a legend entry's Format, categories, picture settings, a
title's characters, a chart's protection and data table: what each reads, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C, and a chart of them of the type it names. It
runs the case's statements, which set r to what they want recorded, reads READS, each as its value or "E" and an error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_members_more.py

writes tests/fixtures/chart_members_more.json, which tests/test_excel_chart_members_more.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_members_more.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'lineMarkers': 'xlLineMarkers', 'scatter': 'xlXYScatter', 'pie': 'xlPie', 'area': 'xlArea', 'bar': 'xlBarClustered', 'doughnut': 'xlDoughnut'}
G = 'ch.ChartGroups(1)'
S1 = 'ch.SeriesCollection(1)'
P2 = f'{S1}.Points(2)'
ENTRY = 'ch.Legend.LegendEntries(1)'
CASES: list[tuple[str, str]] = []

# A series' shadow, over what else it has, and what takes it away.
for kind in TYPES:
    CASES += [(kind, f'r = {S1}.Shadow'), (kind, f'{S1}.Shadow = True: r = {S1}.Shadow'), (kind, f'{S1}.Shadow = True: {S1}.Shadow = False: r = {S1}.Shadow'), (kind, f'{S1}.Shadow = False: r = {S1}.Shadow')]
CASES += [('column', f'{S1}.Format.Fill.ForeColor.RGB = 255: {S1}.Shadow = True: r = {S1}.Shadow'), ('column', f'{S1}.Shadow = True: {S1}.Format.Fill.ForeColor.RGB = 255: r = {S1}.Shadow'),
          ('column', f'{S1}.Shadow = True: {S1}.ClearFormats: r = {S1}.Shadow'), ('line', f'{S1}.Format.Line.Weight = 3: {S1}.Shadow = True: r = {S1}.Shadow'), ('column', f'{S1}.Shadow = True: r = {P2}.Shadow'),
          ('column', f'{S1}.Shadow = True: ch.ChartType = xlLine: r = {S1}.Shadow'), ('column', f'{S1}.Shadow = True: r = ch.SeriesCollection(2).Shadow'), ('column', f'{S1}.Shadow = 5: r = {S1}.Shadow'), ('column', f'{S1}.Shadow = "x": r = {S1}.Shadow'),
          ('column', f'{P2}.Shadow = True: r = {S1}.Shadow & " " & {P2}.Shadow'), ('column', f'{S1}.Shadow = True: {S1}.Format.Fill.Visible = msoFalse: r = {S1}.Shadow')]

# A legend entry's Format.
LOOK = f'{ENTRY}.Format'
FILL_READS = '{0}.Fill.Visible & " " & {0}.Fill.ForeColor.RGB & " " & {0}.Fill.Type & " " & {0}.Fill.Transparency'
LINE_READS = '{0}.Line.Visible & " " & {0}.Line.ForeColor.RGB & " " & {0}.Line.Weight & " " & {0}.Line.DashStyle & " " & {0}.Line.Transparency'
CASES += [('column', f'r = TypeName({LOOK}) & " " & TypeName({LOOK}.Parent)'), ('column', f'r = {FILL_READS.format(LOOK)}'), ('column', f'r = {LINE_READS.format(LOOK)}'), ('column', f'{LOOK}.Fill.ForeColor.RGB = 255: r = {FILL_READS.format(LOOK)}'),
          ('column', f'{LOOK}.Line.ForeColor.RGB = 255: r = {LINE_READS.format(LOOK)}'), ('column', f'{LOOK}.Fill.Visible = msoFalse: r = {FILL_READS.format(LOOK)}'), ('column', f'{LOOK}.Line.Visible = msoFalse: r = {LINE_READS.format(LOOK)}'),
          ('column', f'{LOOK}.Line.Weight = 3: r = {LINE_READS.format(LOOK)}'), ('pie', f'r = {FILL_READS.format(LOOK)}'), ('line', f'r = {LINE_READS.format(LOOK)}'),
          ('column', f'{ENTRY}.Font.Bold = True: {LOOK}.Fill.ForeColor.RGB = 255: r = {ENTRY}.Font.Bold'), ('column', f'{LOOK}.Fill.ForeColor.RGB = 255: {ENTRY}.Delete: r = ch.Legend.LegendEntries.Count'), ('column', f'r = {LOOK}.Glow.Radius'), ('column', f'r = TypeName({LOOK}.TextFrame2)')]

# Categories.
CC = f'{G}.CategoryCollection'
FC = f'{G}.FullCategoryCollection'
CASES += [('column', f'r = TypeName({CC}) & " " & TypeName({CC}.Parent) & " " & TypeName({CC}.Application)'), ('column', f'r = {CC}.Count & " " & {FC}.Count'), ('column', f'r = TypeName({CC}(1)) & " " & TypeName({CC}.Item(1)) & " " & TypeName({CC}(1).Parent)'),
          ('column', f'r = {CC}(1).Name & " " & {CC}(1).IsFiltered & " " & {CC}(3).Name'), ('column', f'r = {FC}(2).Name & " " & {FC}(2).IsFiltered'), ('column', f'r = TypeName({G}.CategoryCollection(1))'), ('column', f'r = {G}.CategoryCollection(2).Name & " " & {G}.FullCategoryCollection(3).Name'),
          ('column', f'r = TypeName({CC}(4))'), ('column', f'r = TypeName({CC}(0))'), ('column', f'r = TypeName({CC}("a"))'), ('column', f'r = TypeName({G}.CategoryCollection(4))'), ('column', 'n = 0: For Each o In ' + CC + ': n = n + 1: t = t & o.Name: Next: r = n & " " & t'),
          ('column', 'n = 0: For Each o In ' + FC + ': n = n + 1: t = t & o.Name: Next: r = n & " " & t'), ('column', f'{CC}(1).Name = "z": r = {CC}(1).Name & " " & ws.Range("A2").Value'), ('column', f'{FC}(2).IsFiltered = True: r = {CC}.Count & " " & {FC}.Count & " " & {FC}(2).IsFiltered'),
          ('column', f'{FC}(2).IsFiltered = False: r = {CC}.Count & " " & {FC}(2).IsFiltered'), ('pie', f'r = {CC}.Count & " " & {CC}(2).Name'), ('scatter', f'r = {CC}.Count & " " & {CC}(1).Name'), ('line', f'r = {CC}.Count & " " & {CC}(3).Name'),
          ('column', f'{S1}.XValues = Array("p", "q"): r = {CC}.Count & " " & {CC}(1).Name'), ('column', f'r = {CC}.Creator & " " & {CC}(1).Creator'), ('column', f'ws.Range("A2").Value = "new": r = {CC}(1).Name'), ('column', f'ws.Rows(3).Hidden = True: r = {CC}.Count & " " & {FC}.Count & " " & {CC}(2).Name')]

# Picture settings, which a series and a point keep.
for owner, start in ((S1, ''), (P2, '')):
    for name, value in (('ApplyPictToEnd', True), ('ApplyPictToFront', True), ('ApplyPictToSides', True), ('PictureType', 2), ('PictureType', 3), ('PictureType', 9), ('PictureUnit2', 2.5), ('PictureUnit2', -1)):
        CASES += [('column', f'r = {owner}.{name}'), ('column', f'{owner}.{name} = {value}: r = {owner}.{name}'), ('line', f'r = {owner}.{name}'), ('pie', f'r = {owner}.{name}')]
CASES += [('pie', f'r = {P2}.PieSliceLocation(xlHorizontalCoordinate)'), ('pie', f'r = {P2}.PieSliceLocation(xlVerticalCoordinate)'), ('column', f'r = {P2}.PieSliceLocation(xlHorizontalCoordinate)'), ('pie', f'r = {P2}.PieSliceLocation(xlHorizontalCoordinate, xlInsideEnd)'), ('pie', f'r = {P2}.PieSliceLocation(5)')]

# What a series has only on the newer kinds of chart.
for name in ('GeoMappingLevel', 'GeoProjectionType', 'ParentDataLabelOption', 'QuartileCalculationInclusiveMedian', 'RegionLabelOption', 'SeriesColorGradientStyle', 'ValueSortOrder'):
    CASES += [('column', f'r = {S1}.{name}'), ('column', f'{S1}.{name} = 1: r = {S1}.{name}'), ('line', f'r = {S1}.{name}')]
for name in ('SeriesColorMaxGradientStop', 'SeriesColorMidGradientStop', 'SeriesColorMinGradientStop'):
    CASES += [('column', f'r = TypeName({S1}.{name})')]
CASES += [('pie', f'{S1}.HasDataLabels = True: r = TypeName({S1}.LeaderLines)'), ('pie', f'{S1}.HasLeaderLines = True: r = TypeName({S1}.LeaderLines)'), ('column', f'r = TypeName({S1}.LeaderLines)'), ('column', f'{S1}.HasLeaderLines = True: r = TypeName({S1}.LeaderLines)'),
          ('pie', f'{S1}.HasDataLabels = True: {S1}.HasLeaderLines = True: r = {S1}.LeaderLines.Border.Color & " " & TypeName({S1}.LeaderLines.Format) & " " & TypeName({S1}.LeaderLines.Parent)'),
          ('pie', f'{S1}.HasDataLabels = True: {S1}.HasLeaderLines = True: {S1}.LeaderLines.Border.Color = 255: r = {S1}.LeaderLines.Border.Color'), ('pie', f'{S1}.HasDataLabels = True: {S1}.HasLeaderLines = True: {S1}.LeaderLines.Delete: r = {S1}.HasLeaderLines'),
          ('pie', f'{S1}.HasDataLabels = True: {S1}.HasLeaderLines = True: r = {S1}.LeaderLines.Select')]

# A title's characters.
for owner, before in (('ch.ChartTitle', 'ch.HasTitle = True: ch.ChartTitle.Text = "Hello there": '), ('ch.Axes(xlValue).AxisTitle', 'ch.Axes(xlValue).HasTitle = True: ch.Axes(xlValue).AxisTitle.Text = "Hello there": ')):
    chars = f'{owner}.Characters'
    CASES += [('column', f'{before}r = TypeName({chars}) & " " & TypeName({chars}.Parent) & " " & TypeName({chars}.Application)'), ('column', f'{before}r = {chars}.Text & "|" & {chars}.Count & "|" & {chars}.Caption'),
              ('column', f'{before}r = {owner}.Characters(1, 5).Text & "|" & {owner}.Characters(1, 5).Count'), ('column', f'{before}r = {owner}.Characters(7).Text'), ('column', f'{before}r = {owner}.Characters(, 3).Text'), ('column', f'{before}r = {owner}.Characters(20, 5).Text & "|" & {owner}.Characters(20, 5).Count'),
              ('column', f'{before}r = {owner}.Characters(0, 2).Text'), ('column', f'{before}r = {owner}.Characters(1, 0).Text'), ('column', f'{before}r = {owner}.Characters(1, 5).Font.Bold & " " & {owner}.Characters(1, 5).Font.Size'), ('column', f'{before}{owner}.Characters(1, 5).Font.Bold = True: r = {owner}.Characters(1, 5).Font.Bold & " " & {owner}.Characters(7, 2).Font.Bold & " " & {owner}.Font.Bold'),
              ('column', f'{before}{owner}.Characters(1, 5).Text = "Howdy!": r = {owner}.Text'), ('column', f'{before}{owner}.Characters(7).Delete: r = {owner}.Text'), ('column', f'{before}{owner}.Characters(1, 5).Insert "Oh": r = {owner}.Text'), ('column', f'{before}{owner}.Characters(1, 5).Caption = "x": r = {owner}.Text'), ('column', f'{before}{owner}.Characters.Text = "whole": r = {owner}.Text')]

# A chart's protection, on a chart on a sheet.
READ_LOCK = 'ch.ProtectContents & " " & ch.ProtectDrawingObjects & " " & ch.ProtectFormatting & " " & ch.ProtectData & " " & ch.ProtectGoalSeek & " " & ch.ProtectSelection & " " & ch.ProtectionMode'
CASES += [('column', f'r = {READ_LOCK}'), ('column', f'ch.Protect: r = {READ_LOCK}'), ('column', f'ch.Protect: ch.Unprotect: r = {READ_LOCK}'), ('column', f'ch.Unprotect: r = {READ_LOCK}'), ('column', f'ch.Protect "pw": ch.Unprotect "pw": r = {READ_LOCK}'), ('column', f'ch.Protect "pw": ch.Unprotect "no": r = {READ_LOCK}'), ('column', f'ch.Protect "pw": ch.Unprotect: r = {READ_LOCK}'),
          ('column', f'ch.Protect , True, True, True, True: r = {READ_LOCK}'), ('column', f'ch.Protect DrawingObjects:=False, Contents:=True: r = {READ_LOCK}'), ('column', f'ch.Protect: {S1}.Name = "z": r = {S1}.Name'), ('column', f'ch.Protect: ch.HasTitle = True: r = ch.HasTitle'), ('column', f'ch.Protect: ch.ChartType = xlLine: r = ch.ChartType'), ('column', f'ch.Protect: r = ch.ChartType & " " & {S1}.Name'), ('column', f'ws.Protect: ch.Protect: r = {READ_LOCK}'),
          ('column', f'ch.Protect: ws.Range("B2").Value = 9: r = {S1}.Values(1)'), ('column', f'ch.ProtectFormatting = True: r = {READ_LOCK}'), ('column', f'ch.ProtectData = True: r = {READ_LOCK}'), ('column', f'ch.ProtectGoalSeek = True: r = {READ_LOCK}'), ('column', f'ch.ProtectSelection = True: r = {READ_LOCK}')]

# A chart's page setup and data table.
TABLE = 'ch.DataTable'
CASES += [('column', 'r = TypeName(ch.PageSetup) & " " & TypeName(ch.PageSetup.Parent)'), ('column', 'r = ch.PageSetup.Orientation & " " & ch.PageSetup.ChartSize'), ('column', 'ch.PageSetup.Orientation = xlLandscape: r = ch.PageSetup.Orientation'),
          ('column', f'r = TypeName({TABLE})'), ('column', 'r = ch.HasDataTable'), ('column', f'ch.HasDataTable = True: r = TypeName({TABLE}) & " " & TypeName({TABLE}.Parent) & " " & TypeName({TABLE}.Application)'),
          ('column', f'ch.HasDataTable = True: r = {TABLE}.ShowLegendKey & " " & {TABLE}.HasBorderHorizontal & " " & {TABLE}.HasBorderVertical & " " & {TABLE}.HasBorderOutline'), ('column', f'ch.HasDataTable = True: {TABLE}.ShowLegendKey = True: r = {TABLE}.ShowLegendKey'), ('column', f'ch.HasDataTable = True: {TABLE}.HasBorderHorizontal = False: {TABLE}.HasBorderVertical = False: {TABLE}.HasBorderOutline = False: r = {TABLE}.HasBorderHorizontal'),
          ('column', f'ch.HasDataTable = True: {TABLE}.Delete: r = ch.HasDataTable'), ('column', f'ch.HasDataTable = True: {TABLE}.Select: r = TypeName(Selection) & " " & Selection.Name'), ('column', f'ch.HasDataTable = True: r = TypeName({TABLE}.Border) & " " & TypeName({TABLE}.Font) & " " & TypeName({TABLE}.Format)'),
          ('column', f'ch.HasDataTable = True: {TABLE}.Font.Bold = True: r = {TABLE}.Font.Bold'), ('column', 'ch.HasDataTable = True: ch.HasDataTable = False: r = ch.HasDataTable'), ('pie', 'ch.HasDataTable = True: r = ch.HasDataTable'), ('scatter', 'ch.HasDataTable = True: r = ch.HasDataTable'), ('line', f'ch.HasDataTable = True: r = {TABLE}.ShowLegendKey'), ('column', 'ch.HasDataTable = True: ch.ChartType = xlPie: r = ch.HasDataTable')]
READS = ['ch.SeriesCollection.Count', f'{S1}.Name']


def _probe() -> str:
    """The cases go in procedures of a hundred, since one holding them all is too large for VBA."""
    lines: list[str] = []
    for first in range(0, len(CASES), 100):
        lines += [f'Private Sub Act{first // 100}(ByVal which As Long, ws As Worksheet, ch As Object, r As Variant)', '    Dim n As Long, t As String, o As Object', '    Select Case which']
        for index, (kind, statement) in enumerate(CASES[first:first + 100], first):
            lines += [f'    Case {index}', '        ch.SetSourceData ws.Range("A1:C4"), xlColumns', f'        ch.ChartType = {TYPES[kind]}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
              '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
              '    Set ch = ws.ChartObjects.Add(300, 100, 300, 200).Chart', '    On Error GoTo Failed', '    Select Case which \\ 100']
    for block in range((len(CASES) + 99) // 100):
        lines += [f'    Case {block}: Act{block} which, ws, ch, r']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 30):
            last = min(first + 29, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {statement}', answer, part] for (kind, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
