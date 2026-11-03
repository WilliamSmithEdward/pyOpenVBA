"""The look of a chart group's lines and bars, of tick labels and legend keys, a group's kind-specific settings, and
walking the chart collections: what each reads, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C, and a chart of them of the type it names. It
runs the case's statements, which set r to what they want recorded, reads READS, each as its value or "E" and an error
number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_group_looks.py

writes tests/fixtures/chart_group_looks.json, which tests/test_excel_chart_group_looks.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_group_looks.json'
TYPES = {'column': 'xlColumnClustered', 'line': 'xlLine', 'lineMarkers': 'xlLineMarkers', 'stacked': 'xlColumnStacked', 'scatter': 'xlXYScatter', 'pie': 'xlPie',
         'area': 'xlArea', 'bubble': 'xlBubble', 'radar': 'xlRadar', 'pieOfPie': 'xlPieOfPie', 'barOfPie': 'xlBarOfPie'}
G = 'ch.ChartGroups(1)'
S1 = 'ch.SeriesCollection(1)'
KEY = 'ch.Legend.LegendEntries(1).LegendKey'
#: The element, the chart it takes, and what puts it there.
LINES = {'DropLines': ('line', f'{G}.HasDropLines = True'), 'HiLoLines': ('line', f'{G}.HasHiLoLines = True'), 'UpBars': ('line', f'{G}.HasUpDownBars = True'),
         'DownBars': ('line', f'{G}.HasUpDownBars = True'), 'SeriesLines': ('stacked', f'{G}.HasSeriesLines = True')}
BORDER_SETS = ['.Color = 255', '.ColorIndex = 3', '.ColorIndex = xlNone', '.ColorIndex = xlAutomatic', '.Weight = xlHairline', '.Weight = xlThin', '.Weight = xlMedium', '.Weight = xlThick',
               '.LineStyle = xlNone', '.LineStyle = xlContinuous', '.LineStyle = xlDash', '.LineStyle = xlDot', '.LineStyle = xlAutomatic']
LINE_READS = '{0}.Line.Visible & " " & {0}.Line.ForeColor.RGB & " " & {0}.Line.Weight & " " & {0}.Line.DashStyle & " " & {0}.Line.Transparency'
FILL_READS = '{0}.Fill.Visible & " " & {0}.Fill.ForeColor.RGB & " " & {0}.Fill.Type & " " & {0}.Fill.Transparency'
LINE_SETS = ['.Line.ForeColor.RGB = 255', '.Line.Weight = 3', '.Line.DashStyle = msoLineDash', '.Line.Visible = msoFalse', '.Line.Visible = msoTrue', '.Line.Transparency = 0.5']
FILL_SETS = ['.Fill.ForeColor.RGB = 255', '.Fill.Visible = msoFalse', '.Fill.Visible = msoTrue', '.Fill.Transparency = 0.5', '.Fill.Solid']
CASES: list[tuple[str, str]] = []
for name, (kind, before) in LINES.items():
    element = f'{G}.{name}'
    border = f'{element}.Border'
    shown = f'{border}.Color & " " & {border}.ColorIndex & " " & {border}.Weight & " " & {border}.LineStyle'
    look = f'{element}.Format'
    CASES.append((kind, f'{before}: r = TypeName({border}) & " " & TypeName({look}) & " " & TypeName({border}.Parent) & " " & TypeName({look}.Parent)'))
    CASES.append((kind, f'{before}: r = {shown}'))
    CASES += [(kind, f'{before}: {border}{change}: r = {shown}') for change in BORDER_SETS]
    CASES.append((kind, f'{before}: {border}.Color = 255: {border}.Weight = xlThick: {border}.LineStyle = xlDash: r = {shown}'))
    CASES.append((kind, f'{before}: r = {LINE_READS.format(look)}'))
    CASES.append((kind, f'{before}: r = {FILL_READS.format(look)}'))
    CASES += [(kind, f'{before}: {look}{change}: r = {LINE_READS.format(look)} & " / " & {shown}') for change in LINE_SETS]
    CASES += [(kind, f'{before}: {look}{change}: r = {FILL_READS.format(look)}') for change in FILL_SETS]
    CASES.append((kind, f'{before}: r = {element}.Select & " " & TypeName(Selection) & " " & Selection.Name'))
    CASES.append((kind, f'{G}.Has{name if name in ("DropLines", "HiLoLines", "SeriesLines") else "UpDownBars"} = False: r = TypeName({element}.Border)'))
    CASES.append((kind, f'{before}: r = {element}.GetProperty("Name")'))
    CASES.append((kind, f'{before}: {element}.SetProperty "Name", "z": r = {element}.Name'))
    CASES.append((kind, f'{before}: {look}.Line.ForeColor.RGB = 255: {element}.Delete: {before}: r = {LINE_READS.format(look)}'))

# Tick labels.
for axis in ('xlValue', 'xlCategory'):
    ticks = f'ch.Axes({axis}).TickLabels'
    CASES += [('column', f'r = TypeName({ticks}.Format) & " " & TypeName({ticks}.Format.Parent)'), ('column', f'r = {LINE_READS.format(f"{ticks}.Format")}'), ('column', f'r = {FILL_READS.format(f"{ticks}.Format")}'),
              ('column', f'{ticks}.Format.Fill.ForeColor.RGB = 255: r = {FILL_READS.format(f"{ticks}.Format")}'), ('column', f'{ticks}.Format.Line.ForeColor.RGB = 255: r = {LINE_READS.format(f"{ticks}.Format")}'),
              ('column', f'r = {ticks}.Select & " " & TypeName(Selection) & " " & Selection.Name'), ('column', f'r = {ticks}.Delete'), ('column', f'{ticks}.Delete: r = ch.Axes({axis}).TickLabelPosition')]

# ErrorBars.ClearFormats.
BARS = f'{S1}.ErrorBars'
SETUP = f'{S1}.ErrorBar xlY, xlBoth, xlFixedValue, 1: '
BAR_READS = f'r = {BARS}.Border.Color & " " & {BARS}.Border.Weight & " " & {BARS}.Border.LineStyle & " " & {BARS}.Format.Line.Visible & " " & {BARS}.EndStyle'
CASES += [('column', f'{SETUP}r = {BARS}.ClearFormats'), ('column', f'{SETUP}{BARS}.ClearFormats: {BAR_READS}')]
CASES += [('column', f'{SETUP}{change}: {BARS}.ClearFormats: {BAR_READS}') for change in (f'{BARS}.Border.Color = 255', f'{BARS}.Border.Weight = xlThick', f'{BARS}.Format.Line.Visible = msoFalse', f'{BARS}.EndStyle = xlNoCap')]

# Walking the collections.
WALK = 'n = 0: For Each o In {0}: n = n + 1: t = TypeName(o): Next: r = n & " " & t'
CASES += [('column', WALK.format('ch.Legend.LegendEntries')), ('pie', WALK.format('ch.Legend.LegendEntries')), ('column', WALK.format('ch.ChartGroups')), ('column', WALK.format('ws.ChartObjects')),
          ('column', f'{S1}.HasDataLabels = True: {WALK.format(f"{S1}.DataLabels")}'), ('column', f'ch.Legend.LegendEntries(1).Delete: {WALK.format("ch.Legend.LegendEntries")}'),
          ('column', 'r = TypeName(ch.Legend.LegendEntries.Parent) & " " & TypeName(ch.Legend.LegendEntries.Application)'), ('column', 'r = TypeName(ch.ChartGroups.Parent) & " " & TypeName(ch.ChartGroups.Application)'),
          ('column', 'r = TypeName(ch.Legend.LegendEntries.[_Default](1))'), ('column', 'r = TypeName(ch.Legend.LegendEntries(1).Format) & " " & TypeName(ch.Legend.LegendEntries(1).Format.Parent)'),
          ('column', f'r = {FILL_READS.format("ch.Legend.LegendEntries(1).Format")}')]

# The legend key.
MARKERS = f'{KEY}.MarkerBackgroundColor & " " & {KEY}.MarkerBackgroundColorIndex & " " & {KEY}.MarkerForegroundColor & " " & {KEY}.MarkerForegroundColorIndex & " " & {KEY}.MarkerSize'
for kind in ('column', 'line', 'lineMarkers', 'scatter'):
    CASES += [(kind, f'r = {MARKERS}'), (kind, f'{KEY}.MarkerSize = 12: r = {MARKERS} & " / " & {S1}.MarkerSize'), (kind, f'{KEY}.MarkerBackgroundColor = 255: r = {MARKERS} & " / " & {S1}.MarkerBackgroundColor'),
              (kind, f'{KEY}.MarkerForegroundColor = 255: r = {MARKERS} & " / " & {S1}.MarkerForegroundColor'), (kind, f'{KEY}.MarkerBackgroundColorIndex = 3: r = {MARKERS}'), (kind, f'{KEY}.MarkerForegroundColorIndex = 3: r = {MARKERS}'),
              (kind, f'{KEY}.Shadow = True: r = {KEY}.Shadow & " " & {S1}.Shadow'), (kind, f'r = {KEY}.Shadow & " " & TypeName({KEY}.Application)'), (kind, f'r = {KEY}.PictureType & " " & {KEY}.PictureUnit2'),
              (kind, f'{S1}.Format.Fill.ForeColor.RGB = 255: r = {KEY}.ClearFormats & " " & {S1}.Format.Fill.ForeColor.RGB'), (kind, f'{KEY}.ClearFormats: r = {MARKERS}'), (kind, f'{KEY}.MarkerStyle = xlMarkerStyleSquare: r = {KEY}.MarkerStyle & " " & {S1}.MarkerStyle')]

# What a group has for one kind of chart.
SETTINGS = {'bubble': ['BubbleScale', 'ShowNegativeBubbles', 'SizeRepresents'], 'pieOfPie': ['SplitType', 'SplitValue', 'SecondPlotSize'], 'barOfPie': ['SplitType', 'SplitValue', 'SecondPlotSize'], 'radar': ['HasRadarAxisLabels']}
VALUES = {'BubbleScale': [50, 300, 0, 301, -1], 'ShowNegativeBubbles': [True, False], 'SizeRepresents': [1, 2, 3], 'SplitType': [1, 2, 3, 4, 5], 'SplitValue': [1, 2, 50, 0], 'SecondPlotSize': [5, 75, 200, 4, 201], 'HasRadarAxisLabels': [True, False]}
for kind, names in SETTINGS.items():
    for name in names:
        CASES.append((kind, f'r = {G}.{name}'))
        CASES += [(kind, f'{G}.{name} = {value}: r = {G}.{name}') for value in VALUES[name]]
for name in VALUES:
    CASES += [(kind, f'r = {G}.{name}') for kind in ('column', 'line')]
    CASES.append(('column', f'{G}.{name} = {VALUES[name][0]}: r = {G}.{name}'))
CASES += [('radar', f'r = TypeName({G}.RadarAxisLabels) & " " & TypeName({G}.RadarAxisLabels.Parent)'), ('column', f'r = TypeName({G}.RadarAxisLabels)'), ('radar', f'{G}.HasRadarAxisLabels = False: r = TypeName({G}.RadarAxisLabels)'),
          ('radar', f'r = {G}.Has3DShading'), ('column', f'r = {G}.Has3DShading'), ('bubble', f'r = {G}.Has3DShading'), ('bubble', f'{G}.Has3DShading = True: r = {G}.Has3DShading'), ('column', f'{G}.Has3DShading = True: r = {G}.Has3DShading'),
          ('column', f'r = TypeName({G}.CategoryCollection) & " " & TypeName({G}.FullCategoryCollection)'), ('column', f'r = {G}.CategoryCollection.Count & " " & {G}.FullCategoryCollection.Count')]

# Outside what the group's kind draws, and looks set over looks.
CASES += [('column', f'{G}.{name} = {value}: r = {G}.{name}') for name, value in (('BubbleScale', 301), ('BubbleScale', -1), ('SecondPlotSize', 4), ('SecondPlotSize', 201), ('SplitType', 5), ('SplitType', 3), ('SizeRepresents', 3), ('SizeRepresents', 2), ('SplitValue', -1), ('SplitValue', 2.5), ('Has3DShading', False), ('ShowNegativeBubbles', False))]
CASES += [('column', f'{G}.SplitType = 3: r = {G}.SplitValue'), ('column', f'{G}.SplitType = 4: r = {G}.SplitValue'), ('column', f'{G}.SplitValue = 7: {G}.SplitType = 2: r = {G}.SplitValue'), ('column', f'{G}.BubbleScale = 50: ch.ChartType = xlLine: r = {G}.BubbleScale')]
for name, flag in (('UpBars', 'HasUpDownBars'), ('DownBars', 'HasUpDownBars')):
    before = f'{G}.{flag} = True'
    look = f'{G}.{name}.Format'
    fills = FILL_READS.format(look)
    lines = LINE_READS.format(look)
    CASES += [('line', f'{before}: {look}.Fill.ForeColor.RGB = 255: {look}.Line.ForeColor.RGB = 65280: r = {fills} & " / " & {lines}'), ('line', f'{before}: {look}.Line.ForeColor.RGB = 65280: {look}.Fill.ForeColor.RGB = 255: r = {fills} & " / " & {lines}'),
              ('line', f'{before}: {look}.Fill.ForeColor.RGB = 255: {look}.Fill.Transparency = 0.5: r = {fills}'), ('line', f'{before}: {look}.Fill.ForeColor.RGB = 255: {look}.Fill.Visible = msoFalse: r = {fills}'),
              ('line', f'{before}: {look}.Fill.Visible = msoFalse: {look}.Fill.ForeColor.RGB = 255: r = {fills}'), ('line', f'{before}: {look}.Fill.Visible = msoFalse: {look}.Fill.Visible = msoTrue: r = {fills}'),
              ('line', f'{before}: {look}.Line.Visible = msoFalse: r = {fills}'), ('line', f'{before}: {look}.Fill.Visible = msoFalse: r = {lines}'), ('line', f'{before}: {look}.Line.Weight = 3: r = {fills}'),
              ('line', f'{before}: {look}.Fill.ForeColor.RGB = 255: {G}.HasUpDownBars = False: {before}: r = {fills}'), ('line', f'{before}: {look}.Fill.ForeColor.RGB = 255: ch.ChartType = xlColumnClustered: ch.ChartType = xlLine: r = {fills}'),
              ('line', f'{before}: r = TypeName({look}.Fill) & " " & TypeName({look}.Fill.Parent) & " " & {look}.Fill.BackColor.RGB & " " & {look}.Fill.ForeColor.ObjectThemeColor & " " & {look}.Fill.ForeColor.Type')]
for name, flag, kind in (('DropLines', 'HasDropLines', 'line'), ('HiLoLines', 'HasHiLoLines', 'line'), ('SeriesLines', 'HasSeriesLines', 'stacked')):
    before = f'{G}.{flag} = True'
    look = f'{G}.{name}.Format'
    lines = LINE_READS.format(look)
    CASES += [(kind, f'{before}: {look}.Line.ForeColor.RGB = 255: {G}.{flag} = False: {before}: r = {lines}'), (kind, f'{before}: {look}.Line.ForeColor.RGB = 255: ch.ChartType = xlPie: ch.ChartType = {TYPES[kind]}: r = {lines}'),
              (kind, f'{before}: {look}.Line.ForeColor.RGB = 255: {look}.Line.Visible = msoFalse: r = {lines}'), (kind, f'{before}: {look}.Line.Visible = msoFalse: {look}.Line.Visible = msoTrue: r = {lines}'),
              (kind, f'{before}: {look}.Line.Visible = msoFalse: {look}.Line.ForeColor.RGB = 255: r = {lines}'), (kind, f'{before}: {G}.{name}.Border.LineStyle = xlNone: {look}.Line.ForeColor.RGB = 255: r = {lines}')]

# GetProperty and SetProperty, which every chart element lists.
OWNERS = {'ch': '', 'ch.PlotArea': '', 'ch.Legend': '', 'ch.ChartTitle': 'ch.HasTitle = True: ', 'ch.Axes(xlValue)': '', 'ch.Axes(xlValue).AxisTitle': 'ch.Axes(xlValue).HasTitle = True: ', S1: '', f'{S1}.Points(1)': '',
          f'{S1}.DataLabels': f'{S1}.HasDataLabels = True: ', f'{S1}.Points(1).DataLabel': f'{S1}.Points(1).HasDataLabel = True: ', f'{S1}.ErrorBars': f'{S1}.ErrorBar xlY, xlBoth, xlFixedValue, 1: '}
for owner, before in OWNERS.items():
    CASES += [('column', f'{before}r = {owner}.GetProperty("Name")'), ('column', f'{before}r = {owner}.GetProperty("x")'), ('column', f'{before}{owner}.SetProperty "Name", "z": r = {owner}.Name'), ('column', f'{before}{owner}.SetProperty "x", 1: r = {owner}.GetProperty("x")')]
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
