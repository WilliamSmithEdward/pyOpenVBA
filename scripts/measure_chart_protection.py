"""A chart's protection and page setup, and the effects in a chart element's Format, one member at a time: what each reads,
and the chart part Excel saves. The cases run as scripts/measure_chart_members_more.py runs its own.

    python scripts/measure_chart_protection.py

writes tests/fixtures/chart_protection.json, which tests/test_excel_chart_protection.py replays.
"""
from __future__ import annotations

from pathlib import Path

import measure_chart_members_more as runner

S1 = 'ch.SeriesCollection(1)'
CASES: list[tuple[str, str]] = []

# Each protection member alone, before and after each way of changing it.
LOCKS = ('ProtectContents', 'ProtectDrawingObjects', 'ProtectFormatting', 'ProtectData', 'ProtectGoalSeek', 'ProtectSelection', 'ProtectionMode', 'ProtectScenarios', 'ChartObjects.Count')
WAYS = ('', 'ch.Protect', 'ch.Protect: ch.Unprotect', 'ch.Unprotect', 'ch.Protect "pw"', 'ch.Protect "pw": ch.Unprotect "pw"', 'ch.Protect "pw": ch.Unprotect "no"', 'ch.Protect "pw": ch.Unprotect', 'ch.Protect , True, True, True, True',
        'ch.Protect DrawingObjects:=False, Contents:=True', 'ch.Protect DrawingObjects:=True, Contents:=False', 'ch.Protect Scenarios:=True', 'ch.Protect UserInterfaceOnly:=True', 'ch.Protect: ch.Protect', 'ws.Protect', 'ws.Protect: ch.Protect', 'ch.Protect: ws.Protect')
for way in WAYS:
    for lock in LOCKS[:7]:
        CASES.append(('column', f'{way}{": " if way else ""}r = ch.{lock}'))
for lock in ('ProtectFormatting', 'ProtectData', 'ProtectGoalSeek', 'ProtectSelection'):
    CASES += [('column', f'ch.{lock} = True: r = ch.{lock}'), ('column', f'ch.{lock} = True: ch.{lock} = False: r = ch.{lock}'), ('column', f'ch.Protect: ch.{lock} = True: r = ch.{lock}'), ('column', f'ch.{lock} = True: ch.Protect: ch.Unprotect: r = ch.{lock}')]

# What each protection stops.
EDITS = (f'{S1}.Name = "z"', 'ch.HasTitle = True', 'ch.ChartType = xlLine', f'{S1}.Format.Fill.ForeColor.RGB = 255', 'ch.HasLegend = False', 'ch.SetSourceData ws.Range("A1:B4")', 'ch.ChartArea.Format.Fill.ForeColor.RGB = 255', 'ws.Range("B2").Value = 9', 'ch.Parent.Left = 20', 'ch.Parent.Width = 200')
READBACKS = (f'{S1}.Name', 'ch.HasTitle', 'ch.ChartType', f'{S1}.Format.Fill.ForeColor.RGB', 'ch.HasLegend', 'ch.SeriesCollection.Count', 'ch.ChartArea.Format.Fill.ForeColor.RGB', f'{S1}.Values(1)', 'ch.Parent.Left', 'ch.Parent.Width')
for lock in ('ch.Protect', 'ch.Protect DrawingObjects:=False, Contents:=True', 'ch.Protect DrawingObjects:=True, Contents:=False', 'ch.ProtectFormatting = True', 'ch.ProtectData = True', 'ch.ProtectSelection = True', 'ws.Protect'):
    CASES += [('column', f'{lock}: {edit}: r = {read}') for edit, read in zip(EDITS, READBACKS, strict=True)]

# The chart's page setup.
PAGE = 'ch.PageSetup'
CASES += [('column', f'r = TypeName({PAGE}) & " " & TypeName({PAGE}.Parent) & " " & TypeName({PAGE}.Application)'), ('column', f'r = TypeName(ws.PageSetup.Parent)')]
READS_PAGE = ('Orientation', 'ChartSize', 'Zoom', 'FitToPagesWide', 'FitToPagesTall', 'PaperSize', 'LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin', 'HeaderMargin', 'FooterMargin', 'CenterHorizontally', 'CenterVertically', 'CenterHeader', 'LeftFooter', 'FirstPageNumber', 'PrintArea', 'PrintTitleRows', 'PrintGridlines', 'BlackAndWhite', 'Draft', 'Order', 'PrintQuality', 'PrintNotes')
CASES += [('column', f'r = {PAGE}.{name}') for name in READS_PAGE]
SETS_PAGE = (('Orientation', 'xlLandscape'), ('Orientation', 'xlPortrait'), ('ChartSize', 'xlFullPage'), ('ChartSize', 'xlFitToPage'), ('Zoom', '50'), ('Zoom', 'False'), ('FitToPagesWide', '2'), ('PaperSize', 'xlPaperA4'), ('LeftMargin', '36'), ('TopMargin', '72'), ('BottomMargin', '20'), ('HeaderMargin', '10'),
             ('CenterHorizontally', 'True'), ('CenterHeader', '"Top"'), ('LeftFooter', '"Foot"'), ('FirstPageNumber', '3'), ('BlackAndWhite', 'True'), ('Draft', 'True'), ('PrintArea', '"A1:B2"'), ('PrintTitleRows', '"$1:$1"'))
CASES += [('column', f'{PAGE}.{name} = {value}: r = {PAGE}.{name}') for name, value in SETS_PAGE]
CASES += [('column', f'{PAGE}.Orientation = xlLandscape: r = ws.PageSetup.Orientation'), ('column', f'ws.PageSetup.Orientation = xlLandscape: r = {PAGE}.Orientation'), ('column', f'{PAGE}.LeftMargin = 36: r = ws.PageSetup.LeftMargin'), ('line', f'r = {PAGE}.Orientation'), ('pie', f'{PAGE}.Orientation = xlLandscape: r = {PAGE}.Orientation')]

# Effects, on each kind of element.
ELEMENTS = {'series': (S1, ''), 'point': (f'{S1}.Points(2)', ''), 'area': ('ch.ChartArea', ''), 'plot': ('ch.PlotArea', ''), 'legend': ('ch.Legend', ''), 'title': ('ch.ChartTitle', 'ch.HasTitle = True: '), 'axis': ('ch.Axes(xlValue)', ''), 'grid': ('ch.Axes(xlValue).MajorGridlines', ''),
            'axis title': ('ch.Axes(xlValue).AxisTitle', 'ch.Axes(xlValue).HasTitle = True: '), 'labels': (f'{S1}.DataLabels', f'{S1}.HasDataLabels = True: '), 'trendline': (f'{S1}.Trendlines(1)', f'{S1}.Trendlines.Add: '), 'error bars': (f'{S1}.ErrorBars', f'{S1}.ErrorBar xlY, xlBoth, xlFixedValue, 1: '), 'ticks': ('ch.Axes(xlValue).TickLabels', '')}
for element, before in ELEMENTS.values():
    look = f'{element}.Format'
    CASES += [('column', f'{before}r = TypeName({look}.{member})') for member in ('Glow', 'Shadow', 'SoftEdge', 'ThreeD', 'TextFrame2', 'Adjustments', 'PictureFormat')]
    CASES += [('column', f'{before}r = {look}.AutoShapeType'), ('column', f'{before}r = {look}.Glow.Radius & " " & {look}.Glow.Color.RGB & " " & {look}.Glow.Transparency'), ('column', f'{before}r = {look}.Shadow.Visible & " " & {look}.Shadow.Type & " " & {look}.Shadow.Blur & " " & {look}.Shadow.OffsetX & " " & {look}.Shadow.OffsetY'),
              ('column', f'{before}r = {look}.SoftEdge.Type & " " & {look}.SoftEdge.Radius'), ('column', f'{before}r = {look}.ThreeD.Visible & " " & {look}.ThreeD.BevelTopType'), ('column', f'{before}{look}.Glow.Radius = 5: r = {look}.Glow.Radius'), ('column', f'{before}{look}.Glow.Color.RGB = 255: r = {look}.Glow.Color.RGB'),
              ('column', f'{before}{look}.Shadow.Visible = msoTrue: r = {look}.Shadow.Visible & " " & {look}.Shadow.Type & " " & {look}.Shadow.Blur'), ('column', f'{before}{look}.Shadow.Type = msoShadow21: r = {look}.Shadow.Type'), ('column', f'{before}{look}.SoftEdge.Type = msoSoftEdgeType1: r = {look}.SoftEdge.Type & " " & {look}.SoftEdge.Radius'),
              ('column', f'{before}{look}.SoftEdge.Radius = 4: r = {look}.SoftEdge.Radius'), ('column', f'{before}{look}.ThreeD.BevelTopType = msoBevelCircle: r = {look}.ThreeD.BevelTopType'), ('column', f'{before}r = {look}.TextFrame2.HasText'), ('column', f'{before}r = {look}.Adjustments.Count'), ('column', f'{before}{look}.AutoShapeType = 5: r = {look}.AutoShapeType')]
READS = ['ch.SeriesCollection.Count', f'{S1}.Name']


def main() -> None:
    runner.OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_protection.json'
    runner.CASES[:] = CASES
    runner.main()


if __name__ == '__main__':
    main()
