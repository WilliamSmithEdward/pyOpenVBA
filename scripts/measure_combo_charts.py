"""Combo charts: a series' own ChartType, AxisGroup, and the secondary axes they bring.

Each case makes a workbook whose sheet holds a three-series column chart over A1:D4, with s1, s2 and s3 its series, runs its
statements, and records r, a digest of the chart's state, and the chart part Excel saves.

    python scripts/measure_combo_charts.py

writes tests/fixtures/combo_charts.json, which tests/test_excel_combo_charts.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/combo_charts.json'
LINE = 's2.ChartType = xlLine: '
SECOND = 's2.AxisGroup = xlSecondary: '
BOTH = 's3.ChartType = xlLine: s3.AxisGroup = xlSecondary: '
CASES: list[str] = ['r = ch.ChartType']

# A series' own type: which kinds go with a column, and what the chart reads then.
KINDS = ('xlColumnClustered', 'xlColumnStacked', 'xlColumnStacked100', 'xlBarClustered', 'xlBarStacked', 'xlLine', 'xlLineMarkers', 'xlLineStacked', 'xlLineStacked100', 'xlLineMarkersStacked', 'xlArea', 'xlAreaStacked', 'xlAreaStacked100', 'xlXYScatter', 'xlXYScatterLines',
         'xlXYScatterSmooth', 'xlXYScatterLinesNoMarkers', 'xlXYScatterSmoothNoMarkers', 'xlPie', 'xlDoughnut', 'xlRadar', 'xlRadarMarkers', 'xlRadarFilled', 'xlBubble', 'xl3DColumn', 'xl3DColumnClustered', 'xl3DLine', 'xl3DArea', 'xl3DPie', 'xlSurface', 'xlStockHLC', 'xlPieExploded', 'xlCylinderColClustered')
CASES += [f's2.ChartType = {kind}: r = s2.ChartType & " " & s1.ChartType & " " & ch.ChartType' for kind in KINDS]
CASES += ['s2.ChartType = 0', 's2.ChartType = 9999', 's2.ChartType = "x"', 's2.ChartType = -1']
# Each series its own, the last alone, and back again.
CASES += [f'{LINE}s3.ChartType = xlArea: r = ch.ChartGroups.Count', f'{LINE}s3.ChartType = xlLine: r = ch.ChartGroups.Count', f'{LINE}s2.ChartType = xlColumnClustered: r = ch.ChartType & " " & ch.ChartGroups.Count', f's1.ChartType = xlLine: r = ch.ChartType', f's3.ChartType = xlLine: r = ch.ChartType',
          f's1.ChartType = xlLine: s2.ChartType = xlLine: s3.ChartType = xlLine: r = ch.ChartType & " " & ch.ChartGroups.Count', f'{LINE}s1.ChartType = xlLineMarkers: r = s1.ChartType & " " & s2.ChartType', f'{LINE}s3.ChartType = xlLineMarkers: r = s2.ChartType & " " & s3.ChartType', f'{LINE}s3.ChartType = xlLineStacked: r = s2.ChartType & " " & s3.ChartType',
          's2.ChartType = xlColumnStacked: r = s1.ChartType & " " & s2.ChartType & " " & s3.ChartType', 's2.ChartType = xlBarClustered: r = s1.ChartType & " " & s2.ChartType & " " & s3.ChartType', f'{LINE}ch.ChartType = xlColumnClustered: r = s2.ChartType & " " & ch.ChartGroups.Count', f'{LINE}ch.ChartType = xlLine: r = s1.ChartType & " " & ch.ChartGroups.Count', f'{LINE}ch.ChartType = xlPie: r = s1.ChartType & " " & ch.SeriesCollection.Count']
# The series' order, and members a kind has.
CASES += [f'{LINE}r = s1.PlotOrder & " " & s2.PlotOrder & " " & s3.PlotOrder', f's1.ChartType = xlLine: r = s1.PlotOrder & " " & s2.PlotOrder & " " & s3.PlotOrder & " " & ch.SeriesCollection(1).Name & " " & ch.SeriesCollection(2).Name & " " & ch.SeriesCollection(3).Name', f'{LINE}s3.PlotOrder = 1: r = ch.SeriesCollection(1).Name & " " & ch.SeriesCollection(2).Name',
          f'{LINE}s2.PlotOrder = 1: r = ch.SeriesCollection(1).Name', f'{LINE}r = s2.MarkerStyle & " " & s1.MarkerStyle', f'{LINE}s2.MarkerStyle = xlMarkerStyleCircle: r = s2.MarkerStyle & " " & s2.ChartType', f'{LINE}s2.Smooth = True: r = s2.Smooth & " " & s2.ChartType', f'{LINE}r = s1.Smooth', f'{LINE}r = s2.InvertIfNegative', f'{LINE}r = s1.InvertIfNegative',
          f'{LINE}Set t = ch.SeriesCollection.NewSeries: r = t.ChartType & " " & t.AxisGroup & " " & t.PlotOrder & " " & ch.ChartGroups.Count', f's1.ChartType = xlLine: Set t = ch.SeriesCollection.NewSeries: r = t.ChartType & " " & t.PlotOrder', f'{LINE}s2.Delete: r = ch.ChartType & " " & ch.ChartGroups.Count', f'{LINE}s1.Delete: r = ch.ChartType & " " & ch.ChartGroups.Count', f'{LINE}s1.Delete: s3.Delete: r = ch.ChartType & " " & ch.ChartGroups.Count',
          f'{LINE}ch.SetSourceData ws.Range("A1:C4"): r = ch.SeriesCollection.Count & " " & ch.SeriesCollection(2).ChartType', f'{LINE}ch.SetSourceData ws.Range("A1:B4"): r = ch.SeriesCollection.Count & " " & ch.ChartType', f'{LINE}ch.ApplyLayout 1: r = ch.ChartGroups.Count', f'{LINE}ch.ClearToMatchStyle: r = s2.ChartType', f'{LINE}ch.ChartStyle = 5: r = ch.ChartStyle']

# Groups.
GROUP = 'ch.ChartGroups({0})'
CASES += [f'{LINE}r = {GROUP.format(1)}.SeriesCollection.Count & " " & {GROUP.format(2)}.SeriesCollection.Count & " " & {GROUP.format(1)}.SeriesCollection(1).Name & " " & {GROUP.format(2)}.SeriesCollection(1).Name', f'{LINE}r = TypeName(ch.ChartGroups(3))', f'{LINE}r = {GROUP.format(1)}.Index & " " & {GROUP.format(2)}.Index & " " & {GROUP.format(1)}.AxisGroup & " " & {GROUP.format(2)}.AxisGroup',
          f'{LINE}r = {GROUP.format(1)}.GapWidth & " " & {GROUP.format(1)}.Overlap', f'{LINE}r = {GROUP.format(2)}.GapWidth', f'{LINE}{GROUP.format(1)}.GapWidth = 40: r = {GROUP.format(1)}.GapWidth', f'{LINE}{GROUP.format(1)}.Overlap = 30: r = {GROUP.format(1)}.Overlap', f'{LINE}r = {GROUP.format(2)}.HasDropLines & " " & {GROUP.format(2)}.HasHiLoLines & " " & {GROUP.format(2)}.HasUpDownBars', f'{LINE}{GROUP.format(2)}.HasDropLines = True: r = {GROUP.format(2)}.HasDropLines',
          f'{LINE}r = ch.ColumnGroups.Count & " " & ch.LineGroups.Count & " " & ch.AreaGroups.Count', f'{LINE}r = ch.ColumnGroups(1).SeriesCollection.Count & " " & ch.LineGroups(1).SeriesCollection(1).Name', f'{LINE}r = TypeName(ch.BarGroups(1))', f's2.ChartType = xlArea: r = ch.AreaGroups(1).SeriesCollection(1).Name & " " & ch.ColumnGroups(1).Index & " " & ch.AreaGroups(1).Index', f's2.ChartType = xlXYScatter: r = ch.XYGroups(1).SeriesCollection(1).Name & " " & ch.XYGroups(1).Index', f's2.ChartType = xlPie: r = ch.PieGroups(1).SeriesCollection(1).Name',
          f'{LINE}s3.ChartType = xlArea: r = ch.AreaGroups(1).Index & " " & ch.ColumnGroups(1).Index & " " & ch.LineGroups(1).Index', f'{SECOND}r = ch.ChartGroups.Count & " " & {GROUP.format(1)}.AxisGroup & " " & {GROUP.format(2)}.AxisGroup & " " & {GROUP.format(1)}.SeriesCollection.Count & " " & {GROUP.format(2)}.SeriesCollection.Count', f'{SECOND}r = {GROUP.format(1)}.GapWidth & " " & {GROUP.format(2)}.GapWidth', f'{SECOND}{GROUP.format(2)}.GapWidth = 40: r = {GROUP.format(1)}.GapWidth & " " & {GROUP.format(2)}.GapWidth', f'{SECOND}r = ch.ColumnGroups.Count & " " & ch.ColumnGroups(2).SeriesCollection(1).Name']

# AxisGroup: moving series.
CASES += [f'{SECOND}r = s2.AxisGroup & " " & s1.AxisGroup', 's2.AxisGroup = xlPrimary: r = s2.AxisGroup', f'{SECOND}s2.AxisGroup = xlPrimary: r = s2.AxisGroup & " " & ch.ChartGroups.Count & " " & ch.HasAxis(xlValue, xlSecondary)', 's2.AxisGroup = 0', 's2.AxisGroup = 3', 's2.AxisGroup = "x"', 's2.AxisGroup = -1', 's2.AxisGroup = 2: r = s2.AxisGroup', f'{SECOND}s1.AxisGroup = xlSecondary: r = s1.AxisGroup & " " & s2.AxisGroup & " " & s3.AxisGroup',
          f'{SECOND}s1.AxisGroup = xlSecondary: s3.AxisGroup = xlSecondary: r = s1.AxisGroup & " " & s2.AxisGroup & " " & s3.AxisGroup & " " & ch.ChartGroups.Count', f'{SECOND}r = s1.PlotOrder & " " & s2.PlotOrder & " " & s3.PlotOrder', 's1.AxisGroup = xlSecondary: r = s1.PlotOrder & " " & s2.PlotOrder & " " & ch.SeriesCollection(1).Name', f'{BOTH}r = s3.ChartType & " " & s3.AxisGroup & " " & ch.ChartGroups.Count', 's3.AxisGroup = xlSecondary: s3.ChartType = xlLine: r = s3.ChartType & " " & s3.AxisGroup', f'{BOTH}s3.ChartType = xlColumnClustered: r = s3.AxisGroup & " " & s3.ChartType', f'{BOTH}s2.ChartType = xlLine: r = s2.AxisGroup & " " & ch.ChartGroups.Count', f'{BOTH}s2.ChartType = xlLine: s2.AxisGroup = xlSecondary: r = ch.ChartGroups.Count', f'{SECOND}s2.ChartType = xlLine: r = s2.AxisGroup', f'{SECOND}s1.ChartType = xlLine: r = s1.AxisGroup & " " & s2.AxisGroup', f'{SECOND}s2.ChartType = xlArea: s3.ChartType = xlLine: s3.AxisGroup = xlSecondary: r = ch.ChartGroups.Count',
          f'{SECOND}ch.ChartType = xlLine: r = s2.AxisGroup & " " & ch.HasAxis(xlValue, xlSecondary)', f'{SECOND}ch.ChartType = xlPie: r = s2.AxisGroup', f'{SECOND}ch.ChartType = xlBarClustered: r = s2.AxisGroup & " " & s2.ChartType', f'{SECOND}s2.Delete: r = ch.ChartGroups.Count & " " & ch.HasAxis(xlValue, xlSecondary)', f'{SECOND}ch.SetSourceData ws.Range("A1:C4"): r = ch.SeriesCollection(2).AxisGroup', f'{SECOND}Set t = ch.SeriesCollection.NewSeries: r = t.AxisGroup & " " & t.PlotOrder']
CASES += [f'ch.ChartType = xlPie: ch.SeriesCollection(1).AxisGroup = xlSecondary: r = ch.SeriesCollection(1).AxisGroup', f'ch.ChartType = xlPie: r = ch.SeriesCollection(1).AxisGroup', f'ch.ChartType = xlXYScatter: s2.AxisGroup = xlSecondary: r = s2.AxisGroup & " " & ch.HasAxis(xlCategory, xlSecondary) & " " & ch.HasAxis(xlValue, xlSecondary)', f'ch.ChartType = xlLine: s2.AxisGroup = xlSecondary: r = s2.AxisGroup & " " & s2.ChartType', f'ch.ChartType = xlArea: s2.AxisGroup = xlSecondary: r = s2.AxisGroup', f'ch.ChartType = xlBarClustered: s2.AxisGroup = xlSecondary: r = s2.AxisGroup', f'ch.ChartType = xlRadar: s2.AxisGroup = xlSecondary: r = s2.AxisGroup', f'ch.ChartType = xl3DColumn: s2.AxisGroup = xlSecondary: r = s2.AxisGroup', f'ch.ChartType = xlDoughnut: s2.AxisGroup = xlSecondary: r = s2.AxisGroup']

# Secondary axes.
AXIS = 'ch.Axes({0}, xlSecondary)'
VALUE, CATEGORY = AXIS.format('xlValue'), AXIS.format('xlCategory')
CASES += ['r = ch.HasAxis(xlValue, xlSecondary)', 'r = ch.HasAxis(xlCategory, xlSecondary)', 'r = ch.HasAxis(xlValue, 2) & " " & ch.HasAxis(xlValue, 1)', 'r = TypeName(ch.Axes(xlValue, xlSecondary))', 'r = TypeName(ch.Axes(xlCategory, xlSecondary))', 'ch.HasAxis(xlValue, xlSecondary) = True: r = ch.HasAxis(xlValue, xlSecondary)', 'r = ch.HasAxis(xlValue, 3)', 'r = ch.HasAxis(xlValue, 0)', 'r = ch.HasAxis(3, xlSecondary)', 'r = TypeName(ch.Axes(xlValue, 3))', 'r = ch.Axes.Count',
          f'{SECOND}r = ch.HasAxis(xlValue, xlSecondary) & " " & ch.HasAxis(xlCategory, xlSecondary) & " " & ch.HasAxis(xlValue, xlPrimary) & " " & ch.HasAxis(xlCategory, xlPrimary)', f'{SECOND}r = ch.Axes.Count', f'{SECOND}r = TypeName({VALUE}) & " " & {VALUE}.AxisGroup & " " & {VALUE}.Type', f'{SECOND}r = TypeName({CATEGORY}) & " " & {CATEGORY}.AxisGroup', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: r = TypeName({CATEGORY}) & " " & {CATEGORY}.AxisGroup & " " & {CATEGORY}.Type & " " & ch.Axes.Count', f'{SECOND}ch.HasAxis(xlValue, xlSecondary) = False: r = ch.HasAxis(xlValue, xlSecondary) & " " & s2.AxisGroup & " " & ch.Axes.Count', f'{SECOND}ch.HasAxis(xlValue, xlSecondary) = False: ch.HasAxis(xlValue, xlSecondary) = True: r = ch.HasAxis(xlValue, xlSecondary)', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: ch.HasAxis(xlValue, xlSecondary) = False: r = ch.HasAxis(xlCategory, xlSecondary)',
          f'{SECOND}{VALUE}.Delete: r = ch.HasAxis(xlValue, xlSecondary) & " " & s2.AxisGroup', f'{SECOND}ch.HasAxis(xlValue, xlPrimary) = False: r = ch.HasAxis(xlValue, xlPrimary) & " " & ch.HasAxis(xlValue, xlSecondary)', f'{SECOND}r = ch.Axes(xlValue).AxisGroup & " " & ch.Axes(xlValue, xlPrimary).AxisGroup', f'{SECOND}r = ch.Axes(1).AxisGroup & " " & ch.Axes(2).AxisGroup', f'{SECOND}text = "": For Each o In ch.Axes: text = text & o.Type & "/" & o.AxisGroup & " ": Next: r = text', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: text = "": For Each o In ch.Axes: text = text & o.Type & "/" & o.AxisGroup & " ": Next: r = text', f'{SECOND}r = TypeName(ch.Axes(xlSeriesAxis, xlSecondary))']
READS = ('MinimumScale', 'MaximumScale', 'MajorUnit', 'MinorUnit', 'MinimumScaleIsAuto', 'MaximumScaleIsAuto', 'MajorUnitIsAuto', 'MinorUnitIsAuto', 'Crosses', 'CrossesAt', 'ScaleType', 'ReversePlotOrder', 'HasTitle', 'HasMajorGridlines', 'HasMinorGridlines', 'MajorTickMark', 'MinorTickMark', 'TickLabelPosition', 'TickLabels.NumberFormat', 'TickLabels.Orientation', 'DisplayUnit', 'HasDisplayUnitLabel', 'Border.LineStyle', 'Format.Line.Visible', 'Parent.Name')
CASES += [f'{SECOND}r = {VALUE}.{name}' for name in READS]
CASES += [f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: r = {CATEGORY}.{name}' for name in READS + ('CategoryType', 'TickLabelSpacing', 'TickMarkSpacing', 'AxisBetweenCategories', 'CategoryNames')]
SETS = (('MinimumScale', '5'), ('MaximumScale', '50'), ('MajorUnit', '5'), ('MinorUnit', '1'), ('Crosses', 'xlMaximum'), ('CrossesAt', '3'), ('ScaleType', 'xlScaleLogarithmic'), ('ReversePlotOrder', 'True'), ('HasTitle', 'True'), ('HasMajorGridlines', 'True'), ('HasMinorGridlines', 'True'), ('MajorTickMark', 'xlTickMarkNone'), ('MinorTickMark', 'xlTickMarkCross'), ('TickLabelPosition', 'xlLow'), ('TickLabels.NumberFormat', '"0.0"'), ('DisplayUnit', 'xlThousands'), ('Format.Line.ForeColor.RGB', '255'), ('Border.Color', '255'))
CASES += [f'{SECOND}{VALUE}.{name} = {value}: r = {VALUE}.{name}' for name, value in SETS]
CASES += [f'{SECOND}{VALUE}.HasTitle = True: {VALUE}.AxisTitle.Text = "Right": r = {VALUE}.AxisTitle.Text', f'{SECOND}{VALUE}.MinimumScale = 5: r = ch.Axes(xlValue).MinimumScale', f'{SECOND}ch.Axes(xlValue).MinimumScale = 5: r = {VALUE}.MinimumScale', f'{SECOND}{VALUE}.MinimumScale = 5: s2.AxisGroup = xlPrimary: s2.AxisGroup = xlSecondary: r = {VALUE}.MinimumScale & " " & {VALUE}.MinimumScaleIsAuto', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: {CATEGORY}.Crosses = xlMaximum: r = {CATEGORY}.Crosses', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: {CATEGORY}.ReversePlotOrder = True: r = {CATEGORY}.ReversePlotOrder', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: {CATEGORY}.HasTitle = True: r = {CATEGORY}.AxisTitle.Text']
# The everyday combinations.
CASES += [f'{BOTH}{AXIS.format("xlValue")}.HasTitle = True: {AXIS.format("xlValue")}.AxisTitle.Text = "Rate": ch.HasTitle = True: ch.ChartTitle.Text = "Mix": r = ch.ChartTitle.Text', f'{LINE}s2.HasDataLabels = True: r = s2.DataLabels.Position', f'{LINE}s2.Format.Line.ForeColor.RGB = 255: r = s2.Format.Line.ForeColor.RGB', f'{LINE}s1.Format.Fill.ForeColor.RGB = 255: r = s1.Format.Fill.ForeColor.RGB', f'{LINE}s2.Trendlines.Add: r = s2.Trendlines.Count', f'{LINE}r = s2.HasLeaderLines', f'{LINE}ch.HasLegend = True: r = ch.Legend.LegendEntries.Count', f'{LINE}s2.ErrorBar xlY, xlBoth, xlPercent, 5: r = s2.HasErrorBars', f'{BOTH}r = ch.Legend.LegendEntries.Count', f'{LINE}s2.Points(1).MarkerStyle = xlMarkerStyleSquare: r = s2.Points(1).MarkerStyle', f'{LINE}ws.ChartObjects(1).Copy: ws.Paste ws.Range("H20"): r = ws.ChartObjects.Count', f'{LINE}Set t = ws.ChartObjects(1).Duplicate: r = t.Chart.SeriesCollection(2).ChartType', f'{BOTH}ws.Copy After:=ws: r = ws.Parent.Worksheets(2).ChartObjects(1).Chart.SeriesCollection(3).AxisGroup']

# Which group a new series joins: the last series', the first group, or the chart's first type. Rivals split by where the line is.
NEW = 'Set t = ch.SeriesCollection.NewSeries: r = t.ChartType & " " & t.AxisGroup & " " & t.PlotOrder & " " & ch.ChartGroups.Count'
CASES += [f's3.ChartType = xlLine: {NEW}', f's3.ChartType = xlLine: s3.AxisGroup = xlSecondary: {NEW}', f's1.ChartType = xlLine: s2.ChartType = xlLine: {NEW}', f's1.ChartType = xlArea: s3.ChartType = xlLine: {NEW}', f's1.ChartType = xlLine: s3.ChartType = xlArea: {NEW}', f's1.ChartType = xlLine: s1.AxisGroup = xlSecondary: {NEW}', f's1.ChartType = xlLine: s2.ChartType = xlArea: s3.ChartType = xlArea: {NEW}',
          f's3.ChartType = xlLine: s2.ChartType = xlColumnStacked: {NEW}', f's3.ChartType = xlLine: ch.SetSourceData ws.Range("A1:D3"): r = ch.SeriesCollection.Count & " " & ch.SeriesCollection(2).ChartType']
# What a new group starts with, and which settings each keeps.
CASES += ['ch.ChartGroups(1).GapWidth = 40: s2.AxisGroup = xlSecondary: r = ch.ChartGroups(1).GapWidth & " " & ch.ChartGroups(2).GapWidth', 'ch.ChartGroups(1).Overlap = 30: s2.AxisGroup = xlSecondary: r = ch.ChartGroups(1).Overlap & " " & ch.ChartGroups(2).Overlap', 'ch.ChartGroups(1).GapWidth = 40: s2.ChartType = xlLine: r = ch.ChartGroups(1).GapWidth & " " & ch.ChartGroups(2).GapWidth',
          's2.AxisGroup = xlSecondary: ch.ChartGroups(2).GapWidth = 40: s2.AxisGroup = xlPrimary: r = ch.ChartGroups(1).GapWidth', 's2.AxisGroup = xlSecondary: ch.ChartGroups(1).GapWidth = 40: s2.AxisGroup = xlPrimary: r = ch.ChartGroups(1).GapWidth', 's1.ChartType = xlLine: ch.ChartGroups(1).GapWidth = 40: r = ch.ChartGroups(1).GapWidth & " " & ch.ChartGroups(2).GapWidth', 's2.ChartType = xlLine: ch.ChartGroups(1).GapWidth = 40: s1.ChartType = xlLine: s3.ChartType = xlLine: s1.ChartType = xlColumnClustered: r = ch.ChartGroups(1).GapWidth',
          f'{LINE}{GROUP.format(2)}.HasHiLoLines = True: {GROUP.format(2)}.HasUpDownBars = True: r = {GROUP.format(2)}.HasHiLoLines', f's2.ChartType = xlColumnStacked: s3.ChartType = xlLine: {GROUP.format(1)}.HasSeriesLines = True: r = {GROUP.format(1)}.HasSeriesLines', f'{SECOND}{GROUP.format(2)}.Overlap = 30: r = {GROUP.format(1)}.Overlap & " " & {GROUP.format(2)}.Overlap', f'{LINE}{GROUP.format(1)}.VaryByCategories = True: r = {GROUP.format(1)}.VaryByCategories']
# The axes in a loop, and crossing on the far side.
CASES += [f'{SECOND}text = "": For Each o In ch.Axes: text = text & o.Type & "/" & o.AxisGroup & " ": Next: r = text', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: text = "": For Each o In ch.Axes: text = text & o.Type & "/" & o.AxisGroup & " ": Next: r = text', f'{SECOND}r = ch.Axes(1, 2).AxisGroup & " " & ch.Axes(2, 2).Type', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: {VALUE}.Crosses = xlMinimum: r = {VALUE}.Crosses & " " & {CATEGORY}.Crosses', f'{SECOND}ch.HasAxis(xlCategory, xlSecondary) = True: {CATEGORY}.Crosses = xlMinimum: r = {VALUE}.Crosses & " " & {CATEGORY}.Crosses', f'{SECOND}{VALUE}.Crosses = xlMinimum', f'{SECOND}{CATEGORY}.Crosses = xlMinimum', f'{SECOND}{CATEGORY}.TickLabelPosition = xlLow', f'{SECOND}{CATEGORY}.HasTitle = True']
# On a chart AddChart2 made, as a recorded macro's is.
STYLED = 'ws.ChartObjects(1).Delete: Set co = ws.Shapes.AddChart2(201, xlColumnClustered).Chart.Parent: Set ch = co.Chart: ch.SetSourceData ws.Range("A1:D4"): Set s1 = ch.SeriesCollection(1): Set s2 = ch.SeriesCollection(2): Set s3 = ch.SeriesCollection(3): '
CASES += [f'{STYLED}r = ch.ChartType', f'{STYLED}{LINE}r = s2.ChartType', f'{STYLED}s2.ChartType = xlLineMarkers: r = s2.ChartType', f'{STYLED}s2.ChartType = xlArea: r = s2.ChartType', f'{STYLED}{SECOND}r = s2.AxisGroup', f'{STYLED}{BOTH}r = s3.AxisGroup', f'{STYLED}{BOTH}s1.ChartType = xlArea: r = ch.ChartGroups.Count', f'{STYLED}{LINE}s2.ChartType = xlColumnClustered: r = ch.ChartType', f'{STYLED}{BOTH}s3.AxisGroup = xlPrimary: r = ch.ChartGroups.Count', f'{STYLED}{BOTH}ch.ChartType = xlColumnClustered: r = ch.ChartGroups.Count',
          f'{STYLED}{BOTH}ch.HasAxis(xlCategory, xlSecondary) = True: r = ch.Axes.Count', f'{STYLED}{BOTH}ch.Axes(xlValue, xlSecondary).HasTitle = True: ch.Axes(xlValue, xlSecondary).AxisTitle.Text = "Rate": r = ch.Axes(xlValue, xlSecondary).AxisTitle.Text', f'{STYLED}{BOTH}ch.Axes(xlValue, xlSecondary).HasMajorGridlines = True: r = ch.Axes(xlValue, xlSecondary).HasMajorGridlines', f'{STYLED}{LINE}s2.HasDataLabels = True: r = s2.DataLabels.Position', f'{STYLED}{BOTH}ch.SetElement msoElementSecondaryValueAxisTitleAdjacentToAxis: r = ch.Axes(xlValue, xlSecondary).HasTitle', f'{STYLED}{BOTH}ch.SetElement msoElementSecondaryCategoryAxisShow: r = ch.HasAxis(xlCategory, xlSecondary)', f'{STYLED}{BOTH}ch.SetElement msoElementSecondaryValueAxisNone: r = ch.HasAxis(xlValue, xlSecondary)', f'{STYLED}ch.ChartType = xlLine: s2.ChartType = xlColumnClustered: r = s2.ChartType', f'{STYLED}ch.ChartType = xlLine: s2.AxisGroup = xlSecondary: r = s2.AxisGroup']

MAKE = '''    Set ws = wb.Worksheets(1)
    ws.Range("A1:D1").Value = Array("", "Rev", "Cost", "Rate")
    ws.Range("A2:D2").Value = Array("a", 10, 4, 1)
    ws.Range("A3:D3").Value = Array("b", 20, 6, 2)
    ws.Range("A4:D4").Value = Array("c", 30, 8, 3)
    Set co = ws.ChartObjects.Add(48, 48, 360, 216)
    Set ch = co.Chart
    ch.SetSourceData ws.Range("A1:D4")
    ch.ChartType = xlColumnClustered
    Set s1 = ch.SeriesCollection(1): Set s2 = ch.SeriesCollection(2): Set s3 = ch.SeriesCollection(3)
'''
DIGEST = (('ch.ChartType', 'type'), ('ch.ChartGroups.Count', 'groups'), ('ch.Axes.Count', 'axes'), ('ch.HasAxis(1, 1)', 'c1'), ('ch.HasAxis(2, 1)', 'v1'), ('ch.HasAxis(1, 2)', 'c2'), ('ch.HasAxis(2, 2)', 'v2'), ('ch.SeriesCollection.Count', 'n'))


def probe() -> str:
    lines = ['Private Function Read1(ch As Object) As String', '    Dim s As String, v As Variant, i As Long, o As Object', '    On Error Resume Next']
    for expression, label in DIGEST:
        lines += ['    v = "E": Err.Clear', f'    v = {expression}', f'    If Err.Number Then v = "E" & Err.Number', f'    s = s & "{label}=" & v & " "']
    lines += ['    Err.Clear', '    For i = 1 To ch.SeriesCollection.Count', '        Set o = Nothing: Set o = ch.SeriesCollection(i)', '        v = "E": v = o.Name: s = s & v & ":"', '        v = "E": v = o.ChartType: s = s & v & ":"', '        v = "E": v = o.AxisGroup: s = s & v & ":"', '        v = "E": v = o.PlotOrder: s = s & v & " "', '    Next', '    Err.Clear', '    Read1 = s', 'End Function', '']
    for first in range(0, len(CASES), 60):
        lines += [f'Private Sub Act{first // 60}(ByVal which As Long, ByVal folder As String, ws As Worksheet, co As Object, ch As Object, s1 As Object, s2 As Object, s3 As Object, r As Variant)', '    Dim t As Object, o As Object, n As Long, text As String', '    Select Case which']
        for index, statement in enumerate(CASES[first:first + 60], first):
            lines += [f'    Case {index}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, co As Object, ch As Object, s1 As Object, s2 As Object, s3 As Object, r As Variant, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', MAKE.rstrip('\n'), '    On Error GoTo Failed', '    Select Case which \\ 60']
    lines += [f'    Case {block}: Act{block} which, folder, ws, co, ch, s1, s2, s3, r' for block in range((len(CASES) + 59) // 60)]
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v & "|" & Read1(ws.ChartObjects(1).Chart)', '    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_parts(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name).decode('utf-8') for name in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


def batch(source: str, folder: str, first: int, last: int) -> list[str]:
    """These cases' answers in a session of their own; a batch that fails runs again a case at a time, so that the case that
    took Excel down or stopped at a dialog is recorded as that."""
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        done = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0, line_numbers=False)
        if done.ok and not done.dialogs:
            return str(done.value).split('@@@')
    if first == last:
        return [f'STOPPED {done.outcome}']
    return [answer for index in range(first, last + 1) for answer in batch(source, folder, index, index)]


def main() -> None:
    source = probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        for first in range(0, len(CASES), 25):
            answers += batch(source, folder, first, min(first + 24, len(CASES) - 1))
        parts = [chart_parts(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': [[statement, answer, part] for statement, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statement, answer, _ in record['cases']:
        print(statement, '=>', answer)


if __name__ == '__main__':
    main()
