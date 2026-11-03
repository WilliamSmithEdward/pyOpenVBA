"""Chart.SetElement, every element on three kinds of chart, each from a fresh chart: what it answers, what the chart reads
afterwards, and the chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, of the chart type TYPES names,
calls SetElement, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart part is
recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_set_element.py

writes tests/fixtures/chart_set_element.json, which tests/test_excel_chart_set_element.py replays. Given kinds of chart

    python scripts/measure_chart_set_element.py bar line

it measures those alone, with no prepared cases, into tests/fixtures/chart_set_element_<kinds>.json.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.interpreter._constants_data import __dict__ as _DATA  # noqa: E402

OUT = ROOT / 'tests/fixtures/chart_set_element.json'
OTHER_TYPES = {'bar': 'xlBarClustered', 'plain': 'xlLine', 'scatter': 'xlXYScatter', 'area': 'xlArea', 'doughnut': 'xlDoughnut'}
CHOSEN = sys.argv[1:]
TYPES = {kind: OTHER_TYPES[kind] for kind in CHOSEN} or {'column': 'xlColumnClustered', 'line': 'xlLineMarkers', 'pie': 'xlPie'}
if CHOSEN:
    OUT = OUT.with_name(f'chart_set_element_{"_".join(CHOSEN)}.json')
ELEMENTS = sorted(name for table in _DATA.values() if isinstance(table, dict) for name in table if str(name).startswith('msoElement'))
PREPARED = [('column', 'ch.HasTitle = True', 'msoElementChartTitleNone'), ('column', 'ch.HasLegend = True', 'msoElementLegendNone'),
            ('column', 'ch.SeriesCollection(1).HasDataLabels = True', 'msoElementDataLabelNone'),
            ('column', 'ch.SetElement msoElementChartTitleAboveChart', 'msoElementChartTitleCenteredOverlay'), ('column', 'ch.SetElement msoElementLegendTop', 'msoElementLegendLeft'),
            ('column', 'ch.SetElement msoElementDataLabelShow', 'msoElementDataLabelCenter'), ('column', 'ch.SetElement msoElementPrimaryValueAxisTitleRotated', 'msoElementPrimaryValueAxisTitleNone'),
            ('column', 'ch.HasTitle = True: ch.ChartTitle.Text = "T"', 'msoElementChartTitleAboveChart'), ('column', 'ch.Axes(xlValue).HasMajorGridlines = False', 'msoElementPrimaryValueGridLinesMajor')]
CASES = [(kind, '', element) for kind in TYPES for element in ELEMENTS]
CASES += [] if CHOSEN else [*PREPARED, ('column', '', '0'), ('column', '', '9999')]
READS = ['ch.HasTitle', 'ch.ChartTitle.Text', 'ch.ChartTitle.IncludeInLayout', 'ch.HasLegend', 'ch.Legend.Position', 'ch.Legend.IncludeInLayout', 'ch.HasAxis(xlCategory)', 'ch.HasAxis(xlValue)',
         'ch.Axes(xlValue).HasMajorGridlines', 'ch.Axes(xlValue).HasMinorGridlines', 'ch.Axes(xlCategory).HasMajorGridlines', 'ch.Axes(xlCategory).HasMinorGridlines', 'ch.Axes(xlValue).HasTitle',
         'ch.Axes(xlCategory).HasTitle', 'ch.Axes(xlValue).DisplayUnit', 'ch.Axes(xlValue).ScaleType', 'ch.Axes(xlCategory).ReversePlotOrder', 'ch.SeriesCollection(1).HasDataLabels',
         'ch.SeriesCollection(1).DataLabels.Position', 'ch.SeriesCollection(2).HasDataLabels', 'ch.SeriesCollection(1).HasErrorBars', 'ch.SeriesCollection(1).Trendlines.Count', 'ch.HasDataTable',
         'ch.ChartGroups(1).HasDropLines', 'ch.ChartGroups(1).HasHiLoLines', 'ch.ChartGroups(1).HasUpDownBars', 'ch.ChartGroups(1).HasSeriesLines', 'ch.SeriesCollection.Count', 'ch.ChartType',
         'TypeName(Selection)']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, before, element) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {TYPES[kind]}']
        lines += [f'        {before}'] if before else []
        lines += [f'        ch.SetElement {element}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "," & v']
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
              'cases': [[f'{kind}: {before}: {element}', answer, part] for (kind, before, element), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
