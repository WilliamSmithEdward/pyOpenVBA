"""The chart part Excel saves for a chart made through ChartObjects.Add, for each kind the model writes and each edit
it models: titles, the legend, series made with NewSeries, axis titles and scales, a hidden axis, a deleted or
reordered series. Each case also records what the chart reads back.

Each case makes a workbook, fills cells from one of the DATA blocks, adds a chart, gives it that block with
SetSourceData (unless the case has none), sets its type, runs the case's statements and saves it. The chart part is
recorded with its axis ids numbered in order of appearance and its series' unique ids cut to their first group, as
both are random in Excel.

    python scripts/measure_chart_markup.py

writes tests/fixtures/chart_markup.json, which tests/test_excel_chart_markup.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_markup.json'
#: Each block's columns, A first; the source is the block's used cells.
DATA = {
    'numbers': (('k', 1, 2, 4), ('v', 1, 2, 3)),
    'one': (('k', 'a', 'b', 'c'), ('v', 1, 2, 3)),
    'two': (('k', 'a', 'b', 'c'), ('v', 1, 2, 3), ('w', 4, 0.5, 6)),
}
READ = ('ch.ChartType & "," & ch.HasTitle & "," & ch.HasLegend & "," & ch.SeriesCollection.Count & "," & '
        'ch.PlotBy')
#: (data block or None, plot by or None, chart type, statements)
CASES = [
    ('numbers', None, 'xlColumnClustered', ''),
    ('numbers', None, 'xlLine', ''),
    ('numbers', None, 'xlPie', ''),
    ('numbers', None, 'xlXYScatter', ''),
    ('numbers', None, 'xlArea', ''),
    ('numbers', None, 'xlBarClustered', ''),
    ('numbers', None, 'xlDoughnut', ''),
    ('numbers', None, 'xlLineMarkers', ''),
    ('one', None, 'xlColumnClustered', ''),
    ('one', None, 'xlPie', ''),
    ('one', None, 'xlLine', ''),
    ('one', None, 'xlXYScatter', ''),
    ('one', None, 'xlDoughnut', ''),
    ('one', None, 'xlArea', ''),
    ('two', None, 'xlColumnClustered', ''),
    ('two', 'xlRows', 'xlBarClustered', ''),
    ('two', None, 'xlLineMarkers', ''),
    ('two', None, 'xlXYScatter', ''),
    ('numbers', None, 'xlColumnClustered', 'ch.HasTitle = True'),
    ('numbers', None, 'xlColumnClustered', 'ch.HasTitle = True: ch.ChartTitle.Text = "Revenue & Expenses"'),
    ('one', None, 'xlColumnClustered', 'ch.HasTitle = False'),
    ('one', None, 'xlColumnClustered', 'ch.ChartTitle.Text = "Mine"'),
    ('numbers', None, 'xlColumnClustered', 'ch.HasLegend = False'),
    ('one', None, 'xlPie', 'ch.HasLegend = False: ch.HasTitle = False'),
    (None, None, 'xlColumnClustered', ''),
    (None, None, 'xlColumnClustered',
     'With ch.SeriesCollection.NewSeries: .Name = "Revenue": .Values = Array(2, 4): .XValues = Array("a", "b"): '
     'End With'),
    (None, None, 'xlLine', 'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4"'),
    ('two', None, 'xlColumnClustered',
     'ch.Axes(1).HasTitle = True: ch.Axes(2).HasTitle = True: ch.Axes(2).AxisTitle.Text = "Units"'),
    ('two', None, 'xlColumnClustered',
     'ch.Axes(2).MinimumScale = 0.5: ch.Axes(2).MaximumScale = 10: ch.Axes(2).MajorUnit = 2: '
     'ch.Axes(1).ReversePlotOrder = True'),
    ('two', None, 'xlXYScatter', 'ch.Axes(1).MaximumScale = 7.25'),
    ('two', None, 'xlColumnClustered', 'ch.HasAxis(xlCategory) = False'),
    ('two', None, 'xlColumnClustered', 'ch.HasAxis(xlValue) = False'),
    ('two', None, 'xlColumnClustered', 'ch.SeriesCollection(1).Delete'),
    ('two', None, 'xlColumnClustered', 'ch.SeriesCollection(2).PlotOrder = 1'),
    ('two', None, 'xlLine', 'ch.SeriesCollection(1).Name = "Renamed"'),
    ('two', None, 'xlColumnClustered', 'ch.ChartType = xlPie: ch.ChartType = xlColumnClustered'),
    (None, None, 'xlLine', ''),
    (None, None, 'xlLine', 'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$D$2"'),
    (None, None, 'xlColumnClustered',
     'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4": '
     'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$C$2:$C$4"'),
    (None, None, 'xlColumnClustered',
     'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$C$2:$C$4": '
     'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4"'),
    ('two', None, 'xlColumnClustered', 'ch.SeriesCollection(1).Name = "=Sheet1!$A$1"'),
    ('one', None, 'xlColumnClustered', 's = ch.ChartTitle.Text'),
    ('one', None, 'xlColumnClustered', 's = TypeName(ch.ChartTitle)'),
    ('one', None, 'xlColumnClustered', 'ch.HasTitle = True: s = ch.ChartTitle.Text'),
    ('numbers', None, 'xlColumnClustered', 'ch.HasTitle = True: s = ch.ChartTitle.Text'),
    ('one', None, 'xlPie', 'ch.HasTitle = True: s = ch.ChartTitle.Text'),
    ('two', None, 'xlColumnClustered', 'ch.Axes(2).ReversePlotOrder = True'),
    ('two', None, 'xlBarClustered', 'ch.Axes(2).ReversePlotOrder = True'),
    ('two', None, 'xlXYScatter', 'ch.Axes(1).ReversePlotOrder = True'),
    ('two', None, 'xlXYScatter', 'ch.Axes(2).ReversePlotOrder = True'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    On Error GoTo Failed',
             '    Select Case which']
    for index, (data, plot_by, kind, statements) in enumerate(CASES):
        lines.append(f'    Case {index}')
        if data is not None:
            columns = DATA[data]
            for offset, column in enumerate(columns):
                letter = 'ABC'[offset]
                values = ', '.join(f'"{item}"' if isinstance(item, str) else str(item) for item in column)
                lines.append(f'        ws.Range("{letter}1:{letter}{len(column)}").Value = '
                             f'Application.Transpose(Array({values}))')
            source = f'ws.Range("A1:{"ABC"[len(columns) - 1]}{len(columns[0])}")'
            lines.append(f'        ch.SetSourceData {source}' + (f', {plot_by}' if plot_by else ''))
        else:
            lines.append('        ws.Range("B2:D4").Value = 5')
        lines.append(f'        ch.ChartType = {kind}')
        if statements:
            lines.append(f'        {statements}')
    lines += ['    End Select', f'    Tried = "ok " & {READ} & "," & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def normalized(xml: str) -> str:
    """The chart part with its random ids made comparable."""
    numbers: dict[str, str] = {}

    def axis(match: re.Match[str]) -> str:
        return f'{match.group(1)} val="{numbers.setdefault(match.group(2), str(len(numbers) + 1))}"'

    xml = re.sub(r'(<c:(?:axId|crossAx)) val="(\d+)"', axis, xml)
    return re.sub(r'(<c16:uniqueId val="\{[0-9A-F]{8})-[0-9A-F-]{27}\}', r'\1}', xml)


def chart_part(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        names = [entry for entry in package.namelist() if entry.startswith('xl/charts/')]
        assert len(names) == 1 and re.fullmatch(r'xl/charts/chart\d+\.xml', names[0]), names
        return normalized(package.read(names[0]).decode('utf-8'))


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{data} {plot_by} {kind}: {statements}', answer, part]
                        for (data, plot_by, kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
