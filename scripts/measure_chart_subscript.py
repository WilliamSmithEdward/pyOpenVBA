"""What Font.Subscript does on a chart's text, which is not the same from run to run: each step on each element, ten times over five
fresh sessions, tallied by what Subscript and Superscript read and the baseline Excel saved.

    python scripts/measure_chart_subscript.py

writes tests/fixtures/chart_subscript.json, which tests/test_excel_chart_font_gaps.py reads.
"""
from __future__ import annotations

import collections
import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_subscript.json'
ELEMENTS = {'title': 'ch.ChartTitle.Font', 'legend': 'ch.Legend.Font', 'ticks': 'ch.Axes(2).TickLabels.Font', 'labels': 's.DataLabels.Font', 'area': 'ch.ChartArea.Font', 'axis title': 'ch.Axes(2).AxisTitle.Font'}
STEPS = {'on': '{0}.Subscript = True', 'off': '{0}.Subscript = False', 'on off': '{0}.Subscript = True: {0}.Subscript = False', 'superscript': '{0}.Superscript = True'}
MAKE = '''    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:D1").Value = Array("", "Rev", "Cost", "Rate")
    ws.Range("A2:D2").Value = Array("a", 10, 4, 1)
    ws.Range("A3:D3").Value = Array("b", 20, 6, 2)
    Set ch = ws.ChartObjects.Add(48, 48, 360, 216).Chart
    ch.SetSourceData ws.Range("A1:D3")
    ch.ChartType = xlColumnClustered
    ch.HasTitle = True
    ch.ChartTitle.Text = "Mine"
    ch.HasLegend = True
    ch.Axes(2).HasTitle = True
    Set s = ch.SeriesCollection(1)
    s.HasDataLabels = True
'''


def main() -> None:
    lines = ['Public Function Probe(ByVal folder As String) As String', '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As Object, t As String', '    Application.DisplayAlerts = False']
    labels: list[str] = []
    for element, font in ELEMENTS.items():
        for step, statement in STEPS.items():
            labels.append(f'{element}: {step}')
            lines += [MAKE.rstrip('\n'), '    ' + statement.format(font), f'    t = t & {font}.Subscript & "/" & {font}.Superscript & ";"', f'    wb.SaveAs folder & "\\p{len(labels)}.xlsx", xlOpenXMLWorkbook', '    wb.Close False']
    lines += ['    Application.DisplayAlerts = True', '    Probe = t', 'End Function']
    source = '\n'.join(lines) + '\n'
    tally: dict[str, collections.Counter[str]] = {label: collections.Counter() for label in labels}
    for _session in range(5):
        with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
            excel.new_document()
            for _ in range(2):
                done = excel.run_vba(source, 'Probe', (folder,), timeout=200.0, line_numbers=False)
                assert done.ok, done
                for index, (label, answer) in enumerate(zip(labels, str(done.value).split(';')[:-1], strict=True), start=1):
                    with zipfile.ZipFile(Path(folder) / f'p{index}.xlsx') as package:
                        part = package.read('xl/charts/chart1.xml').decode('utf-8')
                    tally[label][f"{answer} {','.join(sorted(set(re.findall(r'baseline=.(-?[0-9]+).', part))))}"] += 1
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'tally': {label: dict(counts) for label, counts in tally.items()}}, indent=1) + '\n', encoding='utf-8')
    for label, counts in tally.items():
        print(label, dict(counts))


if __name__ == '__main__':
    main()
