"""The VBA type of each Font read on a chart's elements, measured live. Run on Windows with Excel:

    python scripts/measure_chart_font_types.py

writes tests/fixtures/chart_font_types.json, which tests/test_excel_chart_font_gaps.py replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_font_types.json'
OWNERS = ('ch.ChartArea.Font', 'ch.Legend.Font', 'ch.ChartTitle.Font', 'ch.Axes(1).TickLabels.Font', 'ch.SeriesCollection(1).DataLabels.Font')
MEMBERS = ('Size', 'Bold', 'Italic', 'Underline', 'Strikethrough', 'Subscript', 'Superscript', 'Color', 'ColorIndex', 'Name', 'FontStyle',
           'Background', 'OutlineFont', 'Shadow')


def probe() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")', '    ws.Range("A2:C2").Value = Array("a", 10, 4)',
             '    ws.Range("A3:C3").Value = Array("b", 20, 6)', '    Set ch = ws.ChartObjects.Add(48, 96, 240, 144).Chart',
             '    ch.SetSourceData ws.Range("A1:C3")', '    ch.HasTitle = True', '    ch.SeriesCollection(1).HasDataLabels = True']
    lines += [f'    s = s & "{owner}.{member}=" & TypeName({owner}.{member}) & "|"' for owner in OWNERS for member in MEMBERS]
    lines += ['    wb.Close False', '    Probe = s', 'End Function']
    return '\n'.join(lines)


def main() -> None:
    source = probe()
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        done = excel.run_vba(source, 'Probe', (), timeout=120.0, line_numbers=False)
    assert done.outcome == 'passed', done.message
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': source, 'types': done.value}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    print(done.value)


if __name__ == '__main__':
    main()
