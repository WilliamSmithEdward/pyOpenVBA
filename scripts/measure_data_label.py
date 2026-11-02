"""One point's DataLabel: what it reads, what each change does, and the chart part Excel saves.

Each case makes a workbook with a chart of categories in A2:A4 and two series in B and C, gives a point a label one of
the ways SETUPS names, runs the change on el, the second point's DataLabel, reads READS, each as its value or "E" and an
error number, and saves the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_data_label.py

writes tests/fixtures/data_label.json, which tests/test_excel_data_label.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/data_label.json'
SETUPS = {
    'series': 'ch.SeriesCollection(1).HasDataLabels = True',
    'point': 'ch.SeriesCollection(1).Points(2).HasDataLabel = True',
    'pie': 'ch.ChartType = xlPie: ch.SeriesCollection(1).HasDataLabels = True',
}
CHANGES = [
    '', 'el.ShowValue = False', 'el.ShowValue = True', 'el.ShowCategoryName = True', 'el.ShowSeriesName = True', 'el.ShowLegendKey = True',
    'el.ShowPercentage = True', 'el.ShowBubbleSize = True', 'el.ShowRange = True', 'el.ShowCategoryName = True: el.ShowValue = False',
    'el.ShowValue = False: el.ShowCategoryName = False', 'el.Text = "typed"', 'el.Caption = "typed"', 'el.Text = ""', 'el.Text = "typed": el.AutoText = True',
    'el.AutoText = False', 'el.AutoText = True', 'el.Characters.Text = "typed"', 'el.Formula = "=Sheet1!$A$1"', 'el.Formula = "typed"', 'el.FormulaR1C1 = "=Sheet1!R1C1"',
    'el.Position = xlLabelPositionInsideEnd', 'el.Position = xlLabelPositionCenter', 'el.Position = xlLabelPositionInsideBase', 'el.Position = xlLabelPositionOutsideEnd',
    'el.Position = xlLabelPositionAbove', 'el.Position = xlLabelPositionBestFit', 'el.Position = xlLabelPositionCustom', 'el.Position = 99',
    'el.NumberFormat = "0.0"', 'el.NumberFormat = "General"', 'el.NumberFormat = "nope"', 'el.NumberFormatLinked = False', 'el.NumberFormatLinked = True',
    'el.NumberFormat = "0.0": el.NumberFormatLinked = True', 'el.NumberFormatLocal = "0.00"', 'el.Separator = "; "', 'el.Separator = xlDataLabelSeparatorDefault', 'el.Separator = ""',
    'el.Orientation = 45', 'el.Orientation = xlUpward', 'el.Orientation = 100', 'el.HorizontalAlignment = xlLeft', 'el.VerticalAlignment = xlTop', 'el.ReadingOrder = xlRTL',
    'el.Shadow = True', 'el.Font.Size = 14', 'el.Font.Bold = True', 'el.Font.Color = 255', 'el.Format.Fill.ForeColor.RGB = 255', 'el.Format.Line.ForeColor.RGB = 255',
    'el.Delete', 'el.Delete: Set el = ch.SeriesCollection(1).Points(2).DataLabel', 'el.Select', 'el.Left = 50', 'el.Top = 50', 'el.Width = 50', 'el.Height = 50',
    'ch.SeriesCollection(1).Points(1).DataLabel.Text = "other"', 'ch.SeriesCollection(1).DataLabels.ShowCategoryName = True',
    'el.Text = "typed": ch.SeriesCollection(1).DataLabels.ShowCategoryName = True', 'el.ShowCategoryName = True: ch.SeriesCollection(1).DataLabels.ShowValue = False',
    'ch.SeriesCollection(1).DataLabels.NumberFormat = "0.00"', 'ch.SeriesCollection(1).DataLabels.Position = xlLabelPositionCenter',
    'ch.SeriesCollection(1).DataLabels.Font.Size = 14', 'ch.SeriesCollection(1).HasDataLabels = False', 'ch.SeriesCollection(1).Points(2).HasDataLabel = False',
    'el.ShowCategoryName = True: ch.SeriesCollection(1).Points(2).HasDataLabel = False: ch.SeriesCollection(1).Points(2).HasDataLabel = True: Set el = ch.SeriesCollection(1).Points(2).DataLabel',
    'Range("B2").Value = 9', 'Range("B3").Value = 9', 'Range("A3").Value = "z"',
]
CASES = [(setup, change) for setup in SETUPS for change in CHANGES]
READS = [
    'TypeName(el)', 'TypeName(el.Parent)', 'TypeName(el.Application)', 'el.Creator', 'el.Name', 'el.Text', 'el.Caption', 'el.AutoText', 'el.Formula', 'el.FormulaLocal', 'el.FormulaR1C1',
    'el.FormulaR1C1Local', 'el.ShowValue', 'el.ShowCategoryName', 'el.ShowSeriesName', 'el.ShowLegendKey', 'el.ShowPercentage', 'el.ShowBubbleSize', 'el.ShowRange', 'el.Position',
    'el.NumberFormat', 'el.NumberFormatLinked', 'el.NumberFormatLocal', 'el.Separator', 'el.Orientation', 'el.HorizontalAlignment', 'el.VerticalAlignment', 'el.ReadingOrder', 'el.Shadow',
    'TypeName(el.Characters)', 'el.Characters.Text', 'el.Characters.Count', 'TypeName(el.Font)', 'el.Font.Size', 'el.Font.Bold', 'el.Font.Name', 'el.Font.Color', 'TypeName(el.Format)',
    'el.Format.Fill.Visible', 'el.Format.Line.Visible', 'TypeName(el.Border)', 'TypeName(el.Interior)', 'TypeName(el.Fill)', 'ch.SeriesCollection(1).HasDataLabels',
    'ch.SeriesCollection(1).Points(2).HasDataLabel', 'ch.SeriesCollection(1).Points(1).HasDataLabel', 'ch.SeriesCollection(1).Points(1).DataLabel.Text', 'ch.SeriesCollection(1).DataLabels.Count',
    'ch.SeriesCollection(1).DataLabels(2).Text', 'ch.SeriesCollection(1).DataLabels.ShowValue',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, el As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "Sheet1"',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    ch.ChartType = xlColumnClustered', '    On Error GoTo Failed', '    Select Case which']
    for index, (setup, change) in enumerate(CASES):
        lines += [f'    Case {index}', f'        {SETUPS[setup]}', '        Set el = ch.SeriesCollection(1).Points(2).DataLabel']
        lines += [f'        {change}'] if change else []
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
              'cases': [[f'{setup}: {change}', answer, part] for (setup, change), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        found = re.findall(r'<c:dLbls>.*?</c:dLbls>', part, re.DOTALL)
        print(case, '\n  =', answer, '\n  +', found[0] if found else '')


if __name__ == '__main__':
    main()
