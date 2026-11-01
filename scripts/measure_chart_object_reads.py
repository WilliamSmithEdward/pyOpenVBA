"""ChartObject, Chart, ChartGroup, ChartTitle and Legend members that read simply: what each reads on a new chart and
after RoundedCorners or Shadow is set, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add at the case's place, fills categories in A2:A4 and
two series in B and C, gives them to the chart with SetSourceData, gives it a title, runs the case's statements,
then reads every expression of READS, each as its value or "E" and the error number, and saves it. The chart part is
recorded as scripts/measure_chart_markup.py records it, and the drawing part's graphic frame beside it.

    python scripts/measure_chart_object_reads.py

writes tests/fixtures/chart_object_reads.json, which tests/test_excel_chart_object_reads.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_object_reads.json'
CASES = [
    ('100, 10, 300, 200', ''),
    ('0, 0, 48, 15', ''),
    ('48.5, 15.5, 47, 14', ''),
    ('100, 10, 300, 200', 'co.RoundedCorners = True'),
    ('100, 10, 300, 200', 'co.RoundedCorners = True: co.RoundedCorners = False'),
    ('100, 10, 300, 200', 'co.Shadow = True'),
    ('100, 10, 300, 200', 'co.Shadow = True: co.Shadow = False'),
    ('100, 10, 300, 200', 'ws.Columns("A:D").ColumnWidth = 20'),
]
READS = ['TypeName(co.Application)', 'co.TopLeftCell.Address', 'co.BottomRightCell.Address', 'co.RoundedCorners',
         'co.Shadow', 'TypeName(ch.Application)', 'TypeName(ch.ChartGroups(1).Application)',
         'ch.ChartGroups(1).AxisGroup', 'ch.ChartGroups(1).Index', 'ch.ChartGroups(1).SeriesCollection.Count',
         'ch.ChartGroups(1).SeriesCollection(2).Name', 'TypeName(ch.ChartTitle.Application)',
         'TypeName(ch.ChartTitle.Parent)', 'ch.ChartTitle.Name', 'ch.Legend.Name', 'TypeName(ch.Legend.Application)']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, co As Object, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Select Case which']
    for index, (place, _) in enumerate(CASES):
        lines += [f'    Case {index}', f'        Set co = ws.ChartObjects.Add({place})']
    lines += ['    End Select', '    Set ch = co.Chart',
              '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
              '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
              '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
              '    ch.SetSourceData ws.Range("A1:C4")', '    ch.HasTitle = True', '    On Error GoTo Failed',
              '    Select Case which']
    for index, (_, statements) in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {statements}'] if statements else [])
    lines += ['    End Select', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', f'    s = s & "," & {read}', '    If Err.Number Then s = s & "E" & Err.Number']
    lines += ['    On Error GoTo Failed', '    Tried = "ok " & Mid(s, 2)', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & s', '    Resume Saved',
              'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def frame(path: Path) -> str:
    """The drawing's graphic frame properties, which hold a chart object's name and look."""
    with zipfile.ZipFile(path) as book:
        drawing = book.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.search(r'<xdr:graphicFrame.*?</xdr:nvGraphicFramePr>', drawing)
    return found.group(0) if found else ''


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
        frames = [frame(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'reads': READS,
              'cases': [[f'{place}: {statements}', answer, part, drawn]
                        for (place, statements), answer, part, drawn in zip(CASES, answers, parts, frames,
                                                                            strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part, drawn in record['cases']:
        head = part[:part.find('<c:chart>')]
        print(case, '->', answer, '||', head[-200:], '||', drawn[-150:])


if __name__ == '__main__':
    main()
