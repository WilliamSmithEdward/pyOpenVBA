"""The idx and unique id of each series in the chart part Excel saves, when cells already hold data as
ChartObjects.Add runs, and as SetSourceData, NewSeries and Delete change the series afterwards.

ChartObjects.Add plots nothing (SeriesCollection.Count reads 0), but Excel numbers the unique ids of later series past
those it would have plotted from the selection, or from the active cell's current region. Each case fills cells, runs
the statements before the chart is added, adds it, runs the statements after, and saves.

    python scripts/measure_chart_series_ids.py

writes tests/fixtures/chart_series_ids.json, which tests/test_excel_chart_series_ids.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_series_ids.json'
TEXT = ('ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c")): '
        'ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3)): '
        'ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))')
SQUARE = 'ws.Range("A1:D4").Value = 1: ws.Range("A1:D1").Value = Array("k", "v", "w", "x")'
WIDE = 'ws.Range("A1:E3").Value = 2: ws.Range("B1:E1").Value = Array("p", "q", "r", "s")'
#: (statements before ChartObjects.Add, statements after it)
CASES = [
    (TEXT, 'ch.SetSourceData ws.Range("A1:C4")'),
    (TEXT + ': ws.Range("F10").Select', 'ch.SetSourceData ws.Range("A1:C4")'),
    (TEXT + ': ws.Range("B3").Select', 'ch.SetSourceData ws.Range("A1:B4")'),
    (TEXT + ': ws.Range("A1:B4").Select', 'ch.SetSourceData ws.Range("A1:C4")'),
    (TEXT + ': ws.Range("B1:C4").Select', 'ch.SetSourceData ws.Range("A1:C4")'),
    (SQUARE, 'ch.SetSourceData ws.Range("A1:B4")'),
    (SQUARE + ': ws.Range("A1:B4").Select', 'ch.SetSourceData ws.Range("A1:B4")'),
    (WIDE, 'ch.SetSourceData ws.Range("A1:B3")'),
    ('ws.Range("C3").Value = 9', 'ch.SetSourceData ws.Range("C3")'),
    ('ws.Range("C3").Value = 9: ws.Range("C3").Select', 'ch.SetSourceData ws.Range("C3")'),
    ('', TEXT + ': ch.SetSourceData ws.Range("A1:C4"): ch.SetSourceData ws.Range("A1:C4")'),
    ('', TEXT + ': ch.SetSourceData ws.Range("A1:C4"): ch.SetSourceData ws.Range("A1:B4")'),
    ('', TEXT + ': ch.SetSourceData ws.Range("A1:C4"): ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4"'),
    ('', TEXT + ': ch.SetSourceData ws.Range("A1:C4"): ch.SeriesCollection(1).Delete: '
         'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4"'),
    (TEXT, 'ch.SeriesCollection.NewSeries.Values = "=Sheet1!$B$2:$B$4"'),
    (WIDE + ': ws.Range("H9").Select',
     'ch.SetSourceData ws.Range("A1:E3"): Tried = Tried & " " & ch.SeriesCollection(1).Formula'),
    ('ws.Range("H9").Select',
     TEXT + ': ws.Range("A1").Value = 7: ch.SetSourceData ws.Range("A1:C4"): '
     'Tried = Tried & " " & ch.SeriesCollection(1).Formula'),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object', '    Set wb = Workbooks.Add(xlWBATWorksheet)',
             '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed', '    Select Case which']
    for index, (before, _) in enumerate(CASES):
        lines += [f'    Case {index}'] + ([f'        {before}'] if before else [])
    lines += ['    End Select', '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart',
              '    Tried = "ok " & ch.SeriesCollection.Count', '    Select Case which']
    for index, (_, after) in enumerate(CASES):
        lines += [f'    Case {index}', f'        {after}']
    lines += ['    End Select', '    Tried = Tried & " " & ch.SeriesCollection.Count', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def series_ids(path: Path) -> list[list[int]]:
    """Each series' [idx, order, unique id's first group]."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/charts/chart1.xml').decode('utf-8')
    return [[int(idx), int(order), int(unique, 16)] for idx, order, unique in re.findall(
        r'<c:ser><c:idx val="(\d+)"/><c:order val="(\d+)"/>.*?<c16:uniqueId val="\{([0-9A-F]{8})-', xml)]


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        ids = [series_ids(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{before} || {after}', answer, found]
                        for (before, after), answer, found in zip(CASES, answers, ids, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-90:], '->', answer, found)


if __name__ == '__main__':
    main()
