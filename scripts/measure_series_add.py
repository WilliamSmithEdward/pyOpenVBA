"""SeriesCollection.Add and Extend: what each answers, the series it leaves, and the chart part Excel saves.

Each case makes a workbook with categories in A2:A5 and series with names in B, C and D on Data, numbers with no name in F, a
sheet Other, and a chart on Data from the setup it names: empty, or of A1:B5. It runs the case's statements, which set r to
what they want recorded, reads READS, each as its value or "E" and an error number, and saves the workbook. The chart part is
recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_series_add.py

writes tests/fixtures/series_add.json, which tests/test_excel_series_add.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/series_add.json'
S = 'ch.SeriesCollection'
SETUPS = {'empty': '', 'one': 'ch.SetSourceData d.Range("A1:B5"), xlColumns'}
ADDS = [
    'r = TypeName({S}.Add(d.Range("C1:C5")))', 'r = {S}.Add(d.Range("C1:C5")).Name', '{S}.Add d.Range("C1:C5")', '{S}.Add d.Range("C2:C5")', '{S}.Add d.Range("C1:D5")', '{S}.Add d.Range("C2:D5")', '{S}.Add d.Range("A1:C5")', '{S}.Add d.Range("A2:C5")', '{S}.Add d.Range("F1:F4")', '{S}.Add d.Range("C3")', '{S}.Add d.Range("C1")',
    '{S}.Add d.Range("C1:C5"), xlColumns', '{S}.Add d.Range("C1:D5"), xlColumns', '{S}.Add d.Range("A1:C5"), xlRows', '{S}.Add d.Range("B2:D2"), xlRows', '{S}.Add d.Range("A2:D3"), xlRows', '{S}.Add d.Range("C1:C5"), 3', '{S}.Add d.Range("C1:C5"), 0',
    '{S}.Add d.Range("C1:C5"), xlColumns, True', '{S}.Add d.Range("C1:C5"), xlColumns, False', '{S}.Add d.Range("C2:C5"), xlColumns, True', '{S}.Add d.Range("C2:C5"), xlColumns, False', '{S}.Add d.Range("A1:C5"), xlColumns, True, True', '{S}.Add d.Range("A1:C5"), xlColumns, True, False', '{S}.Add d.Range("A1:C5"), xlColumns, False, True', '{S}.Add d.Range("A2:C5"), xlColumns, False, True', '{S}.Add d.Range("A2:C5"), xlColumns, False, False', '{S}.Add d.Range("C1:C5"), xlColumns, True, False, True', '{S}.Add d.Range("C1:C5"), xlColumns, True, False, False', '{S}.Add d.Range("A1:C5"), xlColumns, True, True, True', '{S}.Add d.Range("A1:C5"), xlColumns, True, True, False',
    '{S}.Add Source:=d.Range("C1:C5"), Rowcol:=xlColumns, SeriesLabels:=True', '{S}.Add "=Data!$C$1:$C$5"', '{S}.Add "Data!R1C3:R5C3"', '{S}.Add "nothing"', '{S}.Add 5', '{S}.Add', '{S}.Add o.Range("A1:A3")', '{S}.Add d.Range("C1:C5,D1:D5")', '{S}.Add d.Rows(1)', '{S}.Add d.Columns(3)', '{S}.Add d.Range("H1:H5")', '{S}.Add d.Range("A1:A5")', '{S}.Add d.Range("B1:B5")',
    '{S}.Add d.Range("C1:C5"): {S}.Add d.Range("D1:D5")', '{S}.Add d.Range("C1:D5"): r = {S}.Count & {S}(2).Name', 'ch.ChartType = xlLine: {S}.Add d.Range("C1:C5")', 'ch.ChartType = xlPie: {S}.Add d.Range("C1:C5")', 'ch.ChartType = xlXYScatter: {S}.Add d.Range("C1:C5")', 'ch.ChartType = xlXYScatter: {S}.Add d.Range("A1:C5"), xlColumns, True, True',
]
EXTENDS = [
    'r = {S}.Extend(d.Range("B6:B8"))', '{S}.Extend d.Range("B6:B8")', '{S}.Extend d.Range("C2:C5")', '{S}.Extend d.Range("C1:C5")', '{S}.Extend d.Range("A6:B8")', '{S}.Extend d.Range("A6:B8"), xlColumns', '{S}.Extend d.Range("A6:B8"), xlColumns, True', '{S}.Extend d.Range("A6:B8"), xlColumns, False', '{S}.Extend d.Range("B6:B8"), xlColumns, False', '{S}.Extend d.Range("B6:B8"), xlRows', '{S}.Extend d.Range("B6:C8")', '{S}.Extend d.Range("C6")', '{S}.Extend "=Data!$B$6:$B$8"', '{S}.Extend "nothing"', '{S}.Extend', '{S}.Extend Source:=d.Range("B6:B8")', '{S}.Extend o.Range("A1:A3")', 'ch.ChartType = xlPie: {S}.Extend d.Range("B6:B8")', 'ch.ChartType = xlLine: {S}.Extend d.Range("B6:B8")',
]
CASES = [(setup, statement.format(S=S)) for setup in SETUPS for statement in [*ADDS, *EXTENDS]]
READS = [f'{S}.Count', f'{S}(1).Formula', f'{S}(2).Formula', f'{S}(3).Formula', f'{S}(1).Name', f'{S}(2).Name', f'{S}(1).Points.Count', 'ch.PlotBy', 'ch.HasLegend']


def _probe() -> str:
    lines: list[str] = []
    for first in range(0, len(CASES), 100):
        lines += [f'Private Sub Act{first // 100}(ByVal which As Long, d As Object, o As Object, ch As Object, r As Variant)', '    Select Case which']
        for index, (setup, statement) in enumerate(CASES[first:first + 100], first):
            lines += [f'    Case {index}']
            lines += [f'        {SETUPS[setup]}'] if SETUPS[setup] else []
            lines += [f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
              '    Dim wb As Workbook, d As Object, o As Object, ch As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
              '    d.Range("A1:A8").Value = Application.Transpose(Array("k", "a", "b", "c", "d", "e", "f", "g"))', '    d.Range("B1:B8").Value = Application.Transpose(Array("v", 1, 2, 3, 4, 5, 6, 7))', '    d.Range("C1:C8").Value = Application.Transpose(Array("w", 11, 12, 13, 14, 15, 16, 17))', '    d.Range("D1:D8").Value = Application.Transpose(Array("x", 21, 22, 23, 24, 25, 26, 27))', '    d.Range("F1:F4").Value = Application.Transpose(Array(31, 32, 33, 34))', '    o.Range("A1:A3").Value = Application.Transpose(Array(41, 42, 43))',
              '    d.Range("J1").Select', '    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart', '    On Error GoTo Failed', '    Select Case which \\ 100']
    for block in range((len(CASES) + 99) // 100):
        lines += [f'    Case {block}: Act{block} which, d, o, ch, r']
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
              'cases': [[f'{setup}: {statement}', answer, part] for (setup, statement), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
