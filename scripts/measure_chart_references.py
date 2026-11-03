"""A chart's series references when the cells they read are moved: rows and columns inserted and deleted, cells shifted and
cut, and the sheet renamed or removed. What the series reads afterwards, and the chart part Excel saves.

Each case makes a workbook with a sheet Data holding categories in A2:A4 and two series in B and C, a sheet Other, and a
chart on Data or Other. It runs the case's statements, reads READS, each as its value or "E" and an error number, and saves
the workbook. The chart part is recorded as scripts/measure_chart_markup.py records it.

    python scripts/measure_chart_references.py

writes tests/fixtures/chart_references.json, which tests/test_excel_chart_references.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_references.json'
EDITS = [
    '', 'd.Rows(1).Insert', 'd.Rows(3).Insert', 'd.Rows(4).Insert', 'd.Rows(5).Insert', 'd.Rows(9).Insert', 'd.Rows("2:3").Insert', 'd.Rows(1).Delete', 'd.Rows(2).Delete', 'd.Rows(3).Delete', 'd.Rows(4).Delete', 'd.Rows(9).Delete', 'd.Rows("2:3").Delete', 'd.Rows("2:4").Delete',
    'd.Rows("1:4").Delete', 'd.Columns(1).Insert', 'd.Columns(2).Insert', 'd.Columns(3).Insert', 'd.Columns(4).Insert', 'd.Columns(1).Delete', 'd.Columns(2).Delete', 'd.Columns(3).Delete', 'd.Columns("B:C").Delete', 'd.Columns(5).Delete',
    'd.Range("B3").Insert xlShiftDown', 'd.Range("B3").Delete xlShiftUp', 'd.Range("B3:C3").Insert xlShiftDown', 'd.Range("B3:C3").Delete xlShiftUp', 'd.Range("A2:C4").Insert xlShiftToRight', 'd.Range("B2:B4").Insert xlShiftToRight', 'd.Range("B2:B4").Delete xlShiftToLeft', 'd.Range("B1:B4").Insert xlShiftDown',
    'd.Range("B1:B4").Delete xlShiftUp', 'd.Range("A1:C4").Cut d.Range("E6")', 'd.Range("B1:B4").Cut d.Range("E6")', 'd.Range("B2:B4").Cut d.Range("E6")', 'd.Range("B3").Cut d.Range("E6")', 'd.Range("A1:C4").Copy d.Range("E6")', 'd.Range("B2:B4").Cut o.Range("A1")', 'd.Range("A1:C4").Clear', 'd.Range("A1:C4").ClearContents',
    'd.Name = "Moved"', 'd.Name = "With Space"', 'd.Copy After:=d', 'o.Delete', 'd.Move After:=o', 'd.Rows(1).Insert: d.Rows(1).Insert', 'd.Rows(1).Insert: d.Rows(1).Delete', 'd.Rows(1).Insert: d.Columns(1).Insert', 'd.Range("B2").Value = 9', 'd.Rows(3).Insert: d.Range("B3").Value = 9',
    'd.Rows(1).Hidden = True', 'd.Rows(2).Hidden = True', 'ch.SeriesCollection(1).Name = "=Data!$A$1": d.Rows(1).Insert', 'ch.SeriesCollection(1).Name = "z": d.Rows(1).Insert', 'ch.SeriesCollection(1).Values = Array(7, 8, 9): d.Rows(1).Insert', 'ch.SeriesCollection(1).XValues = Array("x", "y", "z"): d.Rows(1).Insert',
    'ch.SeriesCollection(1).Formula = "=SERIES(Data!$A$1,Data!$A$2:$A$4,Data!$C$2:$C$4,1)": d.Rows(1).Insert', 'ch.SetSourceData d.Range("A1:B4"): d.Rows(1).Insert', 'ch.SeriesCollection(1).Delete: d.Rows(1).Insert', 'ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255: d.Rows(1).Delete', 'ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255: d.Rows(3).Delete',
    'ch.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = 255: d.Rows(3).Delete', 'ch.SeriesCollection(1).Points(2).HasDataLabel = True: d.Rows(2).Delete', 'ch.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = 255: d.Rows("2:3").Delete', 'ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255: d.Rows(3).Insert', 'ch.SeriesCollection(1).ErrorBar xlY, xlBoth, xlCustom, d.Range("C2:C4"), d.Range("C2:C4"): d.Rows(1).Insert',
    'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$A$1": d.Rows(1).Insert', 'ch.HasTitle = True: ch.ChartTitle.Formula = "=Data!$A$1": d.Rows(1).Delete', 'ch.SeriesCollection(1).HasDataLabels = True: d.Rows(1).Insert',
]
ON_OTHER = ['', 'd.Rows(1).Insert', 'd.Columns(1).Delete', 'd.Rows("1:4").Delete', 'd.Name = "Moved"', 'd.Delete', 'o.Rows(1).Insert', 'o.Columns(1).Insert', 'o.Name = "Moved"']
CASES = [('data', edit) for edit in EDITS] + [('other', edit) for edit in ON_OTHER] + [('line', edit) for edit in ('d.Rows(1).Insert', 'd.Rows(3).Delete', 'd.Columns(2).Delete')] + [('scatter', edit) for edit in ('d.Rows(1).Insert', 'd.Rows(3).Delete', 'd.Columns(1).Delete')] + [('pie', edit) for edit in ('d.Rows(1).Insert', 'd.Rows(3).Delete')]
S1, S2 = 'ch.SeriesCollection(1)', 'ch.SeriesCollection(2)'
READS = ['ch.SeriesCollection.Count', f'{S1}.Formula', f'{S2}.Formula', f'{S1}.Name', f'{S1}.Points.Count', f'{S2}.Points.Count', 'ch.PlotBy', 'ch.HasTitle', 'ch.ChartTitle.Text', 'ch.ChartTitle.Formula', 'ch.Parent.TopLeftCell.Address', 'ActiveSheet.Name']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, d As Worksheet, o As Worksheet, ch As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"',
             '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate', '    d.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    d.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    d.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))', '    On Error GoTo Failed', '    Select Case which']
    types = {'data': 'xlColumnClustered', 'other': 'xlColumnClustered', 'line': 'xlLine', 'scatter': 'xlXYScatter', 'pie': 'xlPie'}
    for index, (kind, edit) in enumerate(CASES):
        place = 'o' if kind == 'other' else 'd'
        lines += [f'    Case {index}', f'        Set ch = {place}.ChartObjects.Add(300, 100, 300, 200).Chart', '        ch.SetSourceData d.Range("A1:C4"), xlColumns', f'        ch.ChartType = {types[kind]}']
        lines += [f'        {edit}'] if edit else []
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def _part(path: Path) -> str:
    """The chart part, or nothing where the case took the chart away with its sheet or copied it."""
    try:
        return chart_part(path)
    except (AssertionError, KeyError):
        return ''


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
        parts = [_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[f'{kind}: {edit}', answer, part] for (kind, edit), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case.replace(S1, 'S1'), '=>', answer)


if __name__ == '__main__':
    main()
