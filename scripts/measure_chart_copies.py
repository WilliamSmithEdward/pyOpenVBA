"""A chart duplicated, copied and pasted, cut, or copied with its sheet: what the new one reads, and every chart part Excel saves.

Each case makes a workbook with categories in A2:A4 and two series in B and C on Data, a sheet Other, and a column chart on
Data with a title, a series fill and a point's fill. It runs the case's statements, which set r to what they want recorded,
reads READS, each as its value or "E" and an error number, and saves the workbook. Each chart part is recorded as
scripts/measure_chart_markup.py records it, by its name.

    python scripts/measure_chart_copies.py

writes tests/fixtures/chart_copies.json, which tests/test_excel_chart_copies.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import normalized

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_copies.json'
ABOUT = 'n.Name & " " & n.Left & " " & n.Top & " " & n.Width & " " & n.Height & " " & n.Index & " " & n.ZOrder & " " & n.Chart.SeriesCollection.Count & " " & n.Chart.ChartTitle.Text & " " & n.Chart.SeriesCollection(1).Formula'
CASES = [
    # Duplicate.
    f'Set n = co.Duplicate: r = {ABOUT}', 'r = TypeName(co.Duplicate)', f'Set n = d.Shapes(1).Duplicate.DrawingObject: r = {ABOUT}', 'r = TypeName(d.Shapes(1).Duplicate)', f'co.Duplicate: Set n = co.Duplicate: r = {ABOUT}', f'Set n = co.Duplicate: Set n = n.Duplicate: r = {ABOUT}',
    f'co.Name = "Mine": Set n = co.Duplicate: r = {ABOUT}', f'co.Left = 10: co.Top = 20: Set n = co.Duplicate: r = {ABOUT}', 'Set n = co.Duplicate: n.Chart.ChartTitle.Text = "other": r = ch.ChartTitle.Text & " " & n.Chart.ChartTitle.Text', 'Set n = co.Duplicate: n.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = 65280: r = ch.SeriesCollection(1).Format.Fill.ForeColor.RGB', 'Set n = co.Duplicate: d.Rows(1).Insert: r = n.Chart.SeriesCollection(1).Formula', 'Set n = co.Duplicate: d.Name = "Moved": r = n.Chart.SeriesCollection(1).Formula', 'Set n = co.Duplicate: n.Delete: r = d.ChartObjects.Count', 'Set n = co.Duplicate: co.Delete: r = d.ChartObjects(1).Name',
    'r = TypeName(d.ChartObjects.Duplicate)', 'd.ChartObjects.Duplicate: r = d.ChartObjects.Count', 'd.ChartObjects.Add 5, 5, 50, 50: d.ChartObjects.Duplicate: r = d.ChartObjects.Count',
    'ch.SeriesCollection(1).Trendlines.Add: ch.SeriesCollection(1).HasDataLabels = True: ch.Axes(xlValue).HasTitle = True: ch.Axes(xlValue).AxisTitle.Text = "v": co.Duplicate: r = d.ChartObjects(2).Chart.SeriesCollection(1).Trendlines.Count', 'ch.Legend.Format.Fill.ForeColor.RGB = 255: ch.PlotArea.Format.Fill.ForeColor.RGB = 255: ch.ChartArea.Format.Fill.ForeColor.RGB = 255: co.Duplicate: r = d.ChartObjects(2).Chart.PlotArea.Format.Fill.ForeColor.RGB',
    # Copy and paste.
    'r = co.Copy', 'r = d.Shapes(1).Copy', 'co.Copy: d.Paste: r = d.ChartObjects.Count', f'co.Copy: d.Range("H20").Select: d.Paste: Set n = d.ChartObjects(2): r = {ABOUT}', f'co.Copy: d.Range("H20").Select: d.Paste d.Range("A10"): Set n = d.ChartObjects(2): r = {ABOUT}', f'co.Copy: d.Paste Destination:=d.Range("A10"): Set n = d.ChartObjects(2): r = {ABOUT}', f'd.Shapes(1).Copy: d.Paste d.Range("A10"): Set n = d.ChartObjects(2): r = {ABOUT}',
    f'co.Copy: o.Paste o.Range("B2"): Set n = o.ChartObjects(1): r = {ABOUT} & " " & d.ChartObjects.Count & " " & ActiveSheet.Name', f'co.Copy: o.Activate: o.Range("B2").Select: o.Paste: Set n = o.ChartObjects(1): r = {ABOUT}', 'co.Copy: r = TypeName(d.Paste(d.Range("A10")))', 'co.Copy: d.Range("A10").PasteSpecial: r = d.ChartObjects.Count', 'co.Copy: d.Range("A10").Select: d.PasteSpecial: r = d.ChartObjects.Count', 'co.Copy: d.Paste d.Range("A10"): r = TypeName(Selection)', 'co.Copy: d.Paste d.Range("A10"): d.Paste d.Range("A30"): r = d.ChartObjects.Count', 'co.Copy: co.Delete: d.Paste d.Range("A10"): r = d.ChartObjects.Count & d.ChartObjects(1).Name', 'r = Application.CutCopyMode: co.Copy: r = r & " " & Application.CutCopyMode', 'ch.ChartArea.Copy: d.Paste d.Range("A10"): r = d.ChartObjects.Count', 'ch.Copy',
    # Cut.
    'r = d.Shapes(1).Cut', f'd.Shapes(1).Cut: r = d.ChartObjects.Count', 'd.Shapes(1).Cut: d.Paste d.Range("A10"): r = d.ChartObjects.Count & d.ChartObjects(1).Name & d.ChartObjects(1).Left', 'd.Shapes(1).Cut: o.Paste o.Range("A10"): r = d.ChartObjects.Count & o.ChartObjects.Count & o.ChartObjects(1).Name', 'co.Cut: r = d.ChartObjects.Count', 'co.Cut: d.Paste d.Range("A10"): r = d.ChartObjects.Count',
    # With its sheet.
    'd.Copy After:=d: Set n = wb.Worksheets(2).ChartObjects(1): r = n.Name & " " & n.Chart.SeriesCollection(1).Formula & " " & wb.Worksheets(2).Name', 'd.Copy After:=d: r = wb.Worksheets(2).ChartObjects(1).Chart.ChartTitle.Text', 'o.Range("A1").Value = 1: d.Copy Before:=o: r = wb.Worksheets(2).Name & wb.Worksheets(2).ChartObjects.Count', 'd.Copy After:=o: wb.Worksheets(3).Rows(1).Insert: r = wb.Worksheets(3).ChartObjects(1).Chart.SeriesCollection(1).Formula & " " & ch.SeriesCollection(1).Formula', 'd.Copy After:=d: d.Range("B2").Value = 9: r = wb.Worksheets(2).ChartObjects(1).Chart.SeriesCollection(1).Values()(1)', 'o.ChartObjects.Add 1, 1, 100, 100: o.ChartObjects(1).Chart.SetSourceData d.Range("A1:B4"): o.Copy After:=o: r = wb.Worksheets(3).ChartObjects(1).Chart.SeriesCollection(1).Formula', 'd.Copy: r = ActiveWorkbook.Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Formula: ActiveWorkbook.Close False',
]
# What a pasted chart's unique ids follow.
S = 'ch.SeriesCollection'
PASTE = 'co.Copy: d.Paste d.Range("A10")'
CASES += [f'ch.ClearToMatchStyle: {look}{PASTE}' for look in (
    '', f'{S}(1).Format.Fill.ForeColor.RGB = 255: ', f'{S}(2).Format.Fill.ForeColor.RGB = 255: ', f'{S}(2).Points(2).Format.Fill.ForeColor.RGB = 255: ', f'{S}(1).Points(1).Format.Fill.ForeColor.RGB = 255: ',
    f'{S}(1).Points(1).Format.Fill.ForeColor.RGB = 255: {S}(1).Points(2).Format.Fill.ForeColor.RGB = 255: {S}(2).Points(1).Format.Fill.ForeColor.RGB = 255: ', f'{S}(1).Points(2).HasDataLabel = True: ', f'{S}(1).Points(2).Format.Fill.ForeColor.RGB = 255: {S}(1).Points(2).HasDataLabel = True: ',
    'd.Range("D1:D4").Value = 1: ch.SetSourceData d.Range("A1:D4"), xlColumns: ', f'{S}.NewSeries: ', f'{S}(1).Delete: ')]
READS = ['d.ChartObjects.Count', 'o.ChartObjects.Count', 'wb.Worksheets.Count', 'ActiveSheet.Name', 'TypeName(Selection)', 'ch.ChartTitle.Text', 'ch.SeriesCollection.Count', 'd.Shapes.Count']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, d As Object, o As Object, co As Object, ch As Object, n As Object, s As String, v As String, r As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set d = wb.Worksheets(1): d.Name = "Data"', '    Set o = wb.Worksheets.Add(After:=d): o.Name = "Other": d.Activate',
             '    d.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))', '    d.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))', '    d.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 5, 6))',
             '    Set co = d.ChartObjects.Add(300, 100, 300, 200): Set ch = co.Chart', '    ch.SetSourceData d.Range("A1:C4"), xlColumns', '    ch.HasTitle = True: ch.ChartTitle.Text = "t"', '    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 255', '    ch.SeriesCollection(2).Points(2).Format.Fill.ForeColor.RGB = 65280', '    d.Range("A1").Select', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    Application.CutCopyMode = False', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False', '    Exit Function',
              'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_parts(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: normalized(package.read(name).decode('utf-8')) for name in package.namelist()
                if re.fullmatch(r'xl/charts/chart\d+\.xml', name)}


def main() -> None:
    """Each case alone, since one that hangs Excel takes its session with it: such a case is recorded as "hangs"."""
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder:
        index = 0
        while index < len(CASES):
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
                excel.new_document()
                while index < len(CASES):
                    result = excel.run_vba(source, 'Probe', (folder, index, index), timeout=45.0)
                    index += 1
                    if not result.ok or result.dialogs:
                        print('hangs', CASES[index - 1], result.outcome, flush=True)
                        answers.append('hangs')
                        break
                    answers.append(str(result.value))
        parts = [chart_parts(Path(folder) / f'c{place}.xlsx') if answers[place] != 'hangs' else {}
                 for place in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        print(case, '=>', answer, sorted(part))


if __name__ == '__main__':
    main()
