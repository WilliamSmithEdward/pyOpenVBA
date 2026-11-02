"""Core members everyday macros read, as Excel answers them: each expression of READS as its text, or "E" and the
error number, on a new workbook set up as the probe sets it.

The workbook has 1 to 4 in A1:A4, a table over C1:D3, a rectangle with text, a chart, a note on A1, a sort field on
A1:A4, a three-colour scale and a unique-values rule on A1:A4, and a workbook name Data for A1:A4. Reads that
depend on the machine or the harness's own session, the user's name, paths, the build, the harness's workbooks and
EnableEvents, are left out.

    python scripts/measure_core_reads.py

writes tests/fixtures/core_reads.json, which tests/test_excel_core_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/core_reads.json'
READS = [
    'TypeName(Application.ActiveCell)', 'Application.ActiveCell.Address', 'TypeName(Application.ActiveSheet)',
    'Application.ActiveSheet.Name', 'Application.ActiveWorkbook.Name = wb.Name', 'TypeName(Application.Application)',
    'Application.Cells.Count', 'Application.Columns.Count', 'Application.Rows.Count', 'Application.DisplayAlerts',
    'Application.Names.Count', 'TypeName(Application.Parent)', 'TypeName(Application.Selection)',
    'Application.Selection.Address', 'Application.Sheets.Count', 'Application.Worksheets.Count',
    'Application.Version',
    'ws.Range("A1:B2,D4").Areas.Count', 'ws.Range("A1:B2,D4").Areas(2).Address',
    'ws.Range("A1:B2,D4").Areas.Item(1).Address',
    'TypeName(ws.Range("A1").Borders(xlEdgeTop).Parent)', 'TypeName(ws.Range("A1").Borders(xlEdgeTop).Application)',
    'TypeName(ws.Range("A1").Borders.Parent)', 'TypeName(ws.Range("A1").Borders.Application)',
    'TypeName(ws.Range("A1").Font.Parent)', 'TypeName(ws.Range("A1").Font.Application)',
    'TypeName(ws.Range("A1").Interior.Parent)', 'TypeName(ws.Range("A1").Interior.Application)',
    'TypeName(ws.Range("A1").Comment.Application)', 'TypeName(ws.Comments.Application)',
    'TypeName(ws.Comments.Parent)',
    'TypeName(cs.Parent)', 'cs.Priority', 'uv.Priority',
    'TypeName(cs.ColorScaleCriteria(1).FormatColor.Parent)',
    'TypeName(cs.ColorScaleCriteria(1).FormatColor.Application)',
    'TypeName(ch.Legend.Application)', 'ch.Legend.Creator', 'TypeName(ch.Legend.Parent)',
    'lo.ShowTableStyleColumnStripes', 'lo.ShowTableStyleFirstColumn', 'lo.ShowTableStyleLastColumn',
    'lo.ShowTableStyleRowStripes',
    'wb.Names("Data").NameLocal', 'wb.Names("Data").RefersToLocal', 'wb.Names("Data").RefersToR1C1Local',
    'wb.Names("Data").RefersToRange.Address',
    'TypeName(ws.Range("A1").Application)', 'TypeName(ws.Range("A1").Parent)',
    'TypeName(sh.Chart)', 'sh.HasChart', 'co.ShapeRange(1).HasChart', 'TypeName(co.ShapeRange(1).Chart)',
    'sh.TextFrame2.HasText', 'sh.TextFrame2.TextRange.Text', 'ws.Shapes.Range(Array(1)).HorizontalFlip',
    'ws.Shapes.Range(Array(1)).VerticalFlip', 'ws.Shapes.Range(Array(1)).Item(1).Name', 'ws.Shapes.Item(1).Name',
    'TypeName(wb.Sheets.Application)', 'TypeName(wb.Sheets.Parent)',
    'sf.DataOption', 'sf.Key.Address', 'sf.Order',
    'sh.TextFrame.AutoSize', 'sh.TextFrame.MarginLeft', 'sh.TextFrame.MarginRight', 'sh.TextFrame.MarginTop',
    'sh.TextFrame.MarginBottom', 'sh.TextFrame.Orientation',
    'wb.ActiveSheet.Name', 'TypeName(wb.Application)', 'wb.FullName = wb.Name', 'TypeName(wb.Parent)', 'wb.Path',
    'wb.Saved', 'wb.Sheets.Count', 'wb.Queries.Count',
    'TypeName(Workbooks.Application)', 'Workbooks.Item(wb.Name).Name = wb.Name', 'TypeName(Workbooks.Parent)',
    'TypeName(ws.Application)', 'ws.Cells.Count', 'ws.Columns.Count', 'ws.Rows.Count', 'TypeName(ws.Parent)',
    'ws.Shapes.Count', 'ws.Type', 'ws.DrawingObjects.Count',
]


def _probe() -> str:
    lines = ['Public Function Probe() As String',
             '    Dim wb As Workbook, ws As Worksheet, sh As Shape, co As ChartObject, ch As Chart, lo As ListObject',
             '    Dim cs As Object, uv As Object, sf As Object, s As String, t As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array(1, 2, 3, 4))',
             '    ws.Range("C1:D3").Value = Array("h1", "h2")',
             '    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("C1:D3"), , xlYes)',
             '    Set sh = ws.Shapes.AddShape(msoShapeRectangle, 200, 10, 100, 40)', '    sh.TextFrame2.TextRange.Text = "hi"',
             '    Set co = ws.ChartObjects.Add(320, 10, 200, 120): Set ch = co.Chart',
             '    ch.SetSourceData ws.Range("A1:A4")', '    ws.Range("A1").AddComment "n"',
             '    Set sf = ws.Sort.SortFields.Add(ws.Range("A1:A4"))',
             '    Set cs = ws.Range("A1:A4").FormatConditions.AddColorScale(3)',
             '    Set uv = ws.Range("A1:A4").FormatConditions.AddUniqueValues', '    cs.SetFirstPriority',
             '    wb.Names.Add "Data", "=Sheet1!$A$1:$A$4"', '    ws.Range("B2").Select']
    for read in READS:
        lines += ['    Err.Clear', '    t = ""', '    On Error Resume Next', f'    t = CStr({read})',
                  '    If Err.Number Then t = "E" & Err.Number', '    On Error GoTo 0', '    s = s & "@@@" & t']
    lines += ['    wb.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(READS, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['cases']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
