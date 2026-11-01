"""The second batch of collections' hidden _Default, reached through bracketed names as a macro reaches them: what each
answers next to Item, which it stands for, or the error.

Each collection is held in a variable of its own type, so the macro calls _Default bound early, as code that declares
its types does. Bound late, through an Object, Excel raised 438 for _Default on Shapes and ChartObjects alike; the
model does not tell early binding from late and answers as early binding does, so those cases are left out.

Each case makes a workbook whose sheet holds a list in A1:C4 (regions, products and amounts), a chart over C1:C4, a
comment on A1, a hyperlink on B1, a list object over E1:F3, a page break above row 10, a format condition, a colour
scale and an icon set on C2:C4, a custom property, an autofilter on A1:C4 and a pivot table with a calculated field and
a filter on another sheet, runs the case's statement and answers what it set. HPageBreaks is left out: its _Default
raised -2147352565 bound early as well, Excel working page breaks out with no printer to ask.

    python scripts/measure_hidden_defaults_more.py

writes tests/fixtures/hidden_defaults_more.json, which tests/test_excel_hidden_defaults_more.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/hidden_defaults_more.json'
CASES = [
    'Set cm = ws.Comments: s = cm.[_Default](1).Text & "|" & cm(1).Text',
    'Set hl = ws.Hyperlinks: s = hl.[_Default](1).Address',
    'Set lc = ws.ListObjects(1).ListColumns: s = lc.[_Default](2).Name & "|" & lc.[_Default]("y").Index',
    'Set lr = ws.ListObjects(1).ListRows: s = lr.[_Default](1).Index',
    'Set fc = ws.Range("C2:C4").FormatConditions: s = fc.[_Default](1).Type & "|" & fc(1).Type',
    'Set csc = ws.Range("C2:C4").FormatConditions(2).ColorScaleCriteria: s = csc.[_Default](1).Type',
    'Set ic = ws.Range("C2:C4").FormatConditions(3).IconCriteria: s = ic.[_Default](2).Type',
    'Set cp = ws.CustomProperties: s = cp.[_Default](1).Name',
    'Set fl = ws.AutoFilter.Filters: s = fl.[_Default](1).On',
    'Set cf = pt.CalculatedFields: s = cf.[_Default](1).Name',
    'Set pf = pt.PivotFields("Region").PivotFilters: s = pf.[_Default](1).FilterType',
    'Set co = ws.ChartObjects: s = co.[_Default](1).Name',
    'Set sc = ch.SeriesCollection: s = sc.[_Default](1).Name',
    'Set ax = ch.Axes: s = ax.[_Default](xlValue).Type',
    'Set cm = ws.Comments: s = TypeName(cm.[_Default](5))',
    'Set hl = ws.Hyperlinks: s = TypeName(hl.[_Default](0))',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As PivotTable, ch As Chart, s As String',
             '    Dim cm As Comments, hl As Hyperlinks, lc As ListColumns, lr As ListRows, hp As HPageBreaks',
             '    Dim fc As FormatConditions, csc As ColorScaleCriteria, ic As IconCriteria, cp As CustomProperties',
             '    Dim fl As Filters, cf As CalculatedFields, pf As PivotFilters, co As ChartObjects',
             '    Dim sc As SeriesCollection, ax As Axes',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("West", "Ink", 20)',
             '    ws.Range("A4:C4").Value = Array("East", "Pad", 30)',
             '    Set ch = ws.ChartObjects.Add(300, 10, 200, 100).Chart',
             '    ch.SetSourceData ws.Range("C1:C4")',
             '    ws.Range("A1").AddComment "note"',
             '    ws.Hyperlinks.Add ws.Range("B1"), "https://example.com/"',
             '    ws.Range("E1:F1").Value = Array("x", "y"): ws.Range("E2:F3").Value = 1',
             '    ws.ListObjects.Add xlSrcRange, ws.Range("E1:F3"), , xlYes',
             '    ws.HPageBreaks.Add ws.Range("A10")',
             '    ws.Range("C2:C4").FormatConditions.Add xlCellValue, xlGreater, "15"',
             '    ws.Range("C2:C4").FormatConditions.AddColorScale 3',
             '    ws.Range("C2:C4").FormatConditions.AddIconSetCondition',
             '    ws.CustomProperties.Add "Owner", "me"',
             '    ws.Range("A1:C4").AutoFilter 1, "East"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R4C3").CreatePivotTable('
             'wb.Worksheets.Add.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    pt.CalculatedFields.Add "Bonus", "=Amount*0.1"',
             '    pt.PivotFields("Region").PivotFilters.Add2 xlCaptionBeginsWith, Value1:="E"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case[:80], '->', answer)


if __name__ == '__main__':
    main()
