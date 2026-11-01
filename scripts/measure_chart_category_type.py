"""Axis.CategoryType, BaseUnit, BaseUnitIsAuto and MajorUnitScale: what a chart's category axis reads before and
after each change, the errors, and the chart part Excel saves.

Each case makes a workbook, adds a chart through ChartObjects.Add, fills categories in A2:A4, text or dates as the
case says, and two series in B and C, gives them to the chart with SetSourceData, sets the case's chart type and runs
the case's statements with ax as the category axis, then reads CategoryType, BaseUnit and BaseUnitIsAuto, each as its
value or "E" and the error number, and saves it. The chart part is recorded as scripts/measure_chart_markup.py
records it.

The unit scales are read and set only with an explicit MajorUnit. Left alone, Excel works them out from the chart's
automatic scale, which differs between kinds (a line chart's reads months where a column chart's reads days) and
writes a major unit of its own; the model does not work that scale out.

    python scripts/measure_chart_category_type.py

writes tests/fixtures/chart_category_type.json, which tests/test_excel_chart_category_type.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import chart_part

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_category_type.json'
CASES = [
    ('text', 'xlColumnClustered', ''),
    ('text', 'xlColumnClustered', 'ax.CategoryType = xlTimeScale'),
    ('text', 'xlColumnClustered', 'ax.CategoryType = xlCategoryScale'),
    ('text', 'xlColumnClustered', 'ax.CategoryType = xlCategoryScale: ax.CategoryType = xlAutomaticScale'),
    ('dates', 'xlColumnClustered', ''),
    ('dates', 'xlColumnClustered', 'ax.CategoryType = xlCategoryScale'),
    ('dates', 'xlColumnClustered', 'ax.CategoryType = xlTimeScale'),
    ('dates', 'xlColumnClustered', 'ax.BaseUnit = xlMonths'),
    ('dates', 'xlColumnClustered', 'ax.BaseUnit = xlYears'),
    ('dates', 'xlColumnClustered', 'ax.BaseUnit = xlMonths: ax.BaseUnitIsAuto = True'),
    ('dates', 'xlColumnClustered', 'ax.BaseUnit = 5'),
    ('dates', 'xlColumnClustered', 'ax.MajorUnit = 2: ax.MajorUnitScale = xlMonths'),
    ('dates', 'xlLine', ''),
    ('dates', 'xlLine', 'ax.CategoryType = xlCategoryScale'),
    ('text', 'xlLine', 'ax.BaseUnit = xlDays'),
    ('dates', 'xlColumnClustered', 'ch.Axes(xlValue).CategoryType = xlTimeScale'),
    ('dates', 'xlColumnClustered', 'ax.CategoryType = 7'),
]
NAMES = ('CategoryType', 'BaseUnit', 'BaseUnitIsAuto')


def _probe() -> str:
    lines = ['Private Function Safe(ByVal target As Object, ByVal name As String) As String',
             '    On Error GoTo Failed', '    Safe = CallByName(target, name, VbGet)', '    Exit Function', 'Failed:',
             '    Safe = "E" & Err.Number', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal dated As Boolean, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, ax As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ws.Range("A1").Value = "k"',
             '    If dated Then',
             '        ws.Range("A2").Value = DateSerial(2026, 1, 1): ws.Range("A3").Value = DateSerial(2026, 1, 15)',
             '        ws.Range("A4").Value = DateSerial(2026, 3, 1)',
             '    Else',
             '        ws.Range("A2:A4").Value = Application.Transpose(Array("a", "b", "c"))',
             '    End If',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    ws.Range("C1:C4").Value = Application.Transpose(Array("w", 4, 0.5, 6))',
             '    ch.SetSourceData ws.Range("A1:C4")', '    On Error GoTo Failed', '    Select Case which']
    for index, (_, kind, statements) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}', '        Set ax = ch.Axes(xlCategory)']
        lines += [f'        {statements}'] if statements else []
    lines += ['    End Select', '    Set ax = ch.Axes(xlCategory)',
              '    s = ' + ' & "," & '.join(f'Safe(ax, "{name}")' for name in NAMES),
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String',
              '    Application.DisplayAlerts = False']
    for index, (data, _, _) in enumerate(CASES):
        lines.append(f'    s = s & "@@@" & Tried({index}, {"True" if data == "dates" else "False"}, folder)')
    lines += ['    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = [chart_part(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[f'{data} {kind}: {statements}', answer, part]
                        for (data, kind, statements), answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, part in record['cases']:
        axis = re.search(r'<c:(?:catAx|dateAx)>.*?</c:(?:catAx|dateAx)>', part)
        print(case, '->', answer, '||', re.sub(r'<c:(?:axId|scaling|delete|axPos|numFmt|majorTickMark|minorTickMark|'
                                               r'tickLblPos|crossAx|crosses)\b.*?/>|<c:scaling>.*?</c:scaling>', '',
                                               axis.group(0) if axis else ''))


if __name__ == '__main__':
    main()
