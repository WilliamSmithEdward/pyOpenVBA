"""Interior on a chart's elements, measured live: what it reads on a new chart, plain and made by AddChart2, and after a set of its
Color, ColorIndex or Pattern, beside what Format then reads. Each case in a workbook of its own, through a late-bound Object.

    python scripts/measure_chart_interior.py

writes tests/fixtures/chart_interior.json, which tests/test_excel_chart_interior.py replays.
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_interior.json'
ELEMENTS = ('ChartArea', 'PlotArea', 'Legend', 'ChartTitle', 'SeriesCollection(1)')
SETS = ('', 'Interior.Color = RGB(255, 0, 0)', 'Interior.Color = RGB(10, 20, 30)', 'Interior.ColorIndex = 3', 'Interior.ColorIndex = xlNone',
        'Interior.ColorIndex = xlAutomatic', 'Interior.Pattern = xlSolid', 'Interior.Pattern = xlNone', 'Format.Fill.ForeColor.RGB = RGB(0, 0, 255)',
        'Format.Fill.Visible = msoFalse')
READS = ('Interior.Color', 'Interior.ColorIndex', 'Interior.Pattern', 'Interior.PatternColorIndex', 'Interior.InvertIfNegative',
         'Format.Fill.Visible', 'Format.Fill.ForeColor.RGB')
KINDS = ('plain', 'styled')


def source() -> str:
    lines = ['Public Function Probe(ByVal kind As String, ByVal element As Long, ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Worksheet, c As Object, e As Object, s As String, x As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:C1").Value = Array("", "Rev", "Cost")', '    ws.Range("A2:C2").Value = Array("a", 10, 4)',
             '    ws.Range("A3:C3").Value = Array("b", 20, 6)', '    If kind = "styled" Then',
             '        Set c = ws.Shapes.AddChart2(201, xlColumnClustered).Chart', '    Else',
             '        Set c = ws.ChartObjects.Add(48, 96, 240, 144).Chart', '    End If', '    c.SetSourceData ws.Range("A1:C3")',
             '    c.HasTitle = True', '    Select Case element']
    lines += [f'    Case {number}: Set e = c.{element}' for number, element in enumerate(ELEMENTS)]
    lines += ['    End Select', '    On Error Resume Next', '    Select Case which']
    lines += [f'    Case {number}: e.{statement}' for number, statement in enumerate(SETS) if statement]
    lines += ['    End Select', '    s = Err.Number & "|"']
    for read in READS:
        lines += ['    Err.Clear', '    x = ""', f'    x = TypeName(e.{read}) & " " & e.{read}', '    s = s & x & " " & Err.Number & "|"']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Probe = s', 'End Function']
    return '\n'.join(lines)


def main() -> None:
    code = source()
    cases: list[list[object]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=600)) as excel:
        for kind in KINDS:
            for element in range(len(ELEMENTS)):
                for which in range(len(SETS)):
                    done = excel.run_vba(code, 'Probe', (kind, element, which), timeout=120.0, line_numbers=False)
                    assert done.outcome == 'passed', (kind, element, which, done.message)
                    cases.append([kind, element, which, done.value])
                    print(kind, ELEMENTS[element], SETS[which] or '(read)', '=>', done.value, flush=True)
    record = {'measured_at': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'source': code, 'elements': ELEMENTS,
              'sets': SETS, 'reads': READS, 'cases': cases}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
