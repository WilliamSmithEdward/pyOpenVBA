"""Chart.DisplayBlanksAs, PlotVisibleOnly, ShowDataLabelsOverMaximum and DisplayValueNotAvailableAsBlank: what each
reads before and after a change, the markup the saved chart holds for them, and what they read once a file that
leaves them out, or writes them with no val, is opened again.

Each case makes a workbook with a chart built through ChartObjects.Add and SetSourceData, sets its type and runs the
case's statements, then saves it. The last two cases take the first case's file, strip the four elements (or their
val attributes) from its chart, and open it again.

    python scripts/measure_chart_display.py

writes tests/fixtures/chart_display.json, which tests/test_excel_chart_display.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_display.json'
READ = ('"|" & ch.DisplayBlanksAs & "," & ch.PlotVisibleOnly & "," & ch.ShowDataLabelsOverMaximum & "," & '
        'ch.DisplayValueNotAvailableAsBlank')
#: (the type the chart takes first, the statements run after the first read)
CASES = [
    ('xlColumnClustered', ''),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = xlZero'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = xlInterpolated'),
    ('xlColumnClustered', 'ch.ChartType = xlArea'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = xlZero: ch.ChartType = xlArea'),
    ('xlArea', 'ch.DisplayBlanksAs = xlInterpolated: ch.ChartType = xlColumnClustered'),
    ('xlArea', 'ch.ChartType = xlLine'),
    ('xlPie', 'ch.DisplayBlanksAs = xlZero'),
    ('xlXYScatter', ''),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = 0'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = 4'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = -1'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = 9: ch.ChartType = xlLine'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = "2"'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = 2.6'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = 2.5'),
    ('xlColumnClustered', 'ch.DisplayValueNotAvailableAsBlank = False'),
    ('xlColumnClustered', 'ch.DisplayValueNotAvailableAsBlank = False: ch.DisplayValueNotAvailableAsBlank = True'),
    ('xlColumnClustered', 'ch.PlotVisibleOnly = 0: ch.ShowDataLabelsOverMaximum = -1'),
    ('xlColumnClustered', 'ch.PlotVisibleOnly = False: ch.PlotVisibleOnly = "True"'),
    ('xlColumnClustered', 'ch.PlotVisibleOnly = "x"'),
    ('xlColumnClustered', 'ch.DisplayBlanksAs = "x"'),
]
#: What the reopened files have stripped from the first case's chart.
STRIPS = {
    'absent': (r'<c:plotVisOnly[^>]*/>|<c:dispBlanksAs[^>]*/>|<c:showDLblsOverMax[^>]*/>'
               r'|<c:extLst><c:ext uri="\{56B9EC1D[^"]*"[^>]*>.*?</c:ext></c:extLst>', ''),
    'no val': (r'<(c:plotVisOnly|c:dispBlanksAs|c:showDLblsOverMax|c16r3:dispNaAsBlank) val="[^"]*"/>', r'<\1/>'),
}


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:A4").Value = Application.Transpose(Array("k", "a", "b", "c"))',
             '    ws.Range("B1:B4").Value = Application.Transpose(Array("v", 1, 2, 3))',
             '    Set ch = ws.ChartObjects.Add(100, 10, 300, 200).Chart', '    ch.SetSourceData ws.Range("A1:B4")',
             '    On Error GoTo Failed', '    Select Case which']
    for index, (kind, statement) in enumerate(CASES):
        lines += [f'    Case {index}', f'        ch.ChartType = {kind}: s = {READ}']
        if statement:
            lines.append(f'        {statement}')
    lines += ['    End Select', f'    Tried = "ok" & s & {READ}', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & s', '    Resume Saved',
              'End Function', '', 'Public Function Probe(ByVal folder As String) As String',
              '    Dim i As Long, s As String', '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function', '',
              'Public Function Reopen(ByVal path As String) As String', '    Dim wb As Workbook, ch As Object',
              '    Set wb = Workbooks.Open(path)', '    Set ch = wb.Worksheets(1).ChartObjects(1).Chart',
              f'    Reopen = {READ}', '    wb.Close False', 'End Function']
    return '\n'.join(lines) + '\n'


def chart_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        name = next(entry for entry in package.namelist() if re.fullmatch(r'xl/charts/chart\d+\.xml', entry))
        return package.read(name).decode('utf-8')


def display_markup(path: Path) -> str:
    """The chart's markup after its plot area, its legend left out."""
    found = re.search(r'</c:plotArea>(.*?)</c:chart>', chart_xml(path), re.S)
    assert found is not None
    return re.sub(r'<c:legend>.*?</c:legend>', '', found.group(1), flags=re.S)


def stripped(source: Path, target: Path, pattern: str, replacement: str) -> None:
    with zipfile.ZipFile(source) as old, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as new:
        for item in old.infolist():
            data = old.read(item.filename)
            if re.fullmatch(r'xl/charts/chart\d+\.xml', item.filename):
                data = re.sub(pattern, replacement, data.decode('utf-8'), flags=re.S).encode('utf-8')
            new.writestr(item, data)


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        markups = [display_markup(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
        reopened: list[list[str]] = []
        for name, (pattern, replacement) in STRIPS.items():
            target = Path(folder) / 'reopened.xlsx'
            stripped(Path(folder) / 'c0.xlsx', target, pattern, replacement)
            answer = excel.run_vba(source, 'Reopen', (str(target),), timeout=120.0)
            assert answer.ok and not answer.dialogs, answer
            reopened.append([name, str(answer.value)])
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'strips': STRIPS,
              'cases': [[f'{kind}: {statement}', answer, markup]
                        for (kind, statement), answer, markup in zip(CASES, answers, markups, strict=True)],
              'reopened': reopened}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, markup in record['cases']:
        print(case, '->', answer, '|', re.sub(r'<c:extLst>.*</c:extLst>', '<ext>', markup)[:160])
    print(reopened)


if __name__ == '__main__':
    main()
