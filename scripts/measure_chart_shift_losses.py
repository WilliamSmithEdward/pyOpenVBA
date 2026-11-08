"""Inserts and deletes beside a chart whose points carry their own look or label and whose error bars read cells, in a
chart Excel made in the session and in one read from a file this project wrote: what it reads afterwards, and the
references, cached points, numbered points and error bars of the chart part Excel saves.

The chart reads Data!A1:C5 by columns, its first series' second point is painted red, its second series' third point
has a label, and its first series has custom error bars reading D2:D5 above and E2:E5 below. A case is a way --
"session", Excel making the chart itself, or "loaded", Excel opening the book this project saved -- and an edit.

    python scripts/measure_chart_shift_losses.py [round]
    python scripts/measure_chart_shift_losses.py predict 1   # tests/fixtures/chart_shift_losses1_predictions.json

writes tests/fixtures/chart_shift_losses<round>.json, which tests/test_excel_chart_shift_losses.py replays.
"""
from __future__ import annotations

import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / 'tests/fixtures'
BUILD = '''Public Sub Build()
    Dim d As Worksheet, ch As Object
    Set d = ActiveWorkbook.Worksheets(1): d.Name = "Data"
    d.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))
    d.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))
    d.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))
    d.Range("D2:D5").Value = Application.Transpose(Array(0.5, 0.25, 0.75, 1))
    d.Range("E2:E5").Value = Application.Transpose(Array(0.1, 0.2, 0.3, 0.4))
    Set ch = d.ChartObjects.Add(300, 100, 300, 200).Chart
    ch.ChartType = xlColumnClustered
    ch.SetSourceData d.Range("A1:C5"), xlColumns
    ch.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = 255
    ch.SeriesCollection(2).Points(3).HasDataLabel = True
    ch.SeriesCollection(1).ErrorBar xlY, xlErrorBarIncludeBoth, xlErrorBarTypeCustom, d.Range("D2:D5"), d.Range("E2:E5")
End Sub
'''
SOURCE = BUILD + '''
Private Function Reads(ch As Object, ByVal points As Boolean) As String
    Dim s As String, v As String, n As Long, i As Long, k As Long
    On Error Resume Next
    s = "n=" & ch.SeriesCollection.Count
    For k = 1 To ch.SeriesCollection.Count
        v = "-": v = ch.SeriesCollection(k).Formula
        s = s & " |" & v
        n = -1: n = ch.SeriesCollection(k).Points.Count
        s = s & " pts=" & n
        v = "-": v = ch.SeriesCollection(k).HasErrorBars
        s = s & " eb=" & v
        If Not points Then n = 0
        For i = 1 To n
            v = "-"
            If k = 1 Then v = ch.SeriesCollection(k).Points(i).Format.Fill.ForeColor.RGB _
                Else v = ch.SeriesCollection(k).Points(i).HasDataLabel
            s = s & " " & v
        Next
    Next
    Err.Clear
    Reads = s
End Function

Public Function Probe(ByVal way As String, ByVal path As String, ByVal folder As String, ByVal edit As String) As String
    Dim wb As Workbook, d As Worksheet, s As String
    Application.DisplayAlerts = False
    If way = "loaded" Then
        Set wb = Workbooks.Open(path)
    Else
        Set wb = Workbooks.Add(xlWBATWorksheet)
        Build
    End If
    Set d = wb.Worksheets("Data")
    d.Activate
    On Error Resume Next
    Err.Clear
    Select Case edit
'''
EDITS = ('', 'd.Rows(3).Insert', 'd.Rows(6).Insert', 'd.Rows(2).Delete', 'd.Rows(3).Delete', 'd.Rows("3:4").Delete',
         'd.Rows("2:5").Delete', 'd.Rows(1).Delete', 'd.Columns(1).Delete', 'd.Columns(2).Delete',
         'd.Columns(3).Delete', 'd.Columns(4).Delete', 'd.Columns("D:E").Delete', 'd.Columns(4).Insert',
         'd.Range("B3").Insert xlShiftDown', 'd.Range("B3:C3").Delete xlShiftUp', 'd.Range("B3").Delete xlShiftUp',
         'd.Range("D2:D5").Delete xlShiftUp', 'd.Range("D3").Delete xlShiftUp', 'd.Range("D3").Insert xlShiftDown',
         'd.Range("B2:B5").Delete xlShiftToLeft', 'd.Range("A3:E3").Delete xlShiftUp',
         'd.Range("A3:E3").Insert xlShiftDown')
CASES = [[way, edit] for way in ('session', 'loaded') for edit in EDITS]


def source(edits: list[str]) -> str:
    lines = [SOURCE]
    for edit in sorted(set(edits)):
        label = edit.replace('"', '""')
        lines.append(f'        Case "{label}"' + (f'\n            {edit}' if edit else ''))
    lines.append('''    End Select
    s = "err=" & Err.Number & " " & Reads(d.ChartObjects(1).Chart, way = "session")
    Err.Clear
    On Error GoTo 0
    wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
''')
    return '\n'.join(lines)


def part(path: Path) -> list[object]:
    """The chart part's series: each reference with its cached points, the points it numbers with a look and with a
    label, and its error bars with their caches left out."""
    with zipfile.ZipFile(path) as package:
        xml = package.read('xl/charts/chart1.xml').decode('utf-8')
    found: list[object] = []
    for series in re.findall(r'<c:ser>(.*?)</c:ser>', xml, re.DOTALL):
        bars = re.findall(r'<c:errBars>(.*?)</c:errBars>', series, re.DOTALL)
        plain = re.sub(r'<c:errBars>.*?</c:errBars>', '', series, flags=re.DOTALL)
        found.append({
            'references': [[reference, re.findall(r'<c:v>(.*?)</c:v>', cache)] for reference, cache in re.findall(
                r'<c:f>(.*?)</c:f>(?:<c:(?:num|str)Cache>(.*?)</c:(?:num|str)Cache>)?', plain, re.DOTALL)],
            'looks': re.findall(r'<c:dPt><c:idx val="(\d+)"/>', plain),
            'labels': re.findall(r'<c:dLbl><c:idx val="(\d+)"/>', plain),
            'bars': [re.sub(r'<c:(num|str)Cache>.*?</c:\1>', '', bar, flags=re.DOTALL) for bar in bars],
            'bar values': [re.findall(r'<c:v>(.*?)</c:v>', bar) for bar in bars],
        })
    return found


def build(path: Path) -> None:
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(BUILD, name='Builder')
    app.run('Builder.Build')
    app.save(path)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    stem = 'chart_shift_losses' + (sys.argv[1] if len(sys.argv) > 1 else '')
    predicted = FIXTURES / f'{stem}_predictions.json'
    cases: list[dict[str, Any]] = (json.loads(predicted.read_text('utf-8'))['cases'] if predicted.exists()
                                   else [{'call': one} for one in CASES])
    calls = [case['call'] for case in cases]
    text = source([call[1] for call in calls])
    answers: list[str] = []
    parts: list[object] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        base = Path(folder) / 'base.xlsx'
        build(base)
        excel.new_document()
        for way, edit in calls:
            result = excel.run_vba(text, 'Probe', (way, str(base), folder, edit), timeout=120.0, line_numbers=False)
            answers.append(str(result.value) if result.ok and not result.dialogs
                           else f'ERROR {result.ok} dialogs={len(result.dialogs)} {str(result)[:300]}')
            saved = Path(folder) / 'saved.xlsx'
            parts.append(part(saved) if saved.exists() else None)
            saved.unlink(missing_ok=True)
    (FIXTURES / f'{stem}.json').write_text(json.dumps({
        'measured_at': datetime.now(timezone.utc).isoformat(), 'build': BUILD, 'source': text,
        'cases': [[*call, answer, found] for call, answer, found in zip(calls, answers, parts, strict=True)]},
        indent=1) + '\n', encoding='utf-8', newline='\n')
    if predicted.exists():
        wrong = [case for case, answer, found in zip(cases, answers, parts, strict=True)
                 if case['model'] is not None and [case['model'], case['part']] != [answer, found]]
        refused = sum(case['model'] is None for case in cases)
        print(f'{len(cases) - len(wrong) - refused} right, {len(wrong)} wrong, {refused} refused of {len(cases)}')
        for one in wrong[:10]:
            print(str(one['call']), str(one['model'])[:300])
    else:
        print(len(answers), 'asked,', sum(answer.startswith('ERROR') for answer in answers), 'errors')


def model_answer(call: list[str], folder: Path) -> tuple[str | None, object]:
    """The model's answer to one case and the part it saves; None where it refuses."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAUnsupportedError

    base = folder / 'base.xlsx'
    if not base.exists():
        build(base)
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source([call[1]]), name='Probe')
    saved = folder / 'saved.xlsx'
    saved.unlink(missing_ok=True)
    try:
        answer = str(app.run('Probe.Probe', call[0], str(base), str(folder), call[1]))
    except VBAUnsupportedError:
        return None, None
    return answer, part(saved)


def held_out(round_: str) -> list[list[str]]:
    """Fresh edits for the round, either way: rows and columns through the chart's cells and its error bars' cells,
    and blocks of cells with each Shift."""
    import random

    pick = random.Random('chart_shift_losses' + round_)
    letters = 'ABCDEF'
    found: list[list[str]] = []
    while len(found) < 100:
        way = pick.choice(('session', 'loaded'))
        verb = pick.choice(('Insert', 'Delete'))
        shape = pick.choice(('rows', 'columns', 'cells', 'cells'))
        if shape == 'rows':
            top = pick.randint(1, 7)
            edit = f'd.Rows("{top}:{top + pick.randint(0, 2)}").{verb}'
        elif shape == 'columns':
            left = pick.randint(0, 4)
            edit = f'd.Columns("{letters[left]}:{letters[left + pick.randint(0, 1)]}").{verb}'
        else:
            top, left = pick.randint(1, 6), pick.randint(0, 4)
            bottom, right = top + pick.choice((0, 0, 1, 3)), min(5, left + pick.choice((0, 0, 1)))
            target = f'{letters[left]}{top}:{letters[right]}{bottom}'
            how = pick.choice(('xlShiftDown', 'xlShiftToRight') if verb == 'Insert' else ('xlShiftUp', 'xlShiftToLeft'))
            edit = f'd.Range("{target}").{verb} {how}'
        if [way, edit] not in found:
            found.append([way, edit])
    return found


def predict(round_: str) -> None:
    folder = Path(tempfile.mkdtemp())
    found = []
    for call in held_out(round_):
        model, part_found = model_answer(call, folder)
        found.append({'call': call, 'model': model, 'part': part_found})
    (FIXTURES / f'chart_shift_losses{round_}_predictions.json').write_text(json.dumps({'cases': found}, indent=0),
                                                                           encoding='utf-8', newline='\n')
    print(len(found), 'predicted,', sum(one['model'] is None for one in found), 'refused')


if __name__ == '__main__':
    if len(sys.argv) > 2 and sys.argv[1] == 'predict':
        predict(sys.argv[2])
    else:
        main()
