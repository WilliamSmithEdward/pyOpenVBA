"""Which styles Shapes.AddChart2 takes, the style part each writes, what each reads for each chart type the engine writes, and
the style each chart type takes by default; then, with "checks", a digest of the chart part of each chart type at each style
some chart type takes by default.

    python scripts/measure_chart_style_parts.py
    python scripts/measure_chart_style_parts.py checks

writes tests/fixtures/chart_style_parts.json, its parts and reads packed, from which scripts/build_chart_styles.py makes the
styles table and tests/test_excel_add_chart2.py checks the digests. The second takes the first's answers from the file.
"""
from __future__ import annotations

import base64
import hashlib
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import zlib

from pyvbaharness import ExcelSession, HarnessConfig

from measure_chart_markup import normalized

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/chart_style_parts.json'
STYLES = list(range(201, 521))
#: Every XlChartType, as the type library lists them.
KINDS = [51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85,
         86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 15, 87, 1, 4, 5, 76,
         -4098, -4100, -4101, -4102, -4120, -4151, -4169]
PROBE = """Public Function Probe(ByVal folder As String, ByVal style As Long, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Object, n As Object, s As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))
    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))
    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))
    ws.Range("A1:C5").Select
    On Error Resume Next
    Set n = ws.Shapes.AddChart2(style, kind)
    If Err.Number Then
        s = "Err " & Err.Number
    Else
        s = n.Chart.ChartStyle & " " & n.Chart.ChartType & " " & n.Chart.ChartColor
    End If
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.SaveAs folder & "\\p.xlsx", xlOpenXMLWorkbook
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
End Function
"""

#: The chart types the engine writes.
WRITTEN = [51, 52, 53, 57, 58, 59, 4, 65, 63, 64, 66, 67, 1, 76, 77, 5, -4120, -4169]
READS = """Public Function Reads(ByVal first As Long, ByVal last As Long, ByVal kind As Long) As String
    Dim wb As Workbook, ws As Object, n As Object, s As String, style As Long, one As String
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:A5").Value = Application.Transpose(Array("k", "a", "b", "c", "d"))
    ws.Range("B1:B5").Value = Application.Transpose(Array("v", 1, 2, 3, 4))
    ws.Range("C1:C5").Value = Application.Transpose(Array("w", 5, 6, 7, 8))
    ws.Range("A1:C5").Select
    On Error Resume Next
    For style = first To last
        Set n = ws.Shapes.AddChart2(style, kind)
        one = n.Chart.ChartType
        one = one & "," & Attr(n, "G"): one = one & "," & Attr(n, "O"): one = one & "," & Attr(n, "C"): one = one & "," & Attr(n, "V")
        one = one & "," & Attr(n, "L"): one = one & "," & Attr(n, "M"): one = one & "," & Attr(n, "H"): one = one & "," & Attr(n, "S")
        s = s & "|" & one
        n.Delete
    Next
    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    Reads = Mid(s, 2)
End Function

Private Function Attr(ByVal n As Object, ByVal which As String) As String
    On Error GoTo Failed
    Select Case which
    Case "G": Attr = n.Chart.ChartGroups(1).GapWidth
    Case "O": Attr = n.Chart.ChartGroups(1).Overlap
    Case "C": Attr = n.Chart.Axes(xlCategory).MajorTickMark
    Case "V": Attr = n.Chart.Axes(xlValue).MajorTickMark
    Case "L": Attr = n.Chart.HasLegend
    Case "M": Attr = n.Chart.SeriesCollection(1).MarkerStyle
    Case "H": Attr = n.Chart.Axes(xlCategory).HasMajorGridlines
    Case "S": Attr = n.Chart.Axes(xlValue).HasMajorGridlines
    End Select
    Exit Function
Failed:
    Attr = "E"
End Function
"""


def _style_part(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        if 'xl/charts/style1.xml' not in package.namelist():
            return ''
        return package.read('xl/charts/style1.xml').decode('utf-8')


def main() -> None:
    parts: dict[str, str] = {}
    answers: dict[str, str] = {}
    defaults: dict[str, str] = {}
    looks: dict[str, list[str]] = {}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for style in STYLES:
            result = excel.run_vba(PROBE, 'Probe', (folder, style, 51), timeout=60.0)
            assert result.ok and not result.dialogs, result
            answers[str(style)] = str(result.value)
            part = _style_part(Path(folder) / 'p.xlsx')
            if part:
                parts[str(style)] = part
        for kind in sorted(set(KINDS)):
            result = excel.run_vba(PROBE, 'Probe', (folder, -1, kind), timeout=60.0)
            assert result.ok and not result.dialogs, result
            defaults[str(kind)] = str(result.value)
        for kind in WRITTEN:
            for first in range(201, 498, 60):
                result = excel.run_vba(READS, 'Reads', (first, min(first + 59, 497), kind), timeout=280.0)
                assert result.ok and not result.dialogs, result
                looks.setdefault(str(kind), []).extend(str(result.value).split('|'))
            print(kind, flush=True)
    packed = base64.b85encode(zlib.compress(json.dumps({'looks': looks, 'parts': parts}).encode('utf-8'), 9)).decode('ascii')
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers, 'defaults': defaults, 'packed': packed},
                              indent=1) + '\n', encoding='utf-8')
    print(len(parts), 'parts')


def layout_of(part: str) -> list[object]:
    """What a chart part says of a style's layout: a title, the legend's place, each axis's minor gridlines, minor tick
    marks and whether it is deleted, its first series' data labels, drop lines, high-low lines, each axis's title."""
    legend = re.search(r'<c:legendPos val="(\w+)"', part)
    axes = re.findall(r'<c:(?:catAx|valAx|dateAx)>.*?</c:(?:catAx|valAx|dateAx)>', part, re.S)
    said: list[object] = ['<c:chart><c:title>' in part, legend.group(1) if legend else '']
    for axis in axes[:2]:
        tick = re.search(r'<c:minorTickMark val="(\w+)"', axis)
        said += ['<c:minorGridlines>' in axis, tick.group(1) if tick else 'none', '<c:delete val="1"/>' in axis]
    said += [False, 'none', False] * (2 - len(axes[:2]))
    labels: object = False
    if '<c:ser>' in part:
        found = re.search(r'<c:dLbls>.*?</c:dLbls>', part[part.index('<c:ser>'):part.index('</c:ser>')], re.S)
        if found:
            place = re.search(r'<c:dLblPos val="(\w+)"', found.group(0))
            labels = [*('<c:show%s val="1"/>' % name in found.group(0) for name in ('Val', 'CatName', 'SerName', 'Percent')),
                      place.group(1) if place else '']
    return [*said, labels, '<c:dropLines>' in part, '<c:hiLowLines>' in part,
            *('<c:title>' in axis for axis in (axes + ['', ''])[:2])]


def checks(kinds: list[int]) -> None:
    """A digest of the chart part of each kind at each style with no data labels in its layout, and at each style some kind
    takes by default; added to what the file holds, as this takes minutes a kind. A column chart also gives the layouts."""
    record = json.loads(OUT.read_text(encoding='utf-8'))
    defaults = {int(answer.split()[0]) for answer in record['defaults'].values() if not answer.startswith('Err')}
    laid = record.get('layouts', {})
    styles = list(range(201, 498)) if kinds != [] else sorted(defaults)
    digests: dict[str, str] = record.setdefault('checks', {})
    doughnuts: dict[str, list[int]] = record.setdefault('doughnuts', {})
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for kind in kinds:
            for style in styles:
                if f'{kind},{style}' in digests:
                    continue
                result = excel.run_vba(PROBE, 'Probe', (folder, style, kind), timeout=60.0)
                assert result.ok and not result.dialogs, result
                with zipfile.ZipFile(Path(folder) / 'p.xlsx') as package:
                    part = normalized(package.read('xl/charts/chart1.xml').decode('utf-8'))
                digests[f'{kind},{style}'] = hashlib.sha256(part.encode('utf-8')).hexdigest()
                if kind in (51, 4, -4169, 5):
                    laid.setdefault(str(kind), {})[str(style)] = layout_of(part)
                if kind == -4120:
                    hole = re.search(r'<c:holeSize val="(\d+)"', part)
                    first = re.search(r'<c:firstSliceAng val="(\d+)"', part)
                    assert hole is not None and first is not None
                    doughnuts[str(style)] = [int(hole.group(1)), int(first.group(1))]
            print(kind, flush=True)
            record['layouts'] = laid
            OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(digests), 'digests')


if __name__ == '__main__':
    if sys.argv[1:2] == ['checks']:
        checks([int(kind) for kind in sys.argv[2:]])
    else:
        main()
