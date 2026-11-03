"""ListObject.Unlist: what it leaves of the table on the sheet, what the cells read and what the file holds.

Each case makes a workbook with a table on A1:C5 of Sheet1, runs the case's statements, unlists it, reads READS, each as
its value or "E" and an error number, and saves the workbook. Every built-in style is tried with the table's options as they
start and as OPTIONS sets them, and only a digest of its cells and formats is kept. The definitions of those styles, as
TableStyleElements reads them, are kept packed, for scripts/build_table_styles.py.

    python scripts/measure_list_unlist.py

writes tests/fixtures/list_unlist.json, which tests/test_excel_list_unlist.py replays.
"""
from __future__ import annotations

import base64
import hashlib
import json
import re
import tempfile
import zipfile
import zlib
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/list_unlist.json'
STYLES = [f'TableStyle{name}{number}' for name, last in (('Light', 21), ('Medium', 28), ('Dark', 11))
          for number in range(1, last + 1)]
OPTIONS = ['', 'o.ShowTotals = True: o.ShowTableStyleColumnStripes = True: o.ShowTableStyleFirstColumn = True: o.ShowTableStyleLastColumn = True',
           'o.ShowTotals = True: o.ShowTableStyleRowStripes = False', 'o.ShowTableStyleRowStripes = False: o.ShowTableStyleColumnStripes = True: o.ShowTableStyleLastColumn = True']
SWEEP = [f'o.TableStyle = "{style}"' + (f': {option}' if option else '') for style in STYLES for option in OPTIONS]
CASES = [
    '', 'o.TableStyle = ""',
    'o.ShowTotals = True', 'o.ShowTableStyleRowStripes = False', 'o.ShowTableStyleColumnStripes = True',
    'o.ShowTableStyleFirstColumn = True', 'o.ShowTableStyleLastColumn = True', 'o.ShowHeaders = False',
    'o.ShowAutoFilter = False', 'o.ShowAutoFilterDropDown = False', 'ws.Range("B2").Interior.Color = 255',
    'ws.Range("A1").Font.Italic = True', 'ws.Range("B3").Font.Color = 255', 'ws.Range("C4").Borders.LineStyle = xlContinuous',
    'ws.Range("A1:C5").NumberFormat = "0.00"', 'ws.Range("E1").Formula = "=SUM(T[a])"', 'ws.Range("E1").Formula = "=T[[#Headers],[a]]"',
    'ws.Range("C2").Formula = "=[@a]+1"', 'ws.Range("E1").Formula = "=SUM(T)"', 'ws.Range("E1").Formula = "=SUM(T[[a]:[b]])"', 'ws.Range("E2").Formula = "=[@a]"', 'ws.Range("E1").Formula = "=ROWS(T[#All])"',
    'o.Resize ws.Range("A1:A3")', 'o.Resize ws.Range("A1:B2")', 'o.Resize ws.Range("A1:C2"): o.TableStyle = "TableStyleLight1"', 'o.TableStyle = "TableStyleLight15": o.Resize ws.Range("A1:B3")', 'o.TableStyle = "TableStyleMedium9": o.ShowTotals = True: o.Resize ws.Range("A1:A4")', 'ws.Range("A1:C5").Cut ws.Range("D3")', 'ws.Range("B2").Font.Bold = False', 'ws.Range("A1").Font.Bold = False', 'ws.Range("A1").Font.Color = 0', 'ws.Range("A3").Interior.Pattern = xlNone', 'ws.Range("B2").Interior.Pattern = xlNone', 'ws.Range("A2").Borders(xlEdgeTop).LineStyle = xlNone', 'ws.Range("A1").Style = "Good"', 'o.Range.AutoFilter 1, ">1"', 'o.Range.Sort ws.Range("A1"), xlDescending, Header:=xlYes',
    'o.ListColumns.Add', 'o.ListRows.Add', 'o.Resize ws.Range("A1:D7")', 'ws.Shapes.AddChart2(201, xlColumnClustered).Chart.SetSourceData o.Range',
]
READS = ['wb.Worksheets(1).ListObjects.Count', 'ws.Range("A1").Value', 'ws.Range("A1").Font.Bold', 'ws.Range("A1").Font.Color',
         'ws.Range("A1").Interior.Color', 'ws.Range("A1").Interior.Pattern', 'ws.Range("A2").Interior.Color', 'ws.Range("A3").Interior.Color',
         'ws.Range("A3").Interior.Pattern', 'ws.Range("A3").Interior.ThemeColor', 'ws.Range("A3").Interior.TintAndShade', 'ws.Range("A1").Borders(xlEdgeBottom).LineStyle',
         'ws.Range("A2").Borders(xlEdgeTop).LineStyle', 'ws.AutoFilterMode', 'ws.Range("E1").Formula', 'ws.Range("C2").Formula', 'ws.Range("A1").ListObject Is Nothing',
         'wb.Names.Count', 'ws.Names.Count', 'Err.Number', 'TypeName(ws.Range("A1").Style)', 'ws.Range("A1").Style.Name']


def _probe(cases: list[str]) -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, o As Object, s As String, v As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:C1").Value = Array("a", "b", "c")', '    ws.Range("A2:C2").Value = Array(1, 2, 3)',
             '    ws.Range("A3:C3").Value = Array(4, 5, 6)', '    ws.Range("A4:C4").Value = Array(7, 8, 9)', '    ws.Range("A5:C5").Value = Array(10, 11, 12)',
             '    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C5"), , xlYes)', '    o.Name = "T"', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(cases):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    o.Unlist', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    Application.DisplayAlerts = False', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook',
              '    wb.Close False', '    Application.DisplayAlerts = True', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def parts_of(path: Path) -> dict[str, str]:
    """The sheet, the styles, the workbook and whatever of tables the package holds."""
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if name in ('xl/worksheets/sheet1.xml', 'xl/styles.xml', 'xl/workbook.xml', '[Content_Types].xml') or re.search(r'tables?/|sheet1\.xml\.rels', name):
                found[name] = package.read(name).decode('utf-8')
    return found


def digest(sheet: str, styles: str) -> str:
    """A digest of the cells of a sheet and the formats of its stylesheet."""
    cells = sheet[sheet.index('<sheetData'):sheet.index('</sheetData>')]
    # A row's height follows the display's scale.
    cells = re.sub(r' x14ac:dyDescent="[^"]*"| ht="[^"]*"', '', cells)
    formats = styles[styles.index('<fonts'):styles.index('</cellXfs>')]
    return hashlib.sha256((cells + formats).encode('utf-8')).hexdigest()


DEFINITIONS = """Private Function Attr(ByVal e As Object, ByVal which As Long, ByVal side As Long) As String
    On Error GoTo Failed
    Select Case which
    Case 0: Attr = e.Font.Bold
    Case 1: Attr = e.Font.Italic
    Case 2: Attr = e.Font.Color
    Case 3: Attr = e.Font.ThemeColor
    Case 4: Attr = e.Font.TintAndShade
    Case 5: Attr = e.Interior.Pattern
    Case 6: Attr = e.Interior.Color
    Case 7: Attr = e.Interior.ThemeColor
    Case 8: Attr = e.Interior.TintAndShade
    Case 9: Attr = e.Borders(side).LineStyle
    Case 10: Attr = e.Borders(side).Weight
    Case 11: Attr = e.Borders(side).ThemeColor
    Case 12: Attr = e.Borders(side).TintAndShade
    Case 13: Attr = e.Borders(side).Color
    Case 14: Attr = e.HasFormat
    End Select
    If Attr = "" Then Attr = "-"
    Exit Function
Failed:
    Attr = "E" & Err.Number
End Function

Public Function Probe(ByVal name As String) As String
    Dim t As Object, e As Object, i As Long, k As Long, b As Variant, s As String
    Set t = ActiveWorkbook.TableStyles(name)
    For i = 0 To 12
        Set e = t.TableStyleElements(i)
        s = s & "@" & i
        s = s & "," & Attr(e, 14, 0)
        For k = 0 To 8: s = s & "," & Attr(e, k, 0): Next
        For Each b In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlInsideVertical, xlInsideHorizontal)
            For k = 9 To 13: s = s & "," & Attr(e, k, b): Next
        Next
    Next
    Probe = s
End Function
"""


def main() -> None:
    every = [*CASES, *SWEEP]
    source = _probe(every)
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(every), 25):
            last = min(first + 24, len(every) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [parts_of(Path(folder) / f'c{index}.xlsx') for index in range(len(every))]
        definitions: dict[str, str] = {}
        for style in STYLES:
            result = excel.run_vba(DEFINITIONS, 'Probe', (style,), timeout=60.0)
            assert result.ok, result
            definitions[style] = str(result.value)
    packed = base64.b85encode(zlib.compress(json.dumps(definitions).encode('utf-8'), 9)).decode('ascii')
    swept = {case: digest(part['xl/worksheets/sheet1.xml'], part['xl/styles.xml'])
             for case, part in zip(SWEEP, parts[len(CASES):], strict=True)}
    kept = [{name: text for name, text in part.items() if name in ('xl/worksheets/sheet1.xml', 'xl/styles.xml')}
            for part in parts[:len(CASES)]]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': _probe(CASES), 'definitions': packed,
              'swept': swept, 'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, kept, strict=False)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)
    Path(tempfile.gettempdir(), 'unlist_sweep.json').write_text(json.dumps(dict(zip(SWEEP, parts[len(CASES):], strict=True))), encoding='utf-8')


if __name__ == '__main__':
    main()
