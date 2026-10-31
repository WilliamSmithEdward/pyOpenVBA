"""SortField's Priority, CustomOrder, SubField, Delete, ModifyKey and Parent, and SortFields' default member, For
Each and Parent: what each reads and does on a sheet's Sort, and the sortState a save writes after them.

    python scripts/measure_sort_fields.py

writes tests/fixtures/sort_fields.json, which tests/test_excel_sort_fields.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sort_fields.json'
SETUP = ['ws.Range("A1:C5").Value = Application.Transpose(Application.Transpose(Array('
         'Array(2, "b", 9), Array(1, "a", 8), Array(2, "a", 7), Array(1, "b", 6), Array(3, "c", 5))))',
         'Set sf = ws.Sort.SortFields', 'sf.Clear', 'sf.Add ws.Range("A1:A5")', 'sf.Add ws.Range("B1:B5")',
         'sf.Add ws.Range("C1:C5"), , xlDescending', 'ws.Sort.SetRange ws.Range("A1:C5")', 'ws.Sort.Header = xlNo']
#: Each case: statements after SETUP, an expression read back, and the file saved, if any.
CASES = [
    ([], 'Keys(ws)', 'plain.xlsx'),
    (['sf(3).Priority = 1'], 'Keys(ws)', 'first.xlsx'),
    (['sf(1).Priority = 3'], 'Keys(ws)', ''),
    (['sf(2).Priority = 2'], 'Keys(ws)', ''),
    (['sf(1).Priority = 0'], 'Keys(ws)', ''),
    (['sf(1).Priority = 4'], 'Keys(ws)', ''),
    (['sf(1).Priority = 9'], 'Keys(ws)', ''),
    (['sf(1).Priority = -1'], 'Keys(ws)', ''),
    (['sf(3).Priority = 1', 'ws.Sort.Apply'], 'Grid(ws)', ''),
    (['sf(2).Delete'], 'Keys(ws)', 'deleted.xlsx'),
    (['Set f = sf(2)', 'f.Delete'], 'TypeName(f)', ''),
    (['Set f = sf(2)', 'f.Delete'], 'f.Order', ''),
    (['Set f = sf(3)', 'sf(1).Delete'], 'f.Priority & " " & f.Key.Address', ''),
    (['sf(1).Delete', 'sf(1).Delete', 'sf(1).Delete'], 'sf.Count', ''),
    (['sf(2).Delete', 'ws.Sort.Apply'], 'Grid(ws)', ''),
    (['sf(1).ModifyKey ws.Range("C1:C5")'], 'Keys(ws)', 'modified.xlsx'),
    (['sf(1).ModifyKey ws.Range("E1")'], 'Keys(ws)', ''),
    (['sf(1).ModifyKey ws.Range("A1:B5")'], 'Keys(ws)', ''),
    (['sf(1).ModifyKey ws.Range("A1,C1")'], 'Keys(ws)', ''),
    (['sf(1).ModifyKey wb.Worksheets.Add.Range("A1:A5")'], 'sf(1).Key.Parent.Name', ''),
    (['sf(1).ModifyKey "A1"'], 'Keys(ws)', ''),
    (['sf.Clear', 'sf.Add(ws.Range("A1:A5"), xlSortOnCellColor).SortOnValue.Color = vbRed',
      'sf(1).ModifyKey ws.Range("B1:B5")'],
     'sf(1).SortOn & " " & sf(1).SortOnValue.Color & " " & sf(1).Order', 'colour.xlsx'),
    (['sf.Clear', 'sf.Add(ws.Range("A1:A5"), xlSortOnFontColor).SortOnValue.Color = vbRed',
      'sf(1).ModifyKey ws.Range("B1:B5")'], 'sf(1).SortOn & " " & sf(1).SortOnValue.Color', 'font.xlsx'),
    (['sf.Clear', 'sf.Add(ws.Range("A1:A5"), xlSortOnCellColor, xlDescending).SortOnValue.Color = vbRed',
      'sf(1).ModifyKey ws.Range("C1:C5")'], 'sf(1).SortOn & " " & sf(1).SortOnValue.Color & " " & sf(1).Order', ''),
    (['sf(2).CustomOrder = "c,b"', 'sf(2).ModifyKey ws.Range("B1:B5")'],
     'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    (['sf(2).SubField = "Population"', 'ws.Sort.Apply'], 'Grid(ws)', ''),
    ([], 'TypeName(sf(1).CustomOrder) & " [" & sf(1).CustomOrder & "]"', ''),
    (['sf.Add ws.Range("A1:A5"), , , "x,y, z"'], 'TypeName(sf(4).CustomOrder) & " [" & sf(4).CustomOrder & "]"', ''),
    (['sf.Add ws.Range("A1:A5"), , , 1'], 'TypeName(sf(4).CustomOrder) & " [" & sf(4).CustomOrder & "]"', ''),
    (['sf.Add ws.Range("A1:A5"), , , ""'], 'TypeName(sf(4).CustomOrder) & " [" & sf(4).CustomOrder & "]"', ''),
    (['sf(2).CustomOrder = "c,b , a"'], 'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', 'custom.xlsx'),
    (['sf(2).CustomOrder = "c,b,a"', 'ws.Sort.Apply'], 'Grid(ws)', ''),
    (['sf(2).CustomOrder = 1'], 'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    (['sf(2).CustomOrder = 1'], 'sf.Count', 'listed.xlsx'),
    (['sf(2).CustomOrder = 0'], 'sf.Count', ''),
    (['sf(2).CustomOrder = "c,b"', 'sf(2).CustomOrder = 0'], 'sf.Count', 'zero.xlsx'),
    (['sf(2).CustomOrder = "c,b"', 'sf(2).CustomOrder = 0'],
     'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    (['sf(2).CustomOrder = -1'], 'sf.Count', ''),
    (['sf(2).CustomOrder = "c,b"', 'sf(2).CustomOrder = ""'],
     'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    (['sf(2).CustomOrder = 999'], 'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    (['sf(2).CustomOrder = Array("c", "b")'], 'TypeName(sf(2).CustomOrder) & " [" & sf(2).CustomOrder & "]"', ''),
    ([], 'TypeName(sf(1).SubField) & " [" & sf(1).SubField & "]"', ''),
    (['sf(1).SubField = "Population"'], 'TypeName(sf(1).SubField) & " [" & sf(1).SubField & "]"', ''),
    ([], 'TypeName(sf(1).Parent) & " " & TypeName(sf.Parent) & " " & (sf.Parent Is ws.Sort) & " " '
         '& (sf(1).Parent Is sf)', ''),
    ([], 'TypeName(sf(1).Application) & " " & TypeName(sf.Application) & " " & sf(1).Creator & " " & sf.Creator', ''),
    ([], 'sf(2).Key.Address & " " & sf.Item(3).Order & " " & TypeName(sf(1))', ''),
    ([], 'Walked(sf)', ''),
    ([], 'sf(4).Order', ''),
    ([], 'sf(0).Order', ''),
    ([], 'sf("A1:A5").Order', ''),
]


def _probe() -> str:
    lines = ['Private Function Keys(ByVal ws As Worksheet) As String',
             '    Dim i As Long, s As String', '    For i = 1 To ws.Sort.SortFields.Count',
             '        With ws.Sort.SortFields(i)',
             '            s = s & ";" & .Key.Address(False, False) & ":" & .Priority & ":" & .Order',
             '        End With', '    Next', '    Keys = ws.Sort.SortFields.Count & "=" & Mid(s, 2)', 'End Function',
             '', 'Private Function Grid(ByVal ws As Worksheet) As String', '    Dim c As Range, s As String',
             '    For Each c In ws.Range("A1:C5")', '        s = s & "|" & c.Formula', '    Next',
             '    Grid = Mid(s, 2)', 'End Function', '',
             'Private Function Walked(ByVal sf As Object) As String', '    Dim f As Object, s As String',
             '    For Each f In sf', '        s = s & ";" & f.Key.Address(False, False)', '    Next',
             '    Walked = Mid(s, 2)', 'End Function', '',
             'Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, sf As Object, f As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    On Error GoTo Failed']
    lines += [f'    {statement}' for statement in SETUP]
    lines.append('    Select Case which')
    for index, (statements, read, saved) in enumerate(CASES):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {read}')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i, folder)', '    Next',
              '    Application.DisplayAlerts = True', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in CASES:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<sortState\b[^>]*/>|<sortState\b.*?</sortState>', sheet, re.DOTALL)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'setup': SETUP, 'source': source,
              'cases': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(CASES, answers, strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for statements, read, saved, answer in record['cases']:
        print(statements[-1:] or [read], '->', answer, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
