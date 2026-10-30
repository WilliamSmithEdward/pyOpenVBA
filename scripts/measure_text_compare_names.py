"""Issue #252 follow-ups: which text comparisons ignore kana type and width (the = operator and Like under Option
Compare Text, a Dictionary in vbTextCompare, Replace, Split, InStrRev), and how ListObject.Name and DisplayName take
a name, beyond ASCII and outside the defined-name rules, and what a save keeps.

Predictions written before measuring: comparisons that order or equate strings (=, Dictionary keys) fold kana type
and width as StrComp does; searches (Replace, Split, InStrRev, Like) do not, as InStr does not. Name takes any text and
leaves DisplayName alone; DisplayName follows the defined-name rules, error 1004 otherwise.

    python scripts/measure_text_compare_names.py

writes tests/fixtures/text_compare_names.json, which tests/test_vba_text_compare_names.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/text_compare_names.json'
COMPARE_CASES = [
    'CStr(ChrW(&H3042) = ChrW(&H30A2))', 'CStr(ChrW(&HFF41) = "A")', 'CStr(ChrW(&H3042) < ChrW(&H30A4))',
    'CStr(ChrW(&HFF76) & ChrW(&HFF9E) = ChrW(&H30AC))', 'CStr(ChrW(&HFF41) Like "a")', 'CStr(ChrW(&H3042) Like ChrW(&H30A2))',
    'CStr(InStr(ChrW(&HFF41) & "b", "a"))', 'CStr(InStrRev("b" & ChrW(&HFF41), "a", -1, vbTextCompare))',
    'Replace("x" & ChrW(&HFF41) & "y", "a", "-", , , vbTextCompare)', 'CStr(UBound(Split("x" & ChrW(&HFF41) & "y", "a", , vbTextCompare)))',
    'Keys()',
    # Option Compare Text with plain ASCII: the operators, Like, Select Case and the functions' default compare.
    'CStr("a" = "A")', 'CStr("a" < "B")', 'CStr("b" > "A")', 'CStr("a" Like "A")', 'CStr("aB" Like "A*")',
    'CStr(InStr("xAy", "a"))', 'CStr(InStr(1, "xAy", "a", vbBinaryCompare))', 'CStr(StrComp("a", "A"))',
    'Replace("xAy", "a", "-")', 'CStr(UBound(Split("xAy", "a")))', 'Cased("a")',
]
SOURCE_COMPARE = '''Option Compare Text

Private Function Keys() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d.Add ChrW(&H3042), 1
    Keys = CStr(d.Exists(ChrW(&H30A2))) & " " & CStr(d.Exists(ChrW(&HFF71))) & " " & CStr(d.Exists(ChrW(&H3042)))
End Function

Private Function Cased(ByVal text As String) As String
    Select Case text
    Case "A": Cased = "case-blind"
    Case Else: Cased = "case-kept"
    End Select
End Function

Public Function Compare() As String
    Dim s As String
''' + '\n'.join(f'    s = s & "|" & {case}' for case in COMPARE_CASES) + '''
    Compare = Mid(s, 2)
End Function
'''
NAME_CASES = [
    ('Name', 'ChrW(&H660E) & ChrW(&H7D30)'), ('Name', '"a b"'), ('Name', '"1abc"'), ('Name', '"R1C1"'),
    ('Name', '"Table1"'), ('Name', '"Other"'),
    ('DisplayName', 'ChrW(&H660E) & ChrW(&H7D30)'), ('DisplayName', '"a b"'), ('DisplayName', '"1abc"'),
    ('DisplayName', '"R1C1"'), ('DisplayName', '"Tab.le"'), ('DisplayName', '"_x"'), ('DisplayName', '"A1"'),
    ('DisplayName', '"Other"'), ('DisplayName', '"Caf" & ChrW(&HE9)'),
]
SOURCE_NAMES = '''Private Function Named(ByVal which As String, ByVal wanted As String, ByVal folder As String, _
                       ByVal saved As String) As String
    Dim wb As Workbook, ws As Worksheet, t As ListObject, other As ListObject
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "h1": ws.Range("B1").Value = "h2": ws.Range("A2:B2").Value = 1
    ws.Range("D1").Value = "k": ws.Range("D2").Value = 2
    Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
    Set other = ws.ListObjects.Add(xlSrcRange, ws.Range("D1:D2"), , xlYes)
    other.DisplayName = "Other"
    On Error GoTo Failed
    If which = "Name" Then t.Name = wanted Else t.DisplayName = wanted
    Named = t.Name & "/" & t.DisplayName & "/" & ws.ListObjects(1).Name
    wb.SaveAs folder & "\\" & saved, xlOpenXMLWorkbook
Done:
    wb.Close False
    Exit Function
Failed:
    Named = "Err " & Err.Number & " " & Err.Description & " / " & t.Name & "/" & t.DisplayName
    Resume Done
End Function

Public Function Names(ByVal folder As String) As String
    Dim s As String
    Application.DisplayAlerts = False
''' + '\n'.join(f'    s = s & "@@@" & Named("{which}", {wanted}, folder, "name{index}.xlsx")'
               for index, (which, wanted) in enumerate(NAME_CASES)) + '''
    Application.DisplayAlerts = True
    Names = Mid(s, 4)
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        compared = excel.run_vba(SOURCE_COMPARE, 'Compare', timeout=120.0, module_name='TextCompare')
        assert compared.ok and not compared.dialogs, compared
        named = excel.run_vba(SOURCE_NAMES, 'Names', (folder,), timeout=300.0, module_name='TableNames')
        assert named.ok and not named.dialogs, named
        parts = {}
        for index in range(len(NAME_CASES)):
            path = Path(folder) / f'name{index}.xlsx'
            if path.exists():
                with zipfile.ZipFile(path) as package:
                    head = re.search(r'<table\b[^>]*>', package.read('xl/tables/table1.xml').decode('utf-8'))
                    parts[f'name{index}.xlsx'] = re.findall(r'\s((?:name|displayName)="[^"]*")', head.group(0)) \
                        if head else []
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'compare_source': SOURCE_COMPARE,
              'names_source': SOURCE_NAMES,
              'compare': dict(zip(COMPARE_CASES, str(compared.value).split('|'), strict=True)),
              'names': [[which, wanted, answer] for (which, wanted), answer
                        in zip(NAME_CASES, str(named.value).split('@@@'), strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['compare'].items():
        print(case, '->', answer)
    for index, (which, wanted, answer) in enumerate(record['names']):
        print(which, wanted, '->', answer, parts.get(f'name{index}.xlsx'))


if __name__ == '__main__':
    main()
