"""The rows of issue report #252 (Japanese locale differences), run in this machine's en-US Excel, so the ones that
are not about locale can be told apart and held to Excel.

Predictions written before measuring, for en-US: AscW of a full-width digit is negative (Integer), IsNumeric(Empty)
True, Val of full-width digits 0, CDec("1,200") 1200, IsNumeric("(12)") and IsNumeric("$1,200") True, a non-ASCII
table name allowed, StrComp of hiragana and katakana or of full and half width 0 under vbTextCompare; dates show
m/d/yyyy and the Japanese era format is not read.

    python scripts/measure_locale_report.py

writes tests/fixtures/locale_report.json, which tests/test_vba_locale_report.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/locale_report.json'
#: Each case: a VBA expression on the sheet ws of a new workbook. ChrW keeps the source ASCII.
CASES = [
    'CStr(AscW(ChrW(&HFF11)))',
    'CStr(AscW(ChrW(&H5B9F)))',
    'TypeName(AscW(ChrW(&HFF11)))',
    'CStr(IsNumeric(Empty))',
    'CStr(Val(ChrW(&HFF11) & ChrW(&HFF12)))',
    'CStr(IsNumeric("$1,200"))',
    'CStr(IsNumeric("(12)"))',
    'CStr(CDec("1,200"))',
    'CStr(CDbl("1,200"))',
    'CStr(DateSerial(2026, 1, 5))',
    'Format(DateSerial(2026, 10, 28), "ggge")',
    'CStr(StrComp(ChrW(&H3042), ChrW(&H30A2), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF21), "a", vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF21), "A", vbBinaryCompare))',
    'TableName(ws, ChrW(&H660E) & ChrW(&H7D30))',
    'TableName(ws, "Caf" & ChrW(&HE9))',
    'Typed(ws, "' + "'" + 'AB")',
    'Typed(ws, ChrW(&HFF07) & "AB")',
    'Typed(ws, ChrW(&H2019) & "AB")',
    'Typed(ws, ChrW(&HFF08) & ChrW(&HFF11) & ChrW(&HFF12) & ChrW(&HFF09))',
    'Typed(ws, ChrW(&HFF11) & ChrW(&HFF12) & ChrW(&HFF05))',
    'Typed(ws, ChrW(&HFF34) & ChrW(&HFF32) & ChrW(&HFF35) & ChrW(&HFF25))',
    # Val reads ASCII digits only, and which spellings of a number it takes.
    'CStr(Val("1" & ChrW(&HFF11) & "2"))', 'CStr(Val("inf"))', 'CStr(Val("nan"))', 'CStr(Val("1_000"))',
    'CStr(Val("1d2"))', 'CStr(Val("1e"))', 'CStr(Val("1.5.5"))', 'CStr(Val("-.5e1x"))', 'CStr(Val("+ 1 2"))',
    # Text comparison across kana type and width: equality and order.
    'CStr(StrComp(ChrW(&HFF71), ChrW(&H30A2), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF76) & ChrW(&HFF9E), ChrW(&H30AC), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF76) & ChrW(&HFF9E), ChrW(&H30AB), vbTextCompare))',
    'CStr(StrComp(ChrW(&H3000), " ", vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF41), "b", vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF42), "a", vbTextCompare))',
    'CStr(StrComp(ChrW(&H3042), ChrW(&H30A4), vbTextCompare))',
    'CStr(StrComp(ChrW(&H30A4), ChrW(&H3042), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF11) & ChrW(&HFF10), "9", vbTextCompare))',
    'CStr(StrComp(ChrW(&H3094), ChrW(&H30F4), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF61), ChrW(&H3002), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF70), ChrW(&H30FC), vbTextCompare))',
    'CStr(StrComp(ChrW(&HFF3C), "\\", vbTextCompare))',
    'CStr(StrComp(ChrW(&HFFE5), ChrW(&HA5), vbTextCompare))',
    'CStr(InStr(1, "ab" & ChrW(&HFF23), "c", vbTextCompare))',
    'CStr(InStr(1, "x" & ChrW(&H3042) & "y", ChrW(&H30A2), vbTextCompare))',
    # Table names beyond ASCII follow the defined-name rules.
    'TableName(ws, ChrW(&H660E) & ChrW(&H7D30) & "2")', 'TableName(ws, "1abc")', 'TableName(ws, "a b")',
    'TableName(ws, "R1C1")', 'TableName(ws, "Tab.le")', 'TableName(ws, "_x")', 'TableName(ws, ChrW(&HFF41) & "x")',
]
HELPERS = '''Private Function TableName(ByVal ws As Worksheet, ByVal wanted As String) As String
    Dim t As ListObject
    ws.Range("A1:B2").Value = "x"
    ws.Range("A1").Value = "h1": ws.Range("B1").Value = "h2"
    Set t = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B2"), , xlYes)
    t.Name = wanted
    TableName = CStr(t.Name = wanted)
    t.Delete
End Function

Private Function Typed(ByVal ws As Worksheet, ByVal text As String) As String
    Dim i As Long, s As String, v As Variant
    ws.Range("D1").Value = text
    v = ws.Range("D1").Value
    For i = 1 To Len(CStr(v))
        s = s & Hex(AscW(Mid(CStr(v), i, 1))) & " "
    Next
    Typed = TypeName(v) & " " & Trim(s) & " prefix:" & Hex(AscW(ws.Range("D1").PrefixCharacter & " "))
End Function
'''


def _probe() -> str:
    lines = [HELPERS, 'Private Function Tried(ByVal which As Long, ByVal ws As Worksheet) As String',
             '    On Error GoTo Failed', '    Select Case which']
    for index, expression in enumerate(CASES):
        lines += [f'    Case {index}', f'        Tried = {expression}']
    lines += ['    End Select', '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number & " " & Err.Description',
              'End Function', '', 'Public Function Probe() As String', '    Dim s As String, i As Long, wb As Workbook',
              f'    For i = 0 To {len(CASES) - 1}', '        Set wb = Workbooks.Add(xlWBATWorksheet)',
              '        s = s & "|" & Tried(i, wb.Worksheets(1))', '        wb.Close False', '    Next',
              '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=180.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('|')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': dict(zip(CASES, answers, strict=True))}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases'].items():
        print(case, '->', answer)


if __name__ == '__main__':
    main()
