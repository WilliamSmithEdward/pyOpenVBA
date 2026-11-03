"""Range and WorksheetFunction members the checklist has as missing, those a macro can reach with no screen, no printer and no network.
Left out: CheckSpelling, DialogBox, FunctionWizard, PrintOut, PrintPreview, Speak, ShowCard, CopyPicture and the picture-in-cell
members, ExportAsFixedFormat, the linked data types that call a service, and RTD, StockHistory and WebService.

    python scripts/measure_range_missing.py

writes tests/fixtures/range_missing.json, which tests/test_excel_range_missing.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_missing.json'
XML = "<a><b id='1'>x</b><b id='2'>y</b><c><b>z</b><d>4</d></c><e>1.5</e><f></f></a>"
NAMES = ('ws.Range("A1").Value = "Name"\nws.Range("B1").Value = "First"\nws.Range("A2").Value = "Ann Lee"\nws.Range("A3").Value = "Bob Ray"\n'
         'ws.Range("A4").Value = "Cy Poe"\nws.Range("B2").Value = "Ann"\n')
#: Each reads into v, under On Error Resume Next.
CASES = [
    # Members on a plain cell.
    'Set v = ws.Range("A1").SoundNote',
    'Set v = ws.Range("A1").XPath',
    'v = ws.Range("A1").XPath.Value',
    'v = ws.Range("A1").XPath.Map Is Nothing',
    'Set v = ws.Range("A1").CellControl',
    'v = ws.Range("A1").ResetContents',
    'v = ws.Range("A1").RemoveControls',
    'v = ws.Range("A1").AllocateChanges',
    'v = ws.Range("A1").DiscardChanges',
    'v = ws.Range("A1").DataTypeToText',
    'v = ws.Range("A1").RefreshLinkedDataType',
    'v = ws.Range("A1").SetCellDataTypeFromCell(ws.Range("B1"))',
    'v = ws.Range("A1").EditionOptions(1, 1, 1)',
    'v = ws.Range("A1").TogglePythonMarshalMode',
    # What they do to a filled cell.
    'ws.Range("A1").Value = 5\nws.Range("A1").ResetContents\nv = ws.Range("A1").Value',
    'ws.Range("A1").Value = 5\nws.Range("A1").RemoveControls\nv = ws.Range("A1").Value',
    'ws.Range("A1").Value = 5\nws.Range("A1").DataTypeToText\nv = ws.Range("A1").Value',
    'ws.Range("A1").Value = 5\nws.Range("A1").DiscardChanges\nv = ws.Range("A1").Value',
    # FlashFill.
    NAMES + 'ws.Range("B2:B4").FlashFill\nv = ws.Range("B3").Value & "," & ws.Range("B4").Value',
    NAMES + 'ws.Range("B2").FlashFill\nv = ws.Range("B3").Value & "," & ws.Range("B4").Value',
    'ws.Range("B2:B4").FlashFill\nv = ws.Range("B3").Value',
    # A disconnected recordset.
    'Set o = CreateObject("ADODB.Recordset")\no.Fields.Append "N", 200, 10\no.Fields.Append "Q", 3\no.Open\no.AddNew\no("N") = "a"\no("Q") = 1\no.Update\no.AddNew\no("N") = "b"\no("Q") = 2\no.Update\no.MoveFirst\nv = ws.Range("A1").CopyFromRecordset(o)\nv = v & ws.Range("A1").Value & ws.Range("B2").Value',
    # FilterXML.
    *(f'ws.Range("A1").Formula2 = "=FILTERXML(""{XML}"",""{path}"")"\nv = ws.Range("A1").Text & "/" & ws.Range("A2").Text & "/" & ws.Range("A3").Text & "/" & ws.Range("A4").Text'
      for path in ('//b', '/a/b', '/a/b[1]', '/a/b[2]', '//b[@id]', '//b[@id=2]', "//b[@id='1']", '//b/@id', '/a/c/b', '//d', '//e', '//f', '//z', '/a/*', '/a/c/*',
                   'count(//b)', '//b[last()]', '//b[position()>1]', 'sum(//d)', '//d[.>3]', '//b[text()=\'y\']', 'string(//e)', '//e/text()', '/a', '[', '')),
    'ws.Range("A1").Formula2 = "=FILTERXML(""not xml"",""//a"")"\nv = ws.Range("A1").Text',
    f'ws.Range("A1").Formula2 = "=FILTERXML(""<a><n:b xmlns:n=\'\'u\'\'>1</n:b></a>"",""//b"")"\nv = ws.Range("A1").Text',
    f'v = WorksheetFunction.FilterXML("{XML}", "//e")',
    f'v = WorksheetFunction.FilterXML("{XML}", "//d")',
    f'v = WorksheetFunction.FilterXML("{XML}", "//b")\nv = UBound(v, 1) & LBound(v, 1) & v(2, 1)',
    f'v = WorksheetFunction.FilterXML("{XML}", "//z")',
    # More of FILTERXML's paths.
    *(f'ws.Range("A1").Formula2 = "=FILTERXML(""{XML}"",""{path}"")"\nv = ws.Range("A1").Text & "/" & ws.Range("A2").Text & "/" & ws.Range("A3").Text'
      for path in ('//b | //d', '//b/..', '//d/parent::c/b', '//d/preceding-sibling::b', '/a/b/following-sibling::*[1]', 'a', 'b', './/b', '/a//b', '//c//*',
                   "//b[contains(.,'y')]", "//b[starts-with(@id,'1')]", '//b[not(@id)]', "//b[@id='1' or @id='2']", "//b[@id='1' and .='x']", '//*[self::d]',
                   '//b[1]', '(//b)[1]', '(//b)[3]', '//b[3]', '//b[0]', '//b[@id>1]', '//b[@id!=1]', '//@id', '//*/@id', '/a/b/@id[.=2]', '//B', '//*[name()=\'d\']',
                   "//*[local-name()='d']", '//d[number(.)=4]', '//b[string-length(.)=1]', "//b[translate(.,'x','q')='q']", "//b[concat(.,'!')='x!']", '/a/e[.=1.5]', '//d[.+1=5]',
                   '//d[.*2=8]', '//d[. div 2=2]', '//d[. mod 3=1]', '//d[.-1=3]', '//b[position()=2]', '//b[position()<2]', 'true()', "'x'", '1')),
    # What comes back from text.
    *(f'ws.Range("A1").Formula2 = "=TYPE(FILTERXML(""<a><x>{value}</x></a>"",""//x""))&""~""&FILTERXML(""<a><x>{value}</x></a>"",""//x"")"\nv = ws.Range("A1").Text'
      for value in ('TRUE', 'true', '1e3', ' 5 ', '&lt;', '12/2/2020', '5%', '$5', '1,000', '-3', '007', '.5', 'abc', '  a  ', '#N/A', '=1+1', '<![CDATA[q]]>', '<y/>t', 't<y/>u', 'é', "&#65;")),
    'ws.Range("A1").Formula2 = "=FILTERXML(""<?xml version=\'\'1.0\'\'?><a><x>1</x></a>"",""//x"")"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula2 = "=FILTERXML(""<!DOCTYPE a><a><x>1</x></a>"",""//x"")"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula2 = "=FILTERXML(""<a><x>1</x>"",""//x"")"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula2 = "=FILTERXML("""",""//x"")"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula2 = "=FILTERXML(5,""//x"")"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula2 = "=FILTERXML(""<a>5</a>"",5)"\nv = ws.Range("A1").Text',
    'ws.Range("C1").Value = "<a><x>7</x></a>"\nws.Range("A1").Formula2 = "=FILTERXML(C1,""//x"")+1"\nv = ws.Range("A1").Text',
    'ws.Range("A1").Formula = "=FILTERXML(""<a><x>1</x><x>2</x></a>"",""//x"")"\nv = ws.Range("A1").Text & "/" & ws.Range("A2").Text & "/" & ws.Range("A1").Formula',
    'ws.Range("A1").Formula2 = "=SUM(FILTERXML(""<a><x>1</x><x>2</x></a>"",""//x""))"\nv = ws.Range("A1").Text',
    # Phonetic.
    'ws.Range("A1").Value = "abc"\nws.Range("B1").Formula = "=PHONETIC(A1)"\nv = ws.Range("B1").Text',
    'ws.Range("A1").Value = 12\nws.Range("B1").Formula = "=PHONETIC(A1)"\nv = ws.Range("B1").Text',
    'ws.Range("A1").Value = "a"\nws.Range("A2").Value = "b"\nws.Range("B1").Formula = "=PHONETIC(A1:A2)"\nv = ws.Range("B1").Text',
    'ws.Range("A1").Formula = "=1+1"\nws.Range("B1").Formula = "=PHONETIC(A1)"\nv = ws.Range("B1").Text',
    'ws.Range("A1").Value = "abc"\nv = WorksheetFunction.Phonetic(ws.Range("A1"))',
    'Set o = Application.WorksheetFunction\nv = o.Phonetic("abc")',
    'ws.Range("B1").Formula = "=PHONETIC(""abc"")"\nv = ws.Range("B1").Text',
    'ws.Range("B1").Formula = "=PHONETIC(A1)"\nv = ws.Range("B1").Text',
]
PROBE = """Private Function Show(ByVal v As Variant) As String
    If IsObject(v) Then
        Show = "object"
    ElseIf IsArray(v) Then
        Show = "array"
    ElseIf IsError(v) Then
        Show = CStr(v)
    Else
        Show = CStr(v)
    End If
End Function

Public Function Missing(ByVal which As Long) As String
    Dim ws As Object, t As String, v As Variant, o As Object, n As Long
    Application.DisplayAlerts = False
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    On Error Resume Next
    Select Case which
@@CASES@@
    End Select
    n = Err.Number
    t = TypeName(v) & "=" & Show(v) & "|" & n
    Err.Clear
    ws.Parent.Close False
    Application.DisplayAlerts = True
    Missing = t
End Function
"""


def source() -> str:
    return PROBE.replace('@@CASES@@', '\n'.join(f'    Case {index}\n        {case}' for index, case in enumerate(CASES)))


def main() -> None:
    text = source()
    cases: list[list[str]] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, case in enumerate(CASES):
            result = excel.run_vba(text, 'Missing', (index,), timeout=60.0)
            assert result.ok and not result.dialogs, (case, result)
            cases.append([case, str(result.value)])
            print(index, case.replace('\n', '; ')[-70:], '=>', result.value, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': text, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
