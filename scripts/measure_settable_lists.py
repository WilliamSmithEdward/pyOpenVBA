"""Settable properties of ListObject, PageSetup and Name in live Excel, in the form of the audit of Worksheet, Workbook and
Application: each value set on a new workbook, late bound and typed, since a bad value's error differs between them. A case's
answer is "before|Eset|after|reopened": what the property reads before, the set's error, what it reads after, and for a typed
set what it reads once the workbook is saved and opened again. The first half of a row's cases is late bound.

    python scripts/measure_settable_lists.py

writes tests/fixtures/settable_lists.json, which scripts/compare_settable_lists.py and tests/test_excel_settable_lists.py replay.
A run stops after about eight minutes and the next takes up where it stopped. Each property takes its own Excel.
"""
from __future__ import annotations

import json
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/settable_lists.json'

BOOLEAN = ['True', 'False', '0', '1', '-1', '2', '1.5', '"True"', '"x"', '""', 'Null', 'Empty']
NUMBER = ['0', '1', '2', '3', '-1', '1.5', '99999', 'True', '"1"', '"x"', '""', 'Null', 'Empty']
TEXT = ['"abc"', '""', '5', '1.5', 'True', 'Null', 'Empty', '"' + 'x' * 40 + '"']
HEADER = [*TEXT, '"&L&Bx"', '"&Cmid&Rend"', '"&"', '"' + 'x' * 255 + '"', '"' + 'x' * 256 + '"']
MARGIN = [*NUMBER, '72', '36.5', '1000000']
FIT = ['1', '2', '0', '-1', '32767', '32768', 'False', 'True', '"x"', 'Null', 'Empty', '1.5']
TABLE_NAME = [*TEXT, '"a b"', '"1a"', '"A1"', '"R1C1"', '"Other"', '"_ok"', '"a.b"']
NAME_NAME = [*TEXT, '"a b"', '"1a"', '"A1"', '"R1C1"', '"Other"', '"_ok"', '"a.b"', '"c"', '"r"']
REFERS = ['"=Sheet1!$B$2"', '"=$B$2"', '"=B2"', '"=1+2"', '"abc"', '""', '5', 'True', 'Null', 'Empty', '"=Nope!A1"', '"=("']
REFERS_R1C1 = ['"=Sheet1!R2C2"', '"=R2C2"', '"=RC"', '"=1+2"', '"abc"', '""', '5', 'True', 'Null', 'Empty', '"=("']
LIST_OBJECT = {
    'AlternativeText': TEXT, 'Comment': TEXT, 'DisplayName': TABLE_NAME, 'Name': TABLE_NAME, 'ShowAutoFilter': BOOLEAN, 'ShowAutoFilterDropDown': BOOLEAN, 'ShowHeaders': BOOLEAN,
    'ShowTableStyleColumnStripes': BOOLEAN, 'ShowTableStyleFirstColumn': BOOLEAN, 'ShowTableStyleLastColumn': BOOLEAN, 'ShowTableStyleRowStripes': BOOLEAN, 'ShowTotals': BOOLEAN, 'Summary': TEXT,
    'TableStyle': ['"TableStyleLight1"', '"tablestylelight2"', '"TableStyleMedium2"', '""', '"Nope"', '5', 'True', 'Null', 'Empty'],
}
PAGE_SETUP = {
    'AlignMarginsHeaderFooter': BOOLEAN, 'BlackAndWhite': BOOLEAN, 'BottomMargin': MARGIN, 'CenterFooter': HEADER, 'CenterHeader': HEADER, 'CenterHorizontally': BOOLEAN, 'CenterVertically': BOOLEAN,
    'DifferentFirstPageHeaderFooter': BOOLEAN, 'Draft': BOOLEAN, 'FirstPageNumber': [*NUMBER, '-4105', '32767', '32768'], 'FitToPagesTall': FIT, 'FitToPagesWide': FIT, 'FooterMargin': MARGIN, 'HeaderMargin': MARGIN,
    'LeftFooter': HEADER, 'LeftHeader': HEADER, 'LeftMargin': MARGIN, 'OddAndEvenPagesHeaderFooter': BOOLEAN, 'Order': [*NUMBER, 'xlDownThenOver', 'xlOverThenDown'], 'Orientation': [*NUMBER, 'xlPortrait', 'xlLandscape'],
    'PaperSize': [*NUMBER, '9', '5', '256', '999'], 'PrintArea': ['"$A$1:$B$2"', '"A1"', '"A1:B2,D1"', '""', '"Other!A1"', '"Sheet2!A1"', '"abc"', '5', 'True', 'Null', 'Empty'],
    'PrintComments': [*NUMBER, '-4142', '16', '14'], 'PrintErrors': NUMBER, 'PrintGridlines': BOOLEAN, 'PrintHeadings': BOOLEAN, 'PrintNotes': BOOLEAN, 'PrintQuality': ['600', '300', '0', '-1', '"x"', 'Null', 'Empty', 'True'],
    'PrintTitleColumns': ['"$A:$B"', '"A:B"', '"A1"', '""', '"abc"', '5', 'True', 'Null', 'Empty', '"$1:$2"'], 'PrintTitleRows': ['"$1:$2"', '"1:2"', '"A1"', '""', '"abc"', '5', 'True', 'Null', 'Empty', '"$A:$B"'],
    'RightFooter': HEADER, 'RightHeader': HEADER, 'RightMargin': MARGIN, 'ScaleWithDocHeaderFooter': BOOLEAN, 'TopMargin': MARGIN, 'Zoom': ['100', '50', '10', '400', '9', '401', '75.5', 'False', 'True', '"x"', '"50"', 'Null', 'Empty', '0'],
}
NAME = {
    'Category': TEXT, 'CategoryLocal': TEXT, 'Comment': [*TEXT, '"' + 'x' * 255 + '"', '"' + 'x' * 256 + '"'], 'MacroType': [*NUMBER, '-4142'], 'Name': NAME_NAME, 'NameLocal': NAME_NAME, 'RefersTo': REFERS, 'RefersToLocal': REFERS,
    'RefersToR1C1': REFERS_R1C1, 'RefersToR1C1Local': REFERS_R1C1, 'ShortcutKey': ['"a"', '"A"', '"ab"', '""', '"1"', '5', 'True', 'Null', 'Empty'], 'Value': REFERS, 'Visible': BOOLEAN, 'WorkbookParameter': BOOLEAN,
}
OBJECTS = {'ListObject': LIST_OBJECT, 'PageSetup': PAGE_SETUP, 'Name': NAME}
TARGETS = {'ListObject': ('ws.ListObjects(1)', 'wb.Worksheets(1).ListObjects(1)'), 'PageSetup': ('ws.PageSetup', 'wb.Worksheets(1).PageSetup'), 'Name': ('wb.Names(1)', 'wb.Names(1)')}

TEMPLATE = '''Private Function Read1(o As Object) As String
    On Error GoTo Failed
    Read1 = TypeName(o.{prop}) & ":" & CStr(o.{prop})
    Exit Function
Failed:
    Read1 = "E" & Err.Number
End Function

Public Function T(ByVal which As Long, ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, o As Object, typed As {kind}, s As String, again As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    wb.Worksheets.Add(After:=ws).Name = "Sheet2"
    ws.Activate
    ws.Range("A1:B1").Value = Array("K", "V")
    ws.Range("A2:B2").Value = Array(1, 2)
    ws.Range("A3:B3").Value = Array(3, 4)
    ws.ListObjects.Add xlSrcRange, ws.Range("A1:B3"), , xlYes
    wb.Names.Add "Nm", "=Sheet1!$A$1"
    Set typed = {target}
    Set o = typed
    s = Read1(o)
    On Error Resume Next
    Err.Clear
    Select Case which
{cases}
    End Select
    s = s & "|E" & Err.Number & "|"
    On Error GoTo 0
    s = s & Read1(o) & "|"
    On Error Resume Next
    If which >= {half} Then
        Err.Clear
        wb.SaveAs folder & "\\saved.xlsx", xlOpenXMLWorkbook
        If Err.Number <> 0 Then
            s = s & "E" & Err.Number & " saving"
        Else
            wb.Close False
            Set wb = Workbooks.Open(folder & "\\saved.xlsx", 0)
            Set typed = {reopened}
            Set o = typed
            s = s & Read1(o)
        End If
    End If
    wb.Close False
    Kill folder & "\\saved.xlsx"
    T = s
End Function

Public Function RunAll(ByVal last As Long, ByVal folder As String) As String
    Dim i As Long, s As String
    For i = 0 To last
        s = s & "@@@" & T(i, folder)
    Next
    RunAll = Mid$(s, 4)
End Function
'''


def source(obj: str, prop: str, values: list[str]) -> str:
    late = [f'    Case {index}\n        o.{prop} = {value}' for index, value in enumerate(values)]
    typed = [f'    Case {index + len(values)}\n        typed.{prop} = {value}' for index, value in enumerate(values)]
    target, reopened = TARGETS[obj]
    return TEMPLATE.format(prop=prop, kind=obj, target=target, reopened=reopened, half=len(values), cases='\n'.join([*late, *typed]))


def main() -> None:
    data: dict[str, dict[str, object]] = json.loads(OUT.read_text(encoding='utf-8'))['rows'] if OUT.exists() else {}
    ends = time.monotonic() + 480.0
    for obj, properties in OBJECTS.items():
        for prop, values in properties.items():
            name = f'{obj}.{prop}'
            if name in data or time.monotonic() >= ends:
                continue
            text = source(obj, prop, values)
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
                excel.new_document()
                result = excel.run_vba(text, 'RunAll', (2 * len(values) - 1, folder), timeout=170.0)
                answers = str(result.value).split('@@@') if result.ok else [f'failed {result.outcome}'] * 2 * len(values)
            data[name] = {'object': obj, 'property': prop, 'values': values, 'source': text, 'answers': answers}
            OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'rows': data}, indent=1) + '\n', encoding='utf-8')
            print(name, answers[0][:90], flush=True)
    print(len(data), 'of', sum(len(properties) for properties in OBJECTS.values()), flush=True)


if __name__ == '__main__':
    main()
