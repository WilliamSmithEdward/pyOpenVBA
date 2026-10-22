"""What the 1904 date system does to the features that read and write dates.

tests/fixtures/date1904/native.json covers cells, typing and the setting
itself. These probes cover the paths that turn a date into a serial or a
serial into text on their own: Text cells, the formula-bar text Find and
Replace work on, AutoFill's date series, TextToColumns, AutoFilter and
validation criteria, and formulas and worksheet functions. Each probe
makes its own 1904 workbook and closes it.

    python scripts/measure_date1904_paths.py

writes tests/fixtures/date1904/paths.json, which tests/test_excel_date1904.py
replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/date1904/paths.json'

#: Helpers every probe shares: a cell as Value's type and text, Value2, Text and NumberFormat, or an error number.
HELPERS = '''
Function D(c As Range) As String
D = TypeName(c.Value) & ":" & CStr(c.Value) & "~" & CStr(c.Value2) & "~" & c.Text & "~" & c.NumberFormat
End Function
Function W(c As Range, v As Variant) As String
On Error Resume Next
c.Value = v
If Err.Number <> 0 Then W = "e" & Err.Number Else W = D(c)
End Function
Function Cells1(ws As Worksheet, a As String) As String
Dim c As Range
For Each c In ws.Range(a).Cells
Cells1 = Cells1 & "|" & c.Address(False, False) & "=" & D(c)
Next
End Function
'''
OPEN = ('Dim wb As Workbook, ws As Worksheet, s As String\nSet wb = Workbooks.Add(xlWBATWorksheet)\n'
        'wb.Date1904 = True\nSet ws = wb.Worksheets(1)\nws.Columns("A:H").ColumnWidth = 30\n')
CLOSE = '\nwb.Close False\nProbe = s'

CASES = {
    # A Date written to a Text cell becomes the text of the serial the workbook keeps for it.
    'text_cells': '''ws.Range("A1:A6").NumberFormat = "@"
s = W(ws.Range("A1"), #2/29/2024#)
s = s & "|" & W(ws.Range("A2"), #1/2/1904#)
s = s & "|" & W(ws.Range("A3"), #1/1/1904 6:00:00 AM#)
s = s & "|" & W(ws.Range("A4"), #2/29/2024 1:30:00 PM#)
s = s & "|" & W(ws.Range("A5"), #12/31/1903#)
s = s & "|" & W(ws.Range("A6"), CDate(0.75))''',
    # Range.Formula of a constant is the formula-bar text.
    'formula_text': '''ws.Range("A1").Value = #2/29/2024#
ws.Range("A2").Value = #2/29/2024 1:30:00 PM#
ws.Range("A3").NumberFormat = "h:mm": ws.Range("A3").Value2 = 0.25
ws.Range("A4").NumberFormat = "yyyy-mm-dd": ws.Range("A4").Value2 = 1
ws.Range("A5").NumberFormat = "yyyy-mm-dd": ws.Range("A5").Value2 = -1
ws.Range("A6").NumberFormat = "yyyy-mm-dd": ws.Range("A6").Value2 = 2957003
ws.Range("A7").NumberFormat = "yyyy-mm-dd": ws.Range("A7").Value2 = 2957004
ws.Range("A8").NumberFormat = "yyyy-mm-dd hh:mm": ws.Range("A8").Value2 = 1.5
Dim i As Long
For i = 1 To 8
s = s & "|" & ws.Cells(i, 1).Formula & "~" & ws.Cells(i, 1).FormulaR1C1 & "~" & ws.Cells(i, 1).Text
Next''',
    'find_replace': '''ws.Range("A1").Value = #2/29/2024#
ws.Range("A2").Value = #1/1/1905#
ws.Range("A3").Value = #3/1/2024 6:00:00 AM#
ws.Range("A4").Value = #1/2/1904#
Dim f As Range
Set f = ws.Range("A1:A4").Find("2/29/2024", LookIn:=xlFormulas, LookAt:=xlWhole)
If f Is Nothing Then s = "find=none" Else s = "find=" & f.Address(False, False)
Set f = ws.Range("A1:A4").Find("1/2/1904", LookIn:=xlFormulas, LookAt:=xlWhole)
If f Is Nothing Then s = s & ";find2=none" Else s = s & ";find2=" & f.Address(False, False)
ws.Range("A1").Replace "2024", "2025", LookAt:=xlPart
ws.Range("A2").Replace "1905", "1900", LookAt:=xlPart
ws.Range("A3").Replace "6:00:00", "7:00:00", LookAt:=xlPart
ws.Range("A4").Replace "1904", "1903", LookAt:=xlPart
s = s & Cells1(ws, "A1:A4")''',
    'autofill': '''ws.Range("A1").Value = #1/31/2024#: ws.Range("A2").Value = #2/29/2024#
ws.Range("A1:A2").AutoFill ws.Range("A1:A6")
ws.Range("B1").Value = #1/1/2024#
ws.Range("B1").AutoFill ws.Range("B1:B5"), xlFillMonths
ws.Range("C1").Value = #1/5/2024#
ws.Range("C1").AutoFill ws.Range("C1:C6"), xlFillWeekdays
ws.Range("D1").Value = #2/29/2024#
ws.Range("D1").AutoFill ws.Range("D1:D4"), xlFillYears
ws.Range("E1").NumberFormat = "m/d/yyyy": ws.Range("E1").Value2 = 1
ws.Range("E1").AutoFill ws.Range("E1:E4"), xlFillMonths
ws.Range("F1").Value = #1/1/2024#
ws.Range("F1").AutoFill ws.Range("F1:F3")
ws.Range("G1").NumberFormat = "m/d/yyyy": ws.Range("G1").Value2 = 0
ws.Range("G1").AutoFill ws.Range("G1:G3"), xlFillMonths
ws.Range("H1").NumberFormat = "m/d/yyyy": ws.Range("H1").Value2 = 59
ws.Range("H1").AutoFill ws.Range("H1:H3"), xlFillMonths
s = Cells1(ws, "A1:A6") & Cells1(ws, "B1:B5") & Cells1(ws, "C1:C6") & Cells1(ws, "D1:D4") & Cells1(ws, "E1:E4") _
    & Cells1(ws, "F1:F3") & Cells1(ws, "G1:G3") & Cells1(ws, "H1:H3")''',
    'text_to_columns': '''ws.Range("A1:A7").NumberFormat = "@"
ws.Range("A1").Value = "2/29/2024": ws.Range("A2").Value = "1/1/1900": ws.Range("A3").Value = "1/1/1904"
ws.Range("A4").Value = "2/29/2024 12:30": ws.Range("A5").Value = "12:30": ws.Range("A6").Value = "1/2/1904"
ws.Range("A7").Value = "Feb 2024"
ws.Range("A1:A7").TextToColumns Destination:=ws.Range("B1"), DataType:=xlDelimited, Tab:=True
ws.Range("A1:A7").TextToColumns Destination:=ws.Range("C1"), DataType:=xlDelimited, Tab:=True, FieldInfo:=Array(1, xlMDYFormat)
s = Cells1(ws, "B1:B7") & Cells1(ws, "C1:C7")''',
    'autofilter': '''ws.Range("A1").Value = "When"
ws.Range("A2").Value = #1/1/2024#: ws.Range("A3").Value = #2/1/2024#
ws.Range("A4").Value = #3/1/2024#: ws.Range("A5").Value = #4/1/2024#
ws.Range("A1:A5").AutoFilter 1, ">2/15/2024"
Dim i As Long
For i = 2 To 5
s = s & IIf(ws.Rows(i).Hidden, "h", "v")
Next
s = s & ";" & ws.AutoFilter.Filters(1).Criteria1
ws.Range("A1:A5").AutoFilter 1, "2/1/2024"
s = s & ";"
For i = 2 To 5
s = s & IIf(ws.Rows(i).Hidden, "h", "v")
Next
s = s & ";" & ws.AutoFilter.Filters(1).Criteria1''',
    'validation': '''With ws.Range("B1").Validation
.Add xlValidateDate, xlValidAlertStop, xlBetween, "1/1/2024", "12/31/2024"
s = .Formula1 & "~" & .Formula2
End With
ws.Range("B1").Value = #6/1/2024#
s = s & ";" & ws.Range("B1").Validation.Value
ws.Range("B1").Value = #6/1/2025#
s = s & ";" & ws.Range("B1").Validation.Value
With ws.Range("B2").Validation
.Add xlValidateDate, xlValidAlertStop, xlGreater, "=DATE(2024,1,1)"
s = s & ";" & .Formula1
End With
ws.Range("B2").Value = #6/1/2024#
s = s & ";" & ws.Range("B2").Validation.Value''',
    'formulas': '''ws.Range("A1").Formula = "=""2/29/2024""+0"
ws.Range("A2").Formula = "=DATEVALUE(""1/1/1904"")"
ws.Range("A3").Formula = "=DATE(1904,1,1)"
ws.Range("A4").Formula = "=DATE(1903,12,31)"
ws.Range("A5").Formula = "=TEXT(0,""yyyy-mm-dd ddd"")"
ws.Range("A6").Formula = "=YEAR(-1)"
ws.Range("A7").Formula = "=WEEKDAY(0)"
ws.Range("A8").Formula = "=DATE(2024,2,29)-DATE(1904,1,2)"
ws.Range("A9").Formula = "=EDATE(0,1)"
ws.Range("A10").Formula = "=DAYS360(0,365)"
ws.Range("A11").Formula = "=TIMEVALUE(""1/2/1904 6:00"")"
ws.Range("A12").Formula = "=VALUE(""1/2/1904"")"
s = Cells1(ws, "A1:A12")''',
    # Worksheet functions and Evaluate with the 1904 workbook active, then with a 1900 one active.
    'worksheet_functions': '''Dim other As Workbook, pass As Long
For pass = 1 To 2
If pass = 2 Then Set other = Workbooks.Add(xlWBATWorksheet): other.Activate Else wb.Activate
s = s & "|" & WorksheetFunction.Weekday(0) & "~" & WorksheetFunction.Text(1, "yyyy-mm-dd")
s = s & "~" & Application.Evaluate("YEAR(0)") & "~" & ws.Evaluate("YEAR(0)") & "~" & Application.Evaluate("DATE(2024,2,29)")
s = s & "~" & ws.Evaluate("DATE(2024,2,29)") & "~" & TypeName(Application.Evaluate("DATE(2024,2,29)"))
s = s & "~" & WorksheetFunction.EDate(0, 1) & "~" & Application.Text(0, "yyyy")
Next
other.Close False''',
    # Which text Find and Replace see in a date cell: the date or the serial Range.Formula gives.
    'find_replace_serials': '''Dim f As Range, pass As Long, book As Workbook, sheet As Worksheet
For pass = 1 To 2
If pass = 1 Then Set book = wb Else Set book = Workbooks.Add(xlWBATWorksheet)
Set sheet = book.Worksheets(1)
sheet.Columns("A").ColumnWidth = 30
sheet.Range("A1").Value = #2/29/2024#: sheet.Range("A2").Value = #2/29/2024#
sheet.Range("A3").Value = #2/29/2024#: sheet.Range("A4").Value = #2/29/2024 1:30:00 PM#
s = s & "|" & sheet.Range("A1").Formula
Set f = sheet.Range("A1:A4").Find(CStr(sheet.Range("A1").Value2), LookIn:=xlFormulas, LookAt:=xlWhole)
If f Is Nothing Then s = s & "~none" Else s = s & "~" & f.Address(False, False)
Set f = sheet.Range("A1:A4").Find("2/29", LookIn:=xlFormulas, LookAt:=xlPart)
If f Is Nothing Then s = s & "~none" Else s = s & "~" & f.Address(False, False)
Set f = sheet.Range("A1:A4").Find("2/29/2024", LookIn:=xlValues, LookAt:=xlWhole)
If f Is Nothing Then s = s & "~none" Else s = s & "~" & f.Address(False, False)
sheet.Range("A1").Replace CStr(sheet.Range("A1").Value2), CStr(sheet.Range("A1").Value2 + 1), LookAt:=xlPart
sheet.Range("A2").Replace "29", "28", LookAt:=xlPart
sheet.Range("A3").Replace "/", "-", LookAt:=xlPart
sheet.Range("A4").Replace "1:30", "2:30", LookAt:=xlPart
s = s & Cells1(sheet, "A1:A4")
If pass = 2 Then book.Close False
Next''',
    # A date column reads a colon as it reads a slash; a field it reads no date in is typed as a General one.
    'text_to_columns_colons': '''Dim pass As Long, book As Workbook, sheet As Worksheet, i As Long, f As Variant, fields As Variant
fields = Array("12:30", "1:2:2024", "12:30:45", "12:30 PM", "1:30", "13:45", "12:30:2020", "1:2", "3:4:5", "12::30", _
    "12:30.5")
For pass = 1 To 2
If pass = 1 Then Set book = wb Else Set book = Workbooks.Add(xlWBATWorksheet)
Set sheet = book.Worksheets(1)
sheet.Range("A1:A11").NumberFormat = "@"
For i = 0 To 10
sheet.Cells(i + 1, 1).Value = fields(i)
Next
sheet.Range("A1:A11").TextToColumns Destination:=sheet.Range("B1"), DataType:=xlDelimited, Tab:=True, _
    FieldInfo:=Array(Array(1, xlMDYFormat))
sheet.Range("A1:A11").TextToColumns Destination:=sheet.Range("C1"), DataType:=xlDelimited, Tab:=True, _
    FieldInfo:=Array(Array(1, xlDMYFormat))
sheet.Range("A1:A11").TextToColumns Destination:=sheet.Range("D1"), DataType:=xlDelimited, Tab:=True, _
    FieldInfo:=Array(Array(1, xlYMDFormat))
s = s & "#" & book.Date1904
For i = 1 To 11
For Each f In Array(2, 3, 4)
s = s & "|" & TypeName(sheet.Cells(i, f).Value) & ":" & sheet.Cells(i, f).Value2 & "~" & sheet.Cells(i, f).NumberFormat
Next
Next
If pass = 2 Then book.Close False
Next''',
    # The year a date read without one falls in, which the tests hold typing to.
    'year': 's = CStr(Year(Date))',
    # One date filled by months or years where it ends its month, on both calendars.
    'autofill_month_ends': '''Dim pass As Long, book As Workbook, sheet As Worksheet, i As Long, starts As Variant, kinds As Variant
starts = Array(#2/29/2024#, #4/30/2024#, #2/28/2023#, #1/30/2024#, #11/30/2024#, #1/31/2024#)
kinds = Array(xlFillMonths, xlFillYears, xlFillDefault)
For pass = 1 To 2
If pass = 1 Then Set book = wb Else Set book = Workbooks.Add(xlWBATWorksheet)
Set sheet = book.Worksheets(1)
For i = 0 To 5
sheet.Cells(1, i + 1).Value = starts(i)
sheet.Cells(1, i + 1).AutoFill sheet.Range(sheet.Cells(1, i + 1), sheet.Cells(4, i + 1)), xlFillMonths
sheet.Cells(6, i + 1).Value = starts(i)
sheet.Cells(6, i + 1).AutoFill sheet.Range(sheet.Cells(6, i + 1), sheet.Cells(8, i + 1)), xlFillYears
sheet.Cells(10, i + 1).Value = starts(i)
sheet.Cells(10, i + 1).NumberFormat = "mmm-yy"
sheet.Cells(10, i + 1).AutoFill sheet.Range(sheet.Cells(10, i + 1), sheet.Cells(12, i + 1))
Next
s = s & "#" & book.Date1904
For i = 1 To 6
s = s & "|" & sheet.Cells(2, i).Value2 & "~" & sheet.Cells(3, i).Value2 & "~" & sheet.Cells(4, i).Value2 _
    & "~" & sheet.Cells(7, i).Value2 & "~" & sheet.Cells(8, i).Value2 & "~" & sheet.Cells(11, i).Value2 & "~" & sheet.Cells(12, i).Value2
Next
If pass = 2 Then book.Close False
Next''',
}


#: Module-level code a probe needs beside the helpers: worksheet functions written in VBA.
UDFS = {
    'udf_dates': '''Public Function DateBack() As Variant
DateBack = #2/29/2024#
End Function
Public Function DateEcho(v As Variant) As String
DateEcho = TypeName(v) & ":" & CStr(v)
End Function
''',
}
# A VBA Date crossing into a 1904 workbook: a user-defined function's result and arguments, and
# WorksheetFunction given one. The functions live in the probe's own workbook, set to 1904 for the probe.
CASES['udf_dates'] = '''Dim host As Worksheet
Set host = ThisWorkbook.Worksheets(1)
ThisWorkbook.Date1904 = True
host.Columns("A").ColumnWidth = 30
host.Range("A1").Formula = "=DateBack()"
host.Range("A2").Value = #2/29/2024#
host.Range("A3").Formula = "=DateEcho(A2)"
host.Range("A4").Formula = "=DateEcho(A2+0)"
host.Range("A5").Formula = "=DateEcho(DATE(2024,2,29))"
host.Range("A6").Formula = "=DateBack()+0"
ThisWorkbook.Activate
s = host.Range("A1").Value2 & "~" & host.Range("A1").Text & "~" & host.Range("A1").NumberFormat
s = s & "|" & host.Range("A3").Value & "|" & host.Range("A4").Value & "|" & host.Range("A5").Value & "|" & host.Range("A6").Value2
s = s & "|" & WorksheetFunction.Weekday(#2/29/2024#) & "|" & WorksheetFunction.Text(#2/29/2024#, "yyyy-mm-dd")
s = s & "|" & WorksheetFunction.Max(#2/29/2024#, 1) & "|" & host.Evaluate("A2+0")
host.Range("A1:A6").Clear
ThisWorkbook.Date1904 = False'''


def macro(name: str) -> str:
    return HELPERS + UDFS.get(name, '') + 'Function Probe() As String\n' + OPEN + CASES[name] + CLOSE + '\nEnd Function\n'


if __name__ == '__main__':
    record({name: macro(name) for name in CASES}, OUT)
