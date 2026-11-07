"""A seventeenth sweep: the VBA language itself, as everyday macros use it, read line by line from live Excel.

Each line is an expression read under On Error, with the error, the type and the value recorded. A line that does not compile is
found by halving and recorded, as the rest still answer.

    python scripts/measure_vba_everyday_17.py

writes tests/fixtures/vba_everyday_17.json, keeps each group as it finishes and asks only the rest.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_everyday_17.json'
SEPARATOR = '¦'
GROUPS: dict[str, list[str]] = {
    'conversions': [
        'CInt(2.5)', 'CInt(3.5)', 'CInt(-2.5)', 'CInt("12")', 'CInt("1e2")', 'CInt("&H10")', 'CInt("abc")', 'CInt(32767.5)', 'CInt(True)', 'CInt(Null)', 'CInt(Empty)', 'CInt("")', 'CLng(2147483647.5)', 'CLng("1,000")', 'CLng("$5")', 'CDbl("1.5e3")', 'CDbl("(5)")', 'CDbl("5-")', 'CDbl(" 5 ")', 'CDbl("%5")', 'CDbl("5%")', 'CSng(1.23456789)', 'CCur(1.23456)', 'CCur("1.5")', 'CDec(1.5)', 'CDec("1.23456789012345678901")', 'CByte(255)', 'CByte(256)', 'CByte(-1)', 'CByte(1.5)', 'CBool(0)', 'CBool(2)', 'CBool("True")', 'CBool("false")', 'CBool("yes")', 'CBool("1")', 'CBool("")', 'CBool(Null)', 'CStr(1.5)', 'CStr(True)', 'CStr(Empty)', 'CStr(Null)', 'CStr(#1/2/2020#)', 'CStr(1E+20)', 'CStr(0.1 + 0.2)', 'CStr(1 / 3)', 'CStr(123456789012345#)', 'CDate("1/2/2020")', 'CDate(43832)', 'CDate("abc")', 'CDate("2020-01-02 03:04")', 'CDate(0)', 'CDate(-1)', 'CVar(5)', 'CVErr(2042)', 'Val("12abc")', 'Val("  1 2 ")', 'Val("&H10")', 'Val("1e3")', 'Val("abc")', 'Val("")', 'Val(".5")', 'Val("-.5e-1")', 'Int(-2.5)', 'Fix(-2.5)', 'Int(2.5)', 'Round(2.5)', 'Round(3.5)', 'Round(2.567, 2)', 'Round(-2.5)', 'Round(1.005, 2)', 'Abs(-1.5)', 'Sgn(-3)', 'Sqr(16)', 'Sqr(-1)', 'Log(0)', 'Exp(1)', 'Hex(255)', 'Hex(-1)', 'Oct(8)', 'Asc("A")', 'Asc("")', 'AscW("A")', 'Chr(65)', 'Chr(256)', 'ChrW(8364)', 'Str(5)', 'Str(-5)', 'Str(1.5)', 'IsNumeric("1,000")', 'IsNumeric("$5")', 'IsNumeric("abc")', 'IsNumeric("")', 'IsNumeric(Empty)', 'IsNumeric(Null)', 'IsNumeric(True)', 'IsDate("1/2/2020")', 'IsDate("abc")', 'IsDate(5)', 'IsEmpty(Empty)', 'IsEmpty("")', 'IsNull(Null)', 'IsNull(Empty)', 'IsArray(Array())', 'IsObject(Nothing)', 'IsError(CVErr(1))', 'IsMissing(1)', 'TypeName(1)', 'TypeName(1&)', 'TypeName(1#)', 'TypeName(1!)', 'TypeName(1@)', 'TypeName("a")', 'TypeName(True)', 'TypeName(Empty)', 'TypeName(Null)', 'TypeName(Nothing)', 'TypeName(#1/1/2000#)', 'TypeName(Array())', 'TypeName(CDec(1))', 'TypeName(CByte(1))', 'TypeName(CVErr(1))', 'VarType(1)', 'VarType("a")', 'VarType(Array())', 'VarType(Empty)', 'VarType(Null)', 'VarType(Nothing)',
    ],
    'operators': [
        '1 + 1', '1 + "1"', '"1" + "1"', '"a" + 1', '1 & 1', 'True + True', 'Null + 1', 'Null & "a"', 'Empty + 1', 'Empty & "a"', '5 / 2', '5 \\ 2', '-5 \\ 2', '5.5 \\ 2', '5 Mod 3', '-5 Mod 3', '5.5 Mod 2', '5 / 0', '5 \\ 0', '5 Mod 0', '0 / 0', '2 ^ 3', '-2 ^ 2', '2 ^ -1', '0 ^ 0', '(-8) ^ (1 / 3)', '32767 + 1', '32767 * 2', '32767& + 1', '2147483647 + 1', '200 * 200', '1E+308 * 10', 'CByte(200) + CByte(100)', 'Not 0', 'Not True', 'Not 5', 'True And 5', '5 And 3', '5 Or 3', '5 Xor 3', 'True Eqv False', 'True Imp False', 'Null And False', 'Null Or True', 'Not Null', '1 = 1', '1 = "1"', '"a" = "A"', '"a" < "B"', '"10" < "9"', '10 < "9"', 'Null = Null', 'Empty = 0', 'Empty = ""', 'True = -1', 'True = 1', '"abc" Like "a*"', '"abc" Like "A*"', '"a1" Like "a#"', '"b" Like "[a-c]"', '"b" Like "[!a-c]"', '"a*" Like "a[*]"', '"" Like ""', 'Null Like "a"', '#1/2/2020# + 1', '#1/2/2020# - #1/1/2020#', '#1/2/2020# > 1', 'TypeName(#1/2/2020# + 1)', 'TypeName(#1/2/2020# - #1/1/2020#)', 'TypeName(1 + 1.5)', 'TypeName(1 + 1&)', 'TypeName(1 / 1)', 'TypeName(1 \\ 1)', 'TypeName(1& \\ 1)', 'TypeName(1.5 \\ 1)', 'TypeName(2 ^ 2)', 'TypeName(1@ + 1)', 'TypeName(1! + 1)', 'TypeName(True + True)', 'TypeName(CByte(1) + CByte(1))', 'TypeName("1" + 1)', 'TypeName(Empty + Empty)', 'TypeName(1 & 1)', 'TypeName(Null + 1)',
    ],
    'strings': [
        'Len("abc")', 'Len(Null)', 'Len(Empty)', 'Len(5)', 'Len(5&)', 'Len(1.5)', 'LenB("abc")', 'Left("abc", 2)', 'Left("abc", 0)', 'Left("abc", -1)', 'Left("abc", 9)', 'Left(Null, 1)', 'Right("abc", 2)', 'Right("abc", 9)', 'Mid("abcde", 2, 2)', 'Mid("abcde", 2)', 'Mid("abcde", 9)', 'Mid("abcde", 0)', 'Mid("abcde", 2, -1)', 'Mid(Null, 1)', 'InStr("abc", "b")', 'InStr(2, "abcb", "b")', 'InStr("abc", "")', 'InStr("", "a")', 'InStr(9, "abc", "b")', 'InStr(0, "abc", "b")', 'InStr(1, "aBc", "b", vbTextCompare)', 'InStr(1, "aBc", "b", vbBinaryCompare)', 'InStr(Null, "a")', 'InStr("a", Null)', 'InStrRev("abcb", "b")', 'InStrRev("abcb", "b", 3)', 'InStrRev("abc", "")', 'Replace("aXbX", "X", "Y")', 'Replace("aXbX", "X", "Y", 3)', 'Replace("aXbX", "X", "Y", 1, 1)', 'Replace("aXbX", "x", "Y", , , vbTextCompare)', 'Replace("abc", "", "Y")', 'Replace("", "a", "Y")', 'Replace(Null, "a", "b")', 'Replace("abc", "b", Null)', 'Replace("abc", "b", "c", 9)', 'Replace("abc", "b", "c", 0)', 'Trim("  a  ")', 'LTrim("  a")', 'RTrim("a  ")', 'Trim(Null)', 'Trim(vbTab & "a")', 'UCase("aB")', 'LCase("aB")', 'UCase(Null)', 'StrConv("ab cd", vbProperCase)', 'StrConv("ab", vbUpperCase)', 'StrConv("a-b c", 3)', 'StrComp("a", "A")', 'StrComp("a", "A", vbTextCompare)', 'StrComp("a", "b")', 'StrComp("b", "a")', 'StrComp(Null, "a")', 'StrReverse("abc")', 'StrReverse("")', 'Space(3)', 'Space(0)', 'Space(-1)', 'String(3, "a")', 'String(3, 65)', 'String(3, "abc")', 'String(0, "a")', 'String(-1, "a")', 'Join(Array("a", "b"))', 'Join(Array("a", "b"), "-")', 'Join(Array())', 'Join(Array(1, 2), "")', 'Join(Array(1, Null), "-")', 'Join(Split("a,b", ","), "|")', 'UBound(Split("a,b", ","))', 'UBound(Split("", ","))', 'LBound(Split("a", ","))', 'UBound(Split("abc", ""))', 'UBound(Split("a,b,c", ",", 2))', 'Split("a,b,c", ",", 2)(1)', 'Split("a b")(1)', 'UBound(Split("a,b", ",", 0))', 'UBound(Split(Null))', 'UBound(Filter(Array("ab", "c"), "a"))', 'Filter(Array("ab", "c"), "a", False)(0)', 'UBound(Filter(Array("ab"), "z"))', 'FormatNumber(1234.567)', 'FormatNumber(1234.567, 1)', 'FormatNumber(0.5, 0)', 'FormatNumber(-1, 2, , vbTrue)', 'FormatNumber(1234.5, 2, , , vbFalse)', 'FormatPercent(0.5)', 'FormatPercent(0.123, 1)', 'FormatCurrency(1.5)', 'FormatDateTime(#1/2/2020#)', 'FormatDateTime(#1/2/2020#, vbLongDate)', 'FormatDateTime(#1:02:03 PM#, vbShortTime)', 'Format(1.5)', 'Format(1.5, "0.00")', 'Format(0.5, "0%")', 'Format(1234.5, "#,##0.0")', 'Format("abc", "@@@@")', 'Format("abc", ">")', 'Format(#1/2/2020#, "yyyy-mm-dd")', 'Format(#1/2/2020#, "ddd mmm")', 'Format(1, "Yes/No")', 'Format(0, "True/False")', 'Format(1, "On/Off")', 'Format(Null, "0")', 'Format(Empty, "0")', 'Format("", "0")', 'Format(1.5, "")', 'Format(-1, "0;(0)")', 'Format(0, "0;(0);zero")', 'Format(5, "000")', 'Format(1E+20, "0")', 'Format(1234.5, "Currency")', 'Format(0.5, "Percent")', 'Format(1234.5, "Standard")', 'Format(1234.5, "Fixed")', 'Format(1234.5, "Scientific")', 'Format(#1/2/2020#, "Short Date")', 'Format(#1/2/2020#, "Long Date")', 'Format(#1/2/2020 13:04#, "Short Time")', 'Format(#1/2/2020 13:04#, "h:nn am/pm")',
    ],
    'dates': [
        'DateSerial(2020, 2, 30)', 'DateSerial(2020, 13, 1)', 'DateSerial(20, 1, 1)', 'DateSerial(2020, 0, 0)', 'DateSerial(99, 1, 1)', 'DateSerial(100, 1, 1)', 'DateSerial(10000, 1, 1)', 'TimeSerial(25, 0, 0)', 'TimeSerial(1, 90, 0)', 'TimeSerial(0, 0, -1)', 'DateValue("1/2/2020")', 'DateValue("1/2/2020 3:04")', 'DateValue("abc")', 'TimeValue("3:04 PM")', 'TimeValue("1/2/2020 3:04")', 'Year(#1/2/2020#)', 'Month(#1/2/2020#)', 'Day(#1/2/2020#)', 'Hour(#3:04:05 PM#)', 'Minute(#3:04:05 PM#)', 'Second(#3:04:05 PM#)', 'Weekday(#1/2/2020#)', 'Weekday(#1/2/2020#, vbMonday)', 'Year(0)', 'Year("1/2/2020")', 'Year("abc")', 'Year(Null)', 'MonthName(1)', 'MonthName(1, True)', 'MonthName(13)', 'WeekdayName(1)', 'WeekdayName(1, True)', 'WeekdayName(1, , vbMonday)', 'WeekdayName(8)', 'DateAdd("d", 1, #1/31/2020#)', 'DateAdd("m", 1, #1/31/2020#)', 'DateAdd("yyyy", 1, #2/29/2020#)', 'DateAdd("q", 1, #1/31/2020#)', 'DateAdd("ww", 1, #1/1/2020#)', 'DateAdd("h", 1, #1/1/2020#)', 'DateAdd("n", 1, #1/1/2020#)', 'DateAdd("s", 1, #1/1/2020#)', 'DateAdd("w", 1, #1/1/2020#)', 'DateAdd("y", 1, #1/1/2020#)', 'DateAdd("x", 1, #1/1/2020#)', 'DateAdd("d", 1.7, #1/1/2020#)', 'DateAdd("d", 1, "1/1/2020")', 'DateAdd("yyyy", 9000, #1/1/2020#)', 'DateDiff("d", #1/1/2020#, #1/3/2020#)', 'DateDiff("d", #1/3/2020#, #1/1/2020#)', 'DateDiff("m", #1/31/2020#, #2/1/2020#)', 'DateDiff("yyyy", #12/31/2019#, #1/1/2020#)', 'DateDiff("q", #1/1/2020#, #4/1/2020#)', 'DateDiff("ww", #1/1/2020#, #1/15/2020#)', 'DateDiff("w", #1/1/2020#, #1/15/2020#)', 'DateDiff("h", #1/1/2020#, #1/2/2020#)', 'DateDiff("n", #1/1/2020#, #1/1/2020 1:00#)', 'DateDiff("s", #1/1/2020#, #1/1/2020 1:00#)', 'DateDiff("ww", #1/1/2020#, #1/15/2020#, vbMonday)', 'DateDiff("x", #1/1/2020#, #1/1/2020#)', 'DatePart("d", #1/2/2020#)', 'DatePart("m", #1/2/2020#)', 'DatePart("q", #5/2/2020#)', 'DatePart("yyyy", #1/2/2020#)', 'DatePart("y", #2/2/2020#)', 'DatePart("w", #1/2/2020#)', 'DatePart("ww", #1/2/2020#)', 'DatePart("ww", #1/2/2020#, vbMonday, vbFirstFullWeek)', 'DatePart("h", #1/2/2020 3:00#)', 'DatePart("x", #1/2/2020#)', 'TypeName(Date)', 'TypeName(Time)', 'TypeName(Now)', 'TypeName(Timer)', 'CDbl(#1/2/2020#)', 'CDbl(#12:00:00 PM#)', 'CLng(#1/2/2020#)', '#1/2/2020# = 43832', 'IsDate("2/30/2020")', 'IsDate("13/1/2020")', 'CDate("13/1/2020")', 'CDate("2/30/2020")', 'CDate("Jan 2 2020")', 'CDate("2 January 2020")', 'CDate("3 PM")', 'CDate("15:04")',
    ],
    'arrays': [
        'UBound(Array(1, 2))', 'LBound(Array(1, 2))', 'UBound(Array())', 'Do1(1)', 'A1(0)', 'A1(3)', 'A1(4)', 'A1(-1)', 'A1(1.5)', 'A1(2.5)', 'A1("1")', 'A1(Null)', 'A2()', 'A3()', 'A4()', 'A5()', 'A6()', 'A7()', 'A8()', 'A9()', 'A10()', 'A11()', 'A12()', 'A13()', 'A14()', 'A15()', 'A16()', 'A17()', 'A18()', 'A19()',
    ],
    'flow': [
        'F1()', 'F2()', 'F3()', 'F4()', 'F5()', 'F6()', 'F7()', 'F8()', 'F9()', 'F10()', 'F11()', 'F12()', 'F13()', 'F14()', 'F15()', 'F16()', 'F17()', 'F18()', 'F19()', 'F20()', 'IIf(True, 1, 2)', 'IIf(Null, 1, 2)', 'IIf(False, 1 / 0, 2)', 'Choose(1, "a", "b")', 'Choose(3, "a", "b")', 'Choose(0, "a")', 'Choose(1.9, "a", "b")', 'Switch(False, 1, True, 2)', 'Switch(False, 1)', 'Switch(True)', 'Err.Number', 'Err.Description', 'Error(5)', 'Error(11)', 'Error(0)', 'Error(99999)',
    ],
    'collections': [
        'C1()', 'C2()', 'C3()', 'C4()', 'C5()', 'C6()', 'C7()', 'C8()', 'C9()', 'C10()', 'C11()', 'C12()', 'C13()', 'C14()', 'C15()', 'C16()', 'C17()', 'C18()',
    ],
}
HELPERS = r'''Private Function Do1(ByVal x As Variant) As String
    Do1 = "ran"
End Function

Private Function A1(ByVal i As Variant) As String
    Dim a(1 To 3) As Long
    a(1) = 10: a(2) = 20: a(3) = 30
    A1 = a(i)
End Function

Private Function A2() As String
    Dim a() As Long
    ReDim a(1 To 2)
    a(2) = 5
    ReDim Preserve a(1 To 3)
    A2 = a(2) & "," & a(3) & "," & UBound(a)
End Function

Private Function A3() As String
    Dim a() As Long
    A3 = UBound(a)
End Function

Private Function A4() As String
    Dim a(1, 2) As Long
    A4 = LBound(a) & UBound(a) & UBound(a, 2) & LBound(a, 2)
    A4 = A4 & UBound(a, 3)
End Function

Private Function A5() As String
    Dim a() As Long
    ReDim a(1, 1)
    On Error Resume Next
    ReDim Preserve a(2, 1)
    A5 = Err.Number
End Function

Private Function A6() As String
    Dim a(2) As Long
    On Error Resume Next
    Erase a
    A6 = Err.Number & ":" & UBound(a)
End Function

Private Function A7() As String
    Dim a() As Long
    ReDim a(2)
    Erase a
    On Error Resume Next
    A7 = UBound(a)
    A7 = A7 & "e" & Err.Number
End Function

Private Function A8() As String
    Dim v As Variant
    v = Array(1, 2)
    Dim w As Variant
    w = v
    w(0) = 9
    A8 = v(0) & w(0)
End Function

Private Function A9() As String
    Dim a(1) As Long, b(1) As Long
    a(0) = 1
    On Error Resume Next
    A9 = TypeName(a)
    A9 = A9 & TypeName(Array(a))
End Function

Private Function A10() As String
    Dim v As Variant, s As String
    For Each v In Array(1, "a", Null)
        s = s & TypeName(v)
    Next
    A10 = s
End Function

Private Function A11() As String
    Dim v As Variant, s As String
    Dim a(1 To 2, 1 To 2) As Long
    a(1, 1) = 1: a(2, 1) = 2: a(1, 2) = 3: a(2, 2) = 4
    For Each v In a
        s = s & v
    Next
    A11 = s
End Function

Private Function A12() As String
    Dim a() As String
    ReDim a(1)
    A12 = "[" & a(0) & "]" & TypeName(a(0))
End Function

Private Function A13() As String
    Dim a() As Variant
    ReDim a(1)
    A13 = TypeName(a(0)) & IsEmpty(a(0))
End Function

Private Function A14() As String
    Dim v As Variant
    ReDim v(1 To 2)
    v(1) = 5
    A14 = TypeName(v) & v(1)
End Function

Private Function A15() As String
    Dim a(1) As Long
    Dim i As Long
    On Error Resume Next
    a(5) = 1
    A15 = Err.Number
End Function

Private Function A16() As String
    Dim v As Variant
    On Error Resume Next
    v = 5
    A16 = v(0)
    A16 = A16 & "e" & Err.Number
End Function

Private Function A17() As String
    Dim a() As Long
    ReDim a(-1 To 1)
    A17 = LBound(a) & UBound(a)
End Function

Private Function A18() As String
    Dim a() As Long
    On Error Resume Next
    ReDim a(-1)
    A18 = Err.Number
    ReDim a(1 To 0)
    A18 = A18 & "," & Err.Number
End Function

Private Function A19() As String
    Dim a As Variant
    a = Array(Array(1, 2), Array(3))
    A19 = a(0)(1) & UBound(a(1))
End Function

Private Function F1() As String
    On Error GoTo Bad
    Err.Raise 1234, "Src", "Desc"
    F1 = "no"
    Exit Function
Bad:
    F1 = Err.Number & Err.Source & Err.Description
End Function

Private Function F2() As String
    On Error Resume Next
    Err.Raise 5
    F2 = Err.Number & Err.Description
End Function

Private Function F3() As String
    On Error GoTo Bad
    Dim x As Long
    x = 1 / 0
    F3 = "after" & x
    Exit Function
Bad:
    x = 7
    Resume Next
End Function

Private Function F4() As String
    Dim n As Long
    On Error GoTo Bad
    n = 1 / n
    F4 = "after"
    Exit Function
Bad:
    n = 1
    Resume
End Function

Private Function F5() As String
    On Error GoTo Bad
    Err.Raise 9
Bad:
    F5 = Err.Number
    On Error GoTo 0
    F5 = F5 & "," & Err.Number
End Function

Private Function F6() As String
    On Error Resume Next
    Err.Raise 5
    Err.Clear
    F6 = Err.Number & "[" & Err.Description & "]"
End Function

Private Function Inner7() As Long
    Err.Raise 77
End Function

Private Function F7() As String
    On Error Resume Next
    Dim x As Long
    x = Inner7()
    F7 = Err.Number & ":" & x
End Function

Private Function F8() As String
    On Error GoTo Bad
    GoTo Done
Bad:
    F8 = "bad"
    Exit Function
Done:
    F8 = "done"
End Function

Private Function F9() As String
    Dim i As Long, s As String
    For i = 1 To 3
        If i = 2 Then GoTo Skip
        s = s & i
Skip:
    Next
    F9 = s & "|" & i
End Function

Private Function F10() As String
    Dim i As Long, s As String
    For i = 3 To 1 Step -1
        s = s & i
    Next
    For i = 1 To 0
        s = s & "x"
    Next
    F10 = s & "|" & i
End Function

Private Function F11() As String
    Dim x As Double, s As String
    For x = 0 To 1 Step 0.25
        s = s & x & ","
    Next
    F11 = s
End Function

Private Function F12() As String
    Dim i As Long, n As Long
    n = 2
    For i = 1 To n
        n = 5
    Next
    F12 = i & "," & n
End Function

Private Function F13() As String
    Dim i As Long, s As String
    Do While i < 2
        i = i + 1
        s = s & i
    Loop
    Do
        i = i - 1
        If i = 0 Then Exit Do
    Loop Until False
    F13 = s & "|" & i
End Function

Private Function F14() As String
    Dim s As String, x As Variant
    For Each x In Array(1, 2.5, "a", 7, 11, Null)
        Select Case x
            Case 1: s = s & "one"
            Case 2 To 3: s = s & "range"
            Case "a": s = s & "text"
            Case Is > 10: s = s & "big"
            Case 5, 6, 7: s = s & "list"
            Case Else: s = s & "else"
        End Select
    Next
    F14 = s
End Function

Private Function F15() As String
    Static n As Long
    n = n + 1
    F15 = n
End Function

Private Function Opt16(Optional a As Variant, Optional b As Long = 5, Optional c As String) As String
    Opt16 = IsMissing(a) & b & "[" & c & "]"
End Function

Private Function F16() As String
    F16 = Opt16() & Opt16(1) & Opt16(, 2) & Opt16(c:="x")
End Function

Private Function Par17(ParamArray p() As Variant) As String
    Par17 = UBound(p) & TypeName(p)
End Function

Private Function F17() As String
    F17 = Par17() & Par17(1, 2)
End Function

Private Sub Ref18(ByRef a As Long, ByVal b As Long)
    a = a + 1: b = b + 1
End Sub

Private Function F18() As String
    Dim x As Long, y As Long
    Ref18 x, y
    Ref18 (x), y
    F18 = x & y
End Function

Private Type Rec19
    A As Long
    B As String
    C(1) As Long
End Type

Private Function F19() As String
    Dim r As Rec19, s As Rec19
    r.A = 1: r.B = "x": r.C(1) = 3
    s = r
    s.A = 2: s.C(1) = 4
    F19 = r.A & r.B & r.C(1) & s.A & s.C(1)
End Function

Private Enum Color20
    Red20
    Green20 = 5
    Blue20
End Enum

Private Function F20() As String
    F20 = Red20 & Green20 & Blue20 & TypeName(Blue20)
End Function

Private Function C1() As String
    Dim c As New Collection
    c.Add 1: c.Add "b", "k"
    C1 = c.Count & c(1) & c("k") & c("K")
End Function

Private Function C2() As String
    Dim c As New Collection
    On Error Resume Next
    C2 = c(1)
    C2 = C2 & Err.Number
End Function

Private Function C3() As String
    Dim c As New Collection
    c.Add 1, "k"
    On Error Resume Next
    c.Add 2, "k"
    C3 = Err.Number & c.Count
    Err.Clear
    C3 = C3 & c("nope")
    C3 = C3 & "," & Err.Number
End Function

Private Function C4() As String
    Dim c As New Collection
    c.Add "a": c.Add "b": c.Add "c", Before:=1: c.Add "d", After:=1
    Dim v As Variant, s As String
    For Each v In c: s = s & v: Next
    C4 = s
End Function

Private Function C5() As String
    Dim c As New Collection
    c.Add "a": c.Add "b"
    c.Remove 1
    On Error Resume Next
    c.Remove 5
    C5 = c.Count & c(1) & Err.Number
End Function

Private Function C6() As String
    Dim c As Collection
    On Error Resume Next
    C6 = c.Count & Err.Number
    C6 = C6 & "," & Err.Number
End Function

Private Function C7() As String
    Dim c As New Collection
    c.Add 1
    On Error Resume Next
    C7 = c(0) & "," & Err.Number
    Err.Clear
    C7 = C7 & c(1.5) & "," & Err.Number
    Err.Clear
    C7 = C7 & c(Null) & "," & Err.Number
End Function

Private Function C8() As String
    Dim c As New Collection
    On Error Resume Next
    c.Add 1, 5
    C8 = Err.Number & c.Count
End Function

Private Function C9() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("a") = 1: d.Add "b", 2
    C9 = d.Count & d("a") & d.Exists("b") & d.Exists("z") & d.Count & d("z") & d.Count
End Function

Private Function C10() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    On Error Resume Next
    d.Add "a", 2
    C10 = Err.Number
    Err.Clear
    d.Remove "z"
    C10 = C10 & "," & Err.Number
End Function

Private Function C11() As String
    Dim d As Object, k As Variant, s As String
    Set d = CreateObject("Scripting.Dictionary")
    d("b") = 1: d("a") = 2: d(3) = 3: d("3") = 4
    For Each k In d.Keys: s = s & TypeName(k) & k: Next
    C11 = s & d.Count & UBound(d.Items)
End Function

Private Function C12() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("a") = 1
    C12 = d.Exists("A")
    d.CompareMode = 1
    C12 = C12 & Err.Number
End Function

Private Function C13() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    d.CompareMode = 1
    d("a") = 1
    C13 = d.Exists("A") & d.Key("a")
    C13 = C13 & Err.Number
    d.Key("a") = "b"
    C13 = C13 & d.Exists("b") & d.Exists("a")
End Function

Private Function C14() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("a") = 1
    d.RemoveAll
    C14 = d.Count & TypeName(d.Keys) & UBound(d.Keys)
End Function

Private Function C15() As String
    Dim c As New Collection, v As Variant
    c.Add Array(1, 2)
    C15 = TypeName(c(1)) & c(1)(1)
End Function

Private Function C16() As String
    Dim c As New Collection
    c.Add ActiveSheet
    C16 = TypeName(c(1)) & (c(1) Is ActiveSheet)
End Function

Private Function C17() As String
    Dim c As New Collection
    Set c = Nothing
    C17 = c.Count
End Function

Private Function C18() As String
    Dim o As Object
    On Error Resume Next
    Set o = CreateObject("No.Such.Thing")
    C18 = Err.Number & TypeName(o)
    Err.Clear
    C18 = C18 & o.Name & Err.Number
End Function
'''


def _read_line(expression: str) -> str:
    return (f'v = Empty\nv = {expression}\ne = Err.Number\nErr.Clear\n'
            f's = s & e & "=" & TypeName(v) & ":" & Shown(v) & "{SEPARATOR}"\n')


SHOWN = '''Private Function Shown(ByVal x As Variant) As String
    On Error Resume Next
    Shown = CStr(x)
    If Err.Number Then Shown = "?" & Err.Number: Err.Clear
End Function
'''


def macro(name: str, leaving_out: frozenset[int] = frozenset()) -> str:
    lines = ''.join(f's = s & "skipped{SEPARATOR}"\n' if at in leaving_out else _read_line(one) for at, one in enumerate(GROUPS[name]))
    return HELPERS + SHOWN + 'Function Probe() As String\nDim s As String, v As Variant, e As Long\nOn Error Resume Next\n' + lines + 'Probe = s\nEnd Function\n'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    def run(source: str) -> object:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=180.0, line_numbers=False)
        if result.dialogs or not result.ok:
            print('  stopped by', [dialog.message[:70] for dialog in result.dialogs], str(result.message)[:60], flush=True)
            return None
        return result.value

    probes: list[dict[str, object]] = json.loads(OUT.read_text(encoding='utf-8'))['probes'] if OUT.exists() else []
    done = {probe['name'] for probe in probes}
    for name in GROUPS:
        if name in done:
            continue
        count = len(GROUPS[name])
        bad: set[int] = set()

        def search(lines: list[int]) -> None:
            if run(macro(name, frozenset(range(count)) - frozenset(lines))) is not None:
                return
            if len(lines) == 1:
                bad.add(lines[0])
                return
            middle = len(lines) // 2
            search(lines[:middle])
            search(lines[middle:])

        answer = run(macro(name))
        if answer is None:
            search(list(range(count)))
            answer = run(macro(name, frozenset(bad)))
        probes.append({'name': name, 'asked': GROUPS[name], 'macro': macro(name, frozenset(bad)), 'answer': answer, 'not_compiling': sorted(bad)})
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=1, ensure_ascii=True) + '\n', encoding='utf-8')
        print(name, count, 'does not compile:', [GROUPS[name][at] for at in sorted(bad)], flush=True)


if __name__ == '__main__':
    main()
