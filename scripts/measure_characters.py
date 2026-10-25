"""Range.Characters in live Excel: the characters of a cell's text, their font, and what edits do to them.

Each case reads or edits part of a cell's text through Range.Characters and records the answer: Text, Count and
Caption, a font property, the error raised, and the font of every character after it ("runs"), each as bold,
italic, strikethrough, superscript (^), subscript (_), then u Underline, s Size, c Color and n Name. The cells
stand in a new workbook, which is saved with its rich text and opened again; the script records the saved
sharedStrings.xml beside the answers.

    python scripts/measure_characters.py

writes tests/fixtures/characters.json, which tests/test_excel_characters.py replays.
"""

from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests" / "fixtures" / "characters.json"

PROBE = r'''Private answer As String

Private Function Shown(value As Variant) As String
    If IsError(value) Then
        Shown = "Error:" & CLng(value)
    ElseIf IsNull(value) Then
        Shown = "Null"
    ElseIf IsEmpty(value) Then
        Shown = "Empty"
    ElseIf IsObject(value) Then
        Shown = "Object:" & TypeName(value)
    Else
        Shown = TypeName(value) & ":" & CStr(value)
    End If
End Function

Private Function Runs(c As Range) As String
    Dim i As Long, out As String, f As Object, n As Long
    On Error Resume Next
    n = Len(CStr(c.Value))
    For i = 1 To n
        Set f = c.Characters(i, 1).Font
        out = out & IIf(f.Bold, "B", "") & IIf(f.Italic, "I", "") & IIf(f.Strikethrough, "S", "") & _
            IIf(f.Superscript, "^", "") & IIf(f.Subscript, "_", "") & "u" & f.Underline & "s" & f.Size & _
            "c" & f.Color & "n" & f.Name & ";"
    Next
    If Err.Number <> 0 Then out = out & "E" & Err.Number
    Runs = out
End Function

Private Function Part(c As Range, Optional start As Variant, Optional length As Variant) As String
    Dim ch As Object
    On Error Resume Next
    If IsMissing(start) And IsMissing(length) Then
        Set ch = c.Characters
    ElseIf IsMissing(length) Then
        Set ch = c.Characters(start)
    ElseIf IsMissing(start) Then
        Set ch = c.Characters(, length)
    Else
        Set ch = c.Characters(start, length)
    End If
    If Err.Number <> 0 Then Part = "E" & Err.Number & ":" & Err.Description: Exit Function
    Part = "text=" & Shown(ch.Text)
    If Err.Number <> 0 Then Part = Part & "(E" & Err.Number & ":" & Err.Description & ")": Err.Clear
    Part = Part & ",count=" & Shown(ch.Count)
    If Err.Number <> 0 Then Part = Part & "(E" & Err.Number & ":" & Err.Description & ")": Err.Clear
    Part = Part & ",caption=" & Shown(ch.Caption)
    If Err.Number <> 0 Then Part = Part & "(E" & Err.Number & ":" & Err.Description & ")": Err.Clear
    Part = Part & ",bold=" & Shown(ch.Font.Bold)
    If Err.Number <> 0 Then Part = Part & "(E" & Err.Number & ":" & Err.Description & ")": Err.Clear
    Part = Part & ",type=" & TypeName(ch)
End Function

Private Function Fresh(ws As Worksheet, address As String, text As String) As Range
    Set Fresh = ws.Range(address)
    Fresh.Clear
    Fresh.Value = text
End Function

Private Function Did(c As Range, what As String) As String
    Dim got As Variant, late As Object
    answer = ""
    On Error Resume Next
    Select Case what
        Case "bold15": c.Characters(1, 5).Font.Bold = True
        Case "color75": c.Characters(7, 5).Font.Color = RGB(255, 0, 0)
        Case "size11": c.Characters(1, 1).Font.Size = 20
        Case "name22": c.Characters(2, 2).Font.Name = "Arial"
        Case "italic44": c.Characters(4, 4).Font.Italic = True
        Case "under33": c.Characters(3, 3).Font.Underline = xlUnderlineStyleDouble
        Case "under_single": c.Characters(1, 2).Font.Underline = xlUnderlineStyleSingle
        Case "under_accounting": c.Characters(1, 2).Font.Underline = xlUnderlineStyleSingleAccounting
        Case "strike91": c.Characters(9, 1).Font.Strikethrough = True
        Case "super101": c.Characters(10, 1).Font.Superscript = True
        Case "sub111": c.Characters(11, 1).Font.Subscript = True
        Case "style52": c.Characters(5, 2).Font.FontStyle = "Bold Italic"
        Case "style_regular": c.Characters(1, 3).Font.FontStyle = "Regular"
        Case "colorindex": c.Characters(1, 2).Font.ColorIndex = 5
        Case "theme": c.Characters(2, 2).Font.ThemeColor = xlThemeColorAccent1
        Case "size_zero": c.Characters(1, 1).Font.Size = 0
        Case "size_big": c.Characters(1, 1).Font.Size = 500
        Case "bold_all": c.Characters.Font.Bold = True
        Case "bold_first_off": c.Characters(1, 1).Font.Bold = False
        Case "insert32": c.Characters(3, 2).Insert "XY"
        Case "insert_start": c.Characters(1, 0).Insert "<"
        Case "insert_middle": c.Characters(4, 0).Insert "+"
        Case "insert_end": c.Characters(Len(c.Value) + 1, 0).Insert "!"
        Case "insert_answer": Set late = c.Characters(2, 1): got = late.Insert("Z"): answer = Shown(got)
        Case "delete12": c.Characters(1, 2).Delete
        Case "delete_answer": Set late = c.Characters(2, 1): got = late.Delete: answer = Shown(got)
        Case "delete_all": c.Characters.Delete
        Case "delete_across": c.Characters(2, 3).Delete
        Case "text_set": c.Characters(2, 3).Text = "abc"
        Case "text_all": c.Characters.Text = "All new"
        Case "text_empty": c.Characters.Text = ""
        Case "caption_set": c.Characters(2, 3).Caption = "Q"
        Case "whole_bold": c.Font.Bold = True
        Case "whole_size": c.Font.Size = 14
        Case "whole_name": c.Font.Name = "Arial"
        Case "same_value": c.Value = c.Value
        Case "new_value": c.Value = "Fresh text"
        Case "formula_value": c.Formula = "=""ab"""
        Case "clear_formats": c.ClearFormats
        Case "clear_contents": c.ClearContents
        Case "copy_all": c.Copy c.Offset(0, 1)
        Case "paste_values": c.Copy: c.Offset(0, 1).PasteSpecial xlPasteValues: Application.CutCopyMode = False
        Case "paste_formats": c.Copy: c.Offset(0, 1).PasteSpecial xlPasteFormats: Application.CutCopyMode = False
        Case "replace": c.Replace "ll", "LL", xlPart
        Case "style_normal": c.Style = "Normal"
        Case "value2": c.Value2 = "Two"
        Case "delete_zero": c.Characters(2, 0).Delete
        Case "text_zero": c.Characters(2, 0).Text = "Q"
        Case "delete_negative": c.Characters(3, -1).Delete
        Case "insert_negative": c.Characters(3, -1).Insert "Q"
        Case "bold_past_end": c.Characters(Len(c.Value) + 3, 2).Font.Bold = True
        Case "bold_empty_span": c.Characters(2, 0).Font.Bold = True
    End Select
    If Err.Number <> 0 Then Did = what & "=E" & Err.Number & ":" & Err.Description Else Did = what & " ok"
    If answer <> "" Then Did = Did & " answered " & answer
End Function

Private Function Edit(ws As Worksheet, address As String, text As String, setup As String, what As String) As String
    Dim c As Range, step As Variant
    Set c = Fresh(ws, address, text)
    If setup <> "" Then
        For Each step In Split(setup, " ")
            Call Did(c, CStr(step))
        Next
    End If
    Edit = Did(c, what) & " value=" & Shown(c.Value) & " runs " & Runs(c)
End Function

Private Function LongText(ws As Worksheet, n As Long, what As String) As String
    Dim c As Range
    Set c = Fresh(ws, "E1", String(n, "x"))
    On Error Resume Next
    Select Case what
        Case "insert": c.Characters(3, 2).Insert "YZ"
        Case "insert_grow": c.Characters(3, 0).Insert "YZ"
        Case "delete": c.Characters(3, 2).Delete
        Case "text": c.Characters(3, 2).Text = "QQ"
        Case "bold": c.Characters(n, 1).Font.Bold = True
    End Select
    If Err.Number <> 0 Then LongText = what & n & "=E" & Err.Number Else LongText = what & n
    Err.Clear
    LongText = LongText & " len=" & Len(c.Value) & " at3=" & Mid(c.Value, 3, 2) & " lastbold=" & c.Characters(n, 1).Font.Bold
    c.Clear
End Function

Public Function Probe(folder As String) As String
    Dim out As String, ws As Worksheet, wb As Workbook, c As Range, what As Variant
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    Set c = Fresh(ws, "A1", "Hello World")
    out = "missing " & Part(c)
    out = out & "~(1,5) " & Part(c, 1, 5)
    out = out & "~(7) " & Part(c, 7)
    out = out & "~(,5) " & Part(c, , 5)
    out = out & "~(0,3) " & Part(c, 0, 3)
    out = out & "~(-1,3) " & Part(c, -1, 3)
    out = out & "~(20,3) " & Part(c, 20, 3)
    out = out & "~(12,3) " & Part(c, 12, 3)
    out = out & "~(12) " & Part(c, 12)
    out = out & "~(5,0) " & Part(c, 5, 0)
    out = out & "~(5,-1) " & Part(c, 5, -1)
    out = out & "~(1,100) " & Part(c, 1, 100)
    out = out & "~(2.5,2.5) " & Part(c, 2.5, 2.5)
    out = out & "~(""2"",""3"") " & Part(c, "2", "3")
    out = out & "~(""x"",1) " & Part(c, "x", 1)
    out = out & "~(Null,1) " & Part(c, Null, 1)
    out = out & "~plain runs " & Runs(c)
    For Each what In Array("bold15", "color75", "size11", "name22", "italic44", "under33", "strike91", _
                           "super101", "sub111", "style52", "colorindex")
        out = out & "~" & Did(c, CStr(what))
    Next
    out = out & "~runs " & Runs(c)
    out = out & "~cell bold=" & Shown(c.Font.Bold) & ",size=" & Shown(c.Font.Size) & ",name=" & Shown(c.Font.Name) & _
        ",color=" & Shown(c.Font.Color) & ",italic=" & Shown(c.Font.Italic) & ",underline=" & Shown(c.Font.Underline) & _
        ",style=" & Shown(c.Font.FontStyle) & ",value=" & Shown(c.Value) & ",text=" & Shown(c.Text)
    out = out & "~parts (3,5)bold=" & Shown(c.Characters(3, 5).Font.Bold) & ",(1,5)bold=" & _
        Shown(c.Characters(1, 5).Font.Bold) & ",(7,5)color=" & Shown(c.Characters(7, 5).Font.Color) & _
        ",(7,5)colorindex=" & Shown(c.Characters(7, 5).Font.ColorIndex) & ",(1,2)colorindex=" & _
        Shown(c.Characters(1, 2).Font.ColorIndex) & ",(5,2)style=" & Shown(c.Characters(5, 2).Font.FontStyle) & _
        ",(1,3)name=" & Shown(c.Characters(1, 3).Font.Name) & ",(1,2)size=" & Shown(c.Characters(1, 2).Font.Size) & _
        ",(3,1)underline=" & Shown(c.Characters(3, 1).Font.Underline) & ",(1,11)italic=" & _
        Shown(c.Characters(1, 11).Font.Italic) & ",(20,2)bold=" & Shown(c.Characters(20, 2).Font.Bold)
    Application.DisplayAlerts = False
    wb.SaveAs folder & "\rich.xlsx", 51
    For Each what In Array("insert32", "insert_start", "insert_middle", "insert_end", "insert_answer", "delete12", _
                           "delete_answer", "delete_all", "delete_across", "text_set", "text_all", "text_empty", _
                           "caption_set", "delete_zero", "text_zero", "delete_negative", "insert_negative")
        out = out & "~" & Edit(ws, "B1", "HeLLo", "italic44", CStr(what))
    Next
    For Each what In Array("bold_past_end", "bold_empty_span", "bold_all", "bold_first_off", "style_regular", _
                           "under_single", "under_accounting", "theme", _
                           "size_zero", "size_big", "whole_bold", "whole_size", "whole_name", "same_value", _
                           "new_value", "formula_value", "clear_formats", "clear_contents", "style_normal", "value2", _
                           "replace")
        out = out & "~" & Edit(ws, "B2", "Hello", "italic44 size11", CStr(what))
    Next
    For Each what In Array("copy_all", "paste_values", "paste_formats")
        out = out & "~" & Edit(ws, "B3", "Hello", "italic44 size11", CStr(what)) & " target value=" & _
            Shown(ws.Range("C3").Value) & " target runs " & Runs(ws.Range("C3"))
        ws.Range("C3").Clear
    Next
    ws.Range("B4").Clear
    ws.Range("B4").Font.Bold = True
    ws.Range("B4").Value = "Bold cell"
    ws.Range("B4").Characters(2, 2).Font.Bold = False
    out = out & "~bold cell part off runs " & Runs(ws.Range("B4")) & " cell=" & Shown(ws.Range("B4").Font.Bold)
    Set c = Fresh(ws, "B14", "Hello")
    c.Characters(2, 4).Font.Bold = True
    out = out & "~length zero reads (2,0)bold=" & Shown(c.Characters(2, 0).Font.Bold) & ",(1,0)bold=" & _
        Shown(c.Characters(1, 0).Font.Bold) & ",(1,-1)bold=" & Shown(c.Characters(1, -1).Font.Bold) & _
        ",(2,-1)bold=" & Shown(c.Characters(2, -1).Font.Bold) & ",(6,0)bold=" & Shown(c.Characters(6, 0).Font.Bold) & _
        ",(9,1)bold=" & Shown(c.Characters(9, 1).Font.Bold) & ",(,2)bold=" & Shown(c.Characters(, 2).Font.Bold)
    c.Characters(1, -1).Font.Italic = True
    out = out & "~negative length write runs " & Runs(c)
    ws.Range("B5").Value = "a" & vbLf & "b"
    ws.Range("B5").Characters(3, 1).Font.Bold = True
    out = out & "~line break " & Part(ws.Range("B5"), 2, 2) & " runs " & Runs(ws.Range("B5"))
    ws.Range("B6").Value = String(300, "x")
    out = out & "~long " & Part(ws.Range("B6"), 290, 20)
    ws.Range("B6").Characters(299, 2).Font.Bold = True
    out = out & " bold tail cell=" & Shown(ws.Range("B6").Font.Bold) & ",(300,1)=" & _
        Shown(ws.Range("B6").Characters(300, 1).Font.Bold) & ",(1,1)=" & Shown(ws.Range("B6").Characters(1, 1).Font.Bold)
    out = out & "~" & Did(ws.Range("B6"), "insert32") & " len=" & Len(ws.Range("B6").Value)
    For Each what In Array(253, 254, 255, 256, 300)
        out = out & "~" & LongText(ws, CLng(what), "insert") & "," & LongText(ws, CLng(what), "insert_grow") & "," & _
            LongText(ws, CLng(what), "delete") & "," & LongText(ws, CLng(what), "text") & "," & LongText(ws, CLng(what), "bold")
    Next
    ws.Range("B7").Value = 12345
    out = out & "~number " & Part(ws.Range("B7"), 1, 2) & " all " & Part(ws.Range("B7"))
    ws.Range("B7").Characters(1, 2).Font.Bold = True
    out = out & " after bold cell=" & Shown(ws.Range("B7").Font.Bold)
    out = out & "~" & Did(ws.Range("B7"), "insert32") & " value=" & Shown(ws.Range("B7").Value)
    out = out & "~" & Did(ws.Range("B7"), "delete12") & " value=" & Shown(ws.Range("B7").Value)
    out = out & "~" & Did(ws.Range("B7"), "text_set") & " value=" & Shown(ws.Range("B7").Value)
    ws.Range("B8").Formula = "=1+2"
    out = out & "~formula " & Part(ws.Range("B8"), 1, 1)
    ws.Range("B8").Characters(1, 1).Font.Italic = True
    out = out & " after italic cell=" & Shown(ws.Range("B8").Font.Italic)
    ws.Range("B9").Formula = "=""ab""&""cd"""
    out = out & "~text formula " & Part(ws.Range("B9"), 1, 2)
    ws.Range("B9").Characters(1, 2).Font.Bold = True
    out = out & " after bold cell=" & Shown(ws.Range("B9").Font.Bold)
    out = out & "~empty " & Part(ws.Range("B10"), 1, 1)
    ws.Range("B10").Characters(1, 1).Font.Bold = True
    out = out & " after bold cell=" & Shown(ws.Range("B10").Font.Bold)
    out = out & "~" & Did(ws.Range("B10"), "insert32") & " value=" & Shown(ws.Range("B10").Value)
    ws.Range("B11").Value = True
    out = out & "~boolean " & Part(ws.Range("B11"), 1, 2)
    ws.Range("B12").Value = DateSerial(2020, 1, 2)
    out = out & "~date " & Part(ws.Range("B12"), 1, 3)
    ws.Range("B13").Value = "'0123"
    out = out & "~prefixed " & Part(ws.Range("B13"), 1, 2)
    ws.Range("C1").Value = "ab": ws.Range("C2").Value = "cd"
    out = out & "~two cells " & Part(ws.Range("C1:C2"), 1, 1)
    ws.Range("C1:C2").Characters(1, 1).Font.Bold = True
    out = out & " after C1=" & Runs(ws.Range("C1")) & " C2=" & Runs(ws.Range("C2"))
    ws.Range("D1").Value = "zeta": ws.Range("D2").Value = "alpha"
    ws.Range("D1").Characters(1, 2).Font.Bold = True
    ws.Range("D1:D2").Sort Key1:=ws.Range("D1"), Order1:=xlAscending, Header:=xlNo
    out = out & "~sorted D1=" & Shown(ws.Range("D1").Value) & " runs " & Runs(ws.Range("D1")) & " D2=" & _
        Shown(ws.Range("D2").Value) & " runs " & Runs(ws.Range("D2"))
    wb.Save
    wb.Close False
    Set wb = Workbooks.Open(folder & "\rich.xlsx")
    Set ws = wb.Worksheets(1)
    out = out & "~reopened A1 " & Runs(ws.Range("A1")) & " B4 " & Runs(ws.Range("B4")) & " D2 " & Runs(ws.Range("D2"))
    wb.Close False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1).Range("A1")
    c.Value = "abcdefgh"
    c.Characters(6, 3).Font.Bold = True
    wb.SaveAs folder & "\dirty.xlsx", xlOpenXMLWorkbook
    out = out & "~dirty saved=" & wb.Saved
    c.Characters(6, 3).Font.Bold = True
    out = out & " same_font=" & wb.Saved
    wb.Save
    c.Characters(2, 1).Font.Italic = True
    out = out & " run_alone=" & wb.Saved
    wb.Save
    c.Characters(2, 1).Font.Size = 11
    out = out & " same_size=" & wb.Saved
    wb.Save
    c.Characters(20, 1).Font.Bold = True
    out = out & " past_end_font=" & wb.Saved
    wb.Save
    c.Characters(2, 1).Text = "b"
    out = out & " same_text=" & wb.Saved
    wb.Save
    out = out & " read=" & c.Characters(2, 3).Text & c.Characters(2, 3).Font.Bold & " " & wb.Saved
    c.Value = String(255, "x")
    wb.Save
    c.Characters(3, 0).Insert "YZ"
    out = out & " too_long=" & wb.Saved
    out = out & "~types " & TypeName(c.Characters) & "|" & TypeName(c.Characters(1, 2)) & "|" & _
        TypeName(c.Characters(1, 2).Font) & "|" & TypeName(c.Characters(1, 2).Font.Parent) & "|" & _
        TypeName(c.Characters(1, 2).Parent) & "|" & c.Characters(1, 2).Creator & "|" & _
        (c.Characters(1, 2).Application Is Application) & "|" & (c.Characters(1, 2).Parent Is c)
    wb.Close False
    Probe = out
End Function
'''


def main() -> None:
    folder = Path(tempfile.mkdtemp(prefix="characters_"))
    try:
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(PROBE, "Probe", [str(folder)], timeout=600.0)
            assert result.ok, f"{result.outcome}: {result.message} {result.error}"
        with zipfile.ZipFile(folder / "rich.xlsx") as package:
            strings = package.read("xl/sharedStrings.xml").decode("utf-8")
            sheet = package.read("xl/worksheets/sheet1.xml").decode("utf-8")
            fonts = re.search(r"<fonts\b.*?</fonts>", package.read("xl/styles.xml").decode("utf-8"), re.DOTALL)
    finally:
        shutil.rmtree(folder, ignore_errors=True)
    record = {"probe": PROBE, "answers": str(result.value).split("~"), "shared_strings": strings, "sheet": sheet,
              "fonts": fonts.group(0) if fonts else ""}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for answer in record["answers"]:
        print(answer[:300])


if __name__ == "__main__":
    main()
