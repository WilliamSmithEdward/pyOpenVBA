"""Workbook.LinkSources, LinkInfo, BreakLink, UpdateLinks and SaveLinkValues on a workbook whose links' sources are
closed.

Alpha.xlsx and Beta.xlsx hold the cells and names Dest.xlsx reads; all
three are saved with all three open, Dest's first formula reading Beta and
the rest mostly Alpha. Each case then reopens Dest alone, its links not
updated, as the model opens it: what LinkSources and LinkInfo answer, what
each cell and name is after BreakLink breaks one link, the other or both,
the errors BreakLink raises, the two link settings and what each saves,
how arguments of other types are taken, and last how LinkInfo's status
follows a source changed, moved away or missing a sheet. Plain, Named,
Local and Mixed.xlsx are saved with their sources open and again with them
closed, Mixed after a break, for what a save writes of each link. A
SheetChange and SheetCalculate sink, with events on for the case, logs
what a break raises.

Excel refuses BreakLink with 1004 on a workbook a macro has just opened
until its links are listed (LinkSources, LinkInfo; a Calculate, a sheet's
Activate or idle time do as well), so each case lists them first, and the
refusal is recorded as "unlisted". A break whose formulas include a
protected sheet's raises an alert that DisplayAlerts does not hold back,
and hangs the macro; no case does that.

The folder the files lived in is written as <folder>, so no machine's path
is kept. The parts Excel saved are recorded the same way, for the model to
open and to compare its own saves with.

    python scripts/measure_workbook_links.py

writes tests/fixtures/workbook_links.json, which
tests/test_excel_workbook_links.py replays.
"""
from __future__ import annotations

import base64
import json
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/workbook_links.json'

#: Dest's formulas in the order they are written, Beta's first so that it is the first link made.
DEST = [
    ('A1', '=[Beta.xlsx]Sheet1!A1'), ('A2', '=[Alpha.xlsx]Sheet1!A1'), ('A3', '=[Alpha.xlsx]Sheet1!A2*B1'),
    ('A4', '=SUM([Alpha.xlsx]Sheet1!A1:A5)'), ('A5', '=Alpha.xlsx!Rate*100'), ('A6', '=[Alpha.xlsx]Sheet1!B3'),
    ('A7', '=[Alpha.xlsx]Sheet1!A1+[Beta.xlsx]Sheet1!A2'), ('A8', '=[Alpha.xlsx]Sheet1!A6'), ('A9', '=B1*2'),
    ('A10', '=[Alpha.xlsx]Sheet1!A1/0'), ('A11', '=IF([Alpha.xlsx]Sheet1!A1>0,"yes","no")'),
    ('A12', '=[Alpha.xlsx]Sheet1!A1>0'), ('A13', '=[Alpha.xlsx]Sheet1!C2'), ('A14', '=SUM(Ext)'),
    ('A15', '=BetaCell*2'), ('A16', '=Beta.xlsx!Bonus'), ('A17', '=A2+1'), ('A18', '=Outer'),
]
#: Every cell a case reads back, by sheet.
READ = [('Dest', f'A{row}') for row in range(1, 19)] + [('Dest', 'B1')] \
    + [('Dest', f'C{row}') for row in range(1, 4)] + [('Dest', f'E{row}') for row in range(1, 4)] \
    + [('Second', 'A1'), ('Second', 'A2')]

SINK = '''Public WithEvents App As Application
Public Log As String

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "|change " & Sh.Name & "!" & Target.Address(False, False)
End Sub

Private Sub App_SheetCalculate(ByVal Sh As Object)
    Log = Log & "|calculate " & Sh.Name
End Sub
'''


def _literal(text: str) -> str:
    return '"' + text.replace('"', '""') + '"'


def _module() -> str:
    formulas = '\n'.join(f'    ws.Range("{cell}").Formula = {_literal(formula)}' for cell, formula in DEST[:1]) + '\n' \
        + '    d.Names.Add Name:="Ext", RefersTo:="=[Alpha.xlsx]Sheet1!$A$1:$A$3"\n' \
        + '    d.Names.Add Name:="BetaCell", RefersTo:="=[Beta.xlsx]Sheet1!$A$1"\n' \
        + '    d.Names.Add Name:="Fixed", RefersTo:="=10"\n' \
        + '    d.Names.Add Name:="Outer", RefersTo:="=SUM(Ext)*2"\n' \
        + '\n'.join(f'    ws.Range("{cell}").Formula = {_literal(formula)}' for cell, formula in DEST[1:])
    # A statement a cell, as VBA's lines stop at 1023 characters.
    reads = '\n'.join(f'    s = s & "|" & Described(wb.Worksheets("{sheet}").Range("{cell}"))' for sheet, cell in READ)
    values = '\n'.join(f'    s = s & "|" & Shown(wb.Worksheets("{sheet}").Range("{cell}").Value)' for sheet, cell in READ)
    return '''Option Explicit

Private sink As LinkSink

Private Function Shown(ByVal v As Variant) As String
    If IsArray(v) Then
        Shown = "Array"
    ElseIf IsEmpty(v) Then
        Shown = "Empty"
    Else
        Shown = CStr(v)
    End If
End Function

Private Function Described(ByVal r As Range) As String
    Described = r.Address(False, False) & "=" & r.Formula & "~" & r.Formula2 & "~" & r.HasFormula & "~" _
        & r.HasArray & "~" & r.HasSpill & "~" & TypeName(r.Value) & "~" & Shown(r.Value) & "~" & r.Text
End Function

Private Function CellList(ByVal wb As Workbook) As String
    Dim s As String
''' + reads + '''
    CellList = Mid(s, 2)
End Function

Private Function ValueList(ByVal wb As Workbook) As String
    Dim s As String
''' + values + '''
    ValueList = Mid(s, 2)
End Function

Private Function NameList(ByVal wb As Workbook) As String
    Dim n As Name, s As String
    For Each n In wb.Names
        s = s & "|" & n.Name & "=" & n.RefersTo
    Next
    NameList = Mid(s, 2)
End Function

Private Function LinkList(ByVal wb As Workbook, Optional ByVal t As Variant) As String
    Dim v As Variant, i As Long, s As String
    If IsMissing(t) Then v = wb.LinkSources() Else v = wb.LinkSources(t)
    If IsEmpty(v) Then
        LinkList = "Empty"
        Exit Function
    End If
    s = TypeName(v) & "(" & LBound(v) & " To " & UBound(v) & ")"
    For i = LBound(v) To UBound(v)
        s = s & "|" & v(i)
    Next
    LinkList = s
End Function

Private Function LinksTried(ByVal wb As Workbook, ByVal t As Variant) As String
    On Error GoTo Failed
    LinksTried = LinkList(wb, t)
    Exit Function
Failed:
    LinksTried = "Err " & Err.Number & " " & Err.Description
End Function

Private Function InfoTried(ByVal wb As Workbook, ByVal nm As Variant, ByVal what As Variant, _
                           Optional ByVal t As Variant) As String
    Dim v As Variant
    On Error GoTo Failed
    If IsMissing(t) Then v = wb.LinkInfo(nm, what) Else v = wb.LinkInfo(nm, what, t)
    InfoTried = TypeName(v) & " " & Shown(v)
    Exit Function
Failed:
    InfoTried = "Err " & Err.Number & " " & Err.Description
End Function

Private Function Opened(ByVal folder As String, ByVal file As String) As Workbook
    Set Opened = Workbooks.Open(Filename:=folder & "\\" & file, UpdateLinks:=0)
End Function

Public Function BuildFiles(ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, d As Workbook, ws As Worksheet, w2 As Worksheet
    Application.DisplayAlerts = False
    Set a = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Name = "Sheet1"
    a.Worksheets.Add(After:=a.Worksheets(1)).Name = "Other"
    With a.Worksheets("Sheet1")
        .Range("A1").Value = 1: .Range("B1").Value = "a": .Range("C1").Value = 10
        .Range("A2").Value = 2: .Range("B2").Value = "b": .Range("C2").Value = 20
        .Range("A3").Value = 3: .Range("B3").Value = "c": .Range("C3").Value = 30
        .Range("A4").Value = 4: .Range("B4").Value = "d": .Range("C4").Value = 40
        .Range("A5").Value = 5: .Range("B5").Value = "e": .Range("C5").Value = 50
    End With
    a.Worksheets("Other").Range("A1").Value = 7
    a.Names.Add Name:="Rate", RefersTo:="=Sheet1!$C$1"
    a.SaveAs Filename:=folder & "\\Alpha.xlsx", FileFormat:=xlOpenXMLWorkbook
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.Worksheets(1).Name = "Sheet1"
    b.Worksheets(1).Range("A1").Value = 100
    b.Worksheets(1).Range("A2").Value = 200
    b.Names.Add Name:="Bonus", RefersTo:="=Sheet1!$A$2"
    b.SaveAs Filename:=folder & "\\Beta.xlsx", FileFormat:=xlOpenXMLWorkbook
    Set d = Workbooks.Add(xlWBATWorksheet)
    Set ws = d.Worksheets(1)
    ws.Name = "Dest"
    Set w2 = d.Worksheets.Add(After:=ws)
    w2.Name = "Second"
    ws.Range("B1").Value = 10
''' + formulas + '''
    ws.Range("A13").NumberFormat = "0.00"
    ws.Range("C1:C3").FormulaArray = "=[Alpha.xlsx]Sheet1!A1:A3*2"
    ws.Range("E1").Formula2 = "=[Alpha.xlsx]Sheet1!A1:A3"
    w2.Range("A1").Formula = "=[Alpha.xlsx]Other!A1"
    w2.Range("A2").Formula = "=Dest!A2*3"
    ws.Activate
    d.SaveAs Filename:=folder & "\\Dest.xlsx", FileFormat:=xlOpenXMLWorkbook
    a.Close False
    b.Close False
    d.Close False
    Application.DisplayAlerts = True
    BuildFiles = "ok"
End Function

Public Function SurveyLinks(ByVal folder As String) As String
    Dim d As Workbook, n As Workbook, s As String, v As Variant, i As Long
    Application.DisplayAlerts = False
    Set d = Opened(folder, "Dest.xlsx")
    s = "default:" & LinkList(d) & "#1:" & LinkList(d, 1) & "#2:" & LinkList(d, 2) & "#5:" & LinkList(d, 5)
    s = s & "#6:" & LinkList(d, 6) & "#3:" & LinksTried(d, 3) & "#0:" & LinksTried(d, 0) & "#4:" & LinksTried(d, 4)
    s = s & "#text:" & LinksTried(d, "x") & "#text 2:" & LinksTried(d, "2") & "#1.4:" & LinksTried(d, 1.4)
    v = d.LinkSources(1)
    For i = LBound(v) To UBound(v)
        s = s & "#info " & i & ":" & InfoTried(d, v(i), 1) & "," & InfoTried(d, v(i), 2) & "," & InfoTried(d, v(i), 3)
        s = s & "," & InfoTried(d, v(i), 3, 2) & "," & InfoTried(d, v(i), 4) & "," & InfoTried(d, LCase(v(i)), 3)
    Next
    s = s & "#info nowhere:" & InfoTried(d, "nowhere.xlsx", 3) & "," & InfoTried(d, "Alpha.xlsx", 3)
    s = s & "#UpdateLinks:" & d.UpdateLinks & "#SaveLinkValues:" & d.SaveLinkValues & "#saved:" & d.Saved
    s = s & "#cells:" & CellList(d) & "#names:" & NameList(d)
    d.Close False
    Set n = Workbooks.Add
    s = s & "#new:" & LinkList(n) & "," & LinkList(n, 2) & "," & InfoTried(n, "nowhere.xlsx", 3)
    s = s & "," & n.UpdateLinks & "," & n.SaveLinkValues
    n.Close False
    Application.DisplayAlerts = True
    SurveyLinks = s
End Function

Public Function BreakEvents(ByVal folder As String) As String
    Dim d As Workbook, s As String, v As Variant, wasOn As Boolean
    Application.DisplayAlerts = False
    ' The harness runs with events off; the sink needs them on, and the setting goes back as it was.
    wasOn = Application.EnableEvents
    Application.EnableEvents = True
    Set d = Opened(folder, "Dest.xlsx")
    Set sink = New LinkSink
    Set sink.App = Application
    v = d.LinkSources()
    d.BreakLink Name:=folder & "\\Alpha.xlsx", Type:=xlLinkTypeExcelLinks
    s = "break:" & sink.Log
    sink.Log = ""
    Application.CalculateFull
    s = s & "#full:" & sink.Log
    sink.Log = ""
    d.Worksheets("Dest").Range("B1").Value = 11
    s = s & "#edit:" & sink.Log
    Set sink.App = Nothing
    Set sink = Nothing
    d.Close False
    Application.EnableEvents = wasOn
    Application.DisplayAlerts = True
    BreakEvents = s
End Function

Public Function BreakCase(ByVal folder As String, ByVal which As String, ByVal saveName As String) As String
    Dim d As Workbook, s As String, v As Variant, i As Long
    Application.DisplayAlerts = False
    Set d = Opened(folder, "Dest.xlsx")
    ' Excel refuses BreakLink on a workbook a macro has just opened until its links have been listed.
    v = d.LinkSources()
    Select Case which
    Case "Alpha"
        d.BreakLink Name:=folder & "\\Alpha.xlsx", Type:=xlLinkTypeExcelLinks
    Case "Beta"
        d.BreakLink Name:=folder & "\\Beta.xlsx", Type:=xlLinkTypeExcelLinks
    Case "All"
        v = d.LinkSources(xlExcelLinks)
        For i = LBound(v) To UBound(v)
            d.BreakLink Name:=v(i), Type:=xlLinkTypeExcelLinks
        Next
    End Select
    s = "saved:" & d.Saved & "#links:" & LinkList(d) & "#cells:" & CellList(d) & "#names:" & NameList(d)
    d.SaveAs Filename:=folder & "\\" & saveName, FileFormat:=xlOpenXMLWorkbook
    Application.CalculateFull
    s = s & "#full:" & ValueList(d)
    d.Close False
    Set d = Opened(folder, saveName)
    s = s & "#reopened:" & LinkList(d) & "#" & CellList(d) & "#" & NameList(d)
    d.Close False
    Application.DisplayAlerts = True
    BreakCase = s
End Function

Private Function BreakTried(ByVal folder As String, ByVal nm As Variant, ByVal t As Variant, _
                            Optional ByVal lateBound As Boolean, Optional ByVal listed As Boolean = True) As String
    Dim d As Workbook, o As Object, v As Variant
    Set d = Opened(folder, "Dest.xlsx")
    If listed Then v = d.LinkSources()
    On Error GoTo Failed
    If lateBound Then
        Set o = d
        o.BreakLink nm
    Else
        d.BreakLink Name:=nm, Type:=t
    End If
    BreakTried = "ok " & LinkList(d) & " " & d.Worksheets("Dest").Range("A2").HasFormula
    d.Close False
    Exit Function
Failed:
    BreakTried = "Err " & Err.Number & " " & Err.Description
    d.Close False
End Function

Public Function ErrorCases(ByVal folder As String) As String
    Dim d As Workbook, s As String, alpha As String
    Application.DisplayAlerts = False
    alpha = folder & "\\Alpha.xlsx"
    s = "unlisted:" & BreakTried(folder, alpha, 1, False, False)
    s = s & "#nowhere:" & BreakTried(folder, "C:\\nowhere\\Alpha.xlsx", 1)
    s = s & "#file only:" & BreakTried(folder, "Alpha.xlsx", 1)
    s = s & "#empty:" & BreakTried(folder, "", 1)
    s = s & "#null:" & BreakTried(folder, Null, 1)
    s = s & "#type 2:" & BreakTried(folder, alpha, 2)
    s = s & "#type 3:" & BreakTried(folder, alpha, 3)
    s = s & "#type 0:" & BreakTried(folder, alpha, 0)
    s = s & "#type -1:" & BreakTried(folder, alpha, -1)
    s = s & "#type x:" & BreakTried(folder, alpha, "x")
    s = s & "#type text 1:" & BreakTried(folder, alpha, "1")
    s = s & "#type 1.4:" & BreakTried(folder, alpha, 1.4)
    s = s & "#type missing:" & BreakTried(folder, alpha, 1, True)
    s = s & "#lower:" & BreakTried(folder, LCase(alpha), 1)
    s = s & "#upper:" & BreakTried(folder, UCase(alpha), 1)
    s = s & "#slashes:" & BreakTried(folder, Replace(alpha, "\\", "/"), 1)
    Set d = Opened(folder, "Dest.xlsx")
    s = s & "#twice:" & LinkList(d)
    d.BreakLink Name:=alpha, Type:=xlLinkTypeExcelLinks
    On Error Resume Next
    d.BreakLink Name:=alpha, Type:=xlLinkTypeExcelLinks
    s = s & " " & Err.Number & " " & Err.Description & " " & LinkList(d)
    On Error GoTo 0
    d.Close False
    Application.DisplayAlerts = True
    ErrorCases = s
End Function

Private Function UpdateTried(ByVal wb As Workbook, ByVal v As Variant) As String
    On Error GoTo Failed
    wb.UpdateLinks = v
    UpdateTried = "ok " & wb.UpdateLinks
    Exit Function
Failed:
    UpdateTried = "Err " & Err.Number & " " & Err.Description & " now " & wb.UpdateLinks
End Function

Private Function ValuesTried(ByVal wb As Workbook, ByVal v As Variant) As String
    On Error GoTo Failed
    wb.SaveLinkValues = v
    ValuesTried = "ok " & wb.SaveLinkValues
    Exit Function
Failed:
    ValuesTried = "Err " & Err.Number & " " & Err.Description & " now " & wb.SaveLinkValues
End Function

Public Function LinkSettings(ByVal folder As String) As String
    Dim d As Workbook, s As String
    Application.DisplayAlerts = False
    Set d = Opened(folder, "Dest.xlsx")
    s = "set 2:" & UpdateTried(d, 2) & "#set 3:" & UpdateTried(d, 3) & "#set 1:" & UpdateTried(d, 1)
    s = s & "#set 0:" & UpdateTried(d, 0) & "#set 4:" & UpdateTried(d, 4) & "#set x:" & UpdateTried(d, "x")
    s = s & "#set 2.6:" & UpdateTried(d, 2.6) & "#set text 2:" & UpdateTried(d, "2") & "#set null:" & UpdateTried(d, Null)
    s = s & "#values x:" & ValuesTried(d, "x") & "#values 0:" & ValuesTried(d, 0) & "#values -1:" & ValuesTried(d, -1)
    s = s & "#values text:" & ValuesTried(d, "False") & "#values null:" & ValuesTried(d, Null)
    s = s & "#saved:" & d.Saved
    d.Close False
    Set d = Opened(folder, "Dest.xlsx")
    d.UpdateLinks = xlUpdateLinksNever
    d.SaveLinkValues = False
    s = s & "#set:" & d.UpdateLinks & " " & d.SaveLinkValues & " " & d.Saved
    d.SaveAs Filename:=folder & "\\Never.xlsx", FileFormat:=xlOpenXMLWorkbook
    ' Saved without the values, then with them again: the workbook still has them.
    d.SaveLinkValues = True
    d.SaveAs Filename:=folder & "\\ValuesAgain.xlsx", FileFormat:=xlOpenXMLWorkbook
    d.Close False
    Set d = Opened(folder, "Never.xlsx")
    s = s & "#reopened:" & d.UpdateLinks & " " & d.SaveLinkValues & " " & LinkList(d) & "#" & ValueList(d)
    Application.CalculateFull
    s = s & "#full:" & ValueList(d)
    d.Close False
    Set d = Opened(folder, "Dest.xlsx")
    d.UpdateLinks = xlUpdateLinksAlways
    d.SaveAs Filename:=folder & "\\Always.xlsx", FileFormat:=xlOpenXMLWorkbook
    d.Close False
    Set d = Opened(folder, "Always.xlsx")
    s = s & "#always:" & d.UpdateLinks & " " & d.SaveLinkValues
    d.Close False
    Application.DisplayAlerts = True
    LinkSettings = s
End Function

Private Function SourcesTried(ByVal wb As Workbook, ByVal t As Variant) As String
    Dim v As Variant
    On Error GoTo Failed
    v = wb.LinkSources(t)
    If IsEmpty(v) Then SourcesTried = "Empty" Else SourcesTried = TypeName(v) & " " & (UBound(v) - LBound(v) + 1)
    Exit Function
Failed:
    SourcesTried = "Err " & Err.Number & " " & Err.Description
End Function

Public Function MoreCases(ByVal folder As String) As String
    Dim d As Workbook, n As Workbook, s As String, alpha As String, o As Object, v As Variant
    Application.DisplayAlerts = False
    alpha = folder & "\\Alpha.xlsx"
    Set d = Opened(folder, "Dest.xlsx")
    s = "sources:" & SourcesTried(d, "1") & "," & SourcesTried(d, True) & "," & SourcesTried(d, Null) & "," _
        & SourcesTried(d, Empty) & "," & SourcesTried(d, 0.6) & "," & SourcesTried(d, -1) & "," _
        & SourcesTried(d, 1.5) & "," & SourcesTried(d, 0.5)
    s = s & "#info:" & InfoTried(d, alpha, "3") & "," & InfoTried(d, alpha, "x") & "," & InfoTried(d, alpha, Null) _
        & "," & InfoTried(d, alpha, 3, 1) & "," & InfoTried(d, alpha, 3, "x") & "," & InfoTried(d, alpha, 3, 5) _
        & "," & InfoTried(d, Replace(alpha, "\\", "/"), 3) & "," & InfoTried(d, "Beta.xlsx", 3) _
        & "," & InfoTried(d, Null, 3) & "," & InfoTried(d, alpha, 1.4) & "," & InfoTried(d, alpha, 3.4) _
        & "," & InfoTried(d, "", 3) & "," & InfoTried(d, alpha, 0) & "," & InfoTried(d, "C:\\nowhere\\Alpha.xlsx", 3)
    Set o = d
    On Error Resume Next
    v = o.LinkInfo(alpha)
    s = s & "," & Err.Number & " " & Err.Description
    Err.Clear
    On Error GoTo 0
    s = s & "#saved listed:" & d.Saved
    d.BreakLink Name:=alpha, Type:=xlLinkTypeExcelLinks
    d.SaveAs Filename:=folder & "\\AfterSave.xlsx", FileFormat:=xlOpenXMLWorkbook
    On Error Resume Next
    d.BreakLink Name:=alpha, Type:=xlLinkTypeExcelLinks
    s = s & "#after save:" & Err.Number & " " & Err.Description
    Err.Clear
    d.BreakLink Name:="Alpha.xlsx", Type:=xlLinkTypeExcelLinks
    s = s & "," & Err.Number & " " & Err.Description & "," & InfoTried(d, alpha, 3)
    On Error GoTo 0
    d.Close False
    Set d = Opened(folder, "Dest.xlsx")
    d.UpdateLinks = 1
    s = s & "#saved update:" & d.Saved
    d.UpdateLinks = 2
    s = s & "," & d.Saved
    d.Close False
    Set d = Opened(folder, "Dest.xlsx")
    d.SaveLinkValues = True
    s = s & "#saved values:" & d.Saved
    d.SaveLinkValues = False
    s = s & "," & d.Saved
    d.Close False
    Set n = Workbooks.Add(xlWBATWorksheet)
    n.UpdateLinks = 2
    n.SaveLinkValues = False
    n.SaveAs Filename:=folder & "\\NoLinks.xlsx", FileFormat:=xlOpenXMLWorkbook
    n.Close False
    Set n = Opened(folder, "NoLinks.xlsx")
    s = s & "#no links:" & n.UpdateLinks & " " & n.SaveLinkValues
    n.Close False
    Application.DisplayAlerts = True
    MoreCases = s
End Function

Public Function ReSaves(ByVal folder As String) As String
    Dim a As Workbook, b As Workbook, d As Workbook, kind As Variant, v As Variant
    Application.DisplayAlerts = False
    Set a = Workbooks.Open(folder & "\\Alpha.xlsx")
    Set b = Workbooks.Open(folder & "\\Beta.xlsx")
    For Each kind In Array("Plain", "Named", "Local", "Mixed", "Blank")
        Set d = Workbooks.Add(xlWBATWorksheet)
        Select Case kind
        Case "Plain"
            d.Worksheets(1).Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1"
        Case "Blank"
            d.Worksheets(1).Range("A1").Formula = "=[Alpha.xlsx]Other!B5"
        Case "Named"
            d.Worksheets(1).Range("A1").Formula = "=Alpha.xlsx!Rate"
        Case "Local"
            d.Names.Add Name:="Ext", RefersTo:="=[Alpha.xlsx]Sheet1!$A$1:$A$3"
        Case "Mixed"
            d.Worksheets(1).Range("A1").Formula = "=[Alpha.xlsx]Sheet1!A1+[Beta.xlsx]Sheet1!A2"
            d.Worksheets(1).Range("A2").Formula = "=[Beta.xlsx]Sheet1!A1"
        End Select
        d.SaveAs Filename:=folder & "\\" & kind & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        d.Close False
    Next
    a.Close False
    b.Close False
    For Each kind In Array("Plain", "Named", "Local", "Mixed", "Blank")
        Set d = Opened(folder, kind & ".xlsx")
        If kind = "Mixed" Then
            v = d.LinkSources()
            d.BreakLink Name:=folder & "\\Alpha.xlsx", Type:=xlLinkTypeExcelLinks
        End If
        d.SaveAs Filename:=folder & "\\" & kind & "Again.xlsx", FileFormat:=xlOpenXMLWorkbook
        d.Close False
    Next
    Set d = Opened(folder, "NamedAgain.xlsx")
    d.SaveAs Filename:=folder & "\\NamedThird.xlsx", FileFormat:=xlOpenXMLWorkbook
    d.Close False
    Application.DisplayAlerts = True
    ReSaves = "ok"
End Function

Public Function DeletedSheet(ByVal folder As String) As String
    Dim b As Workbook, s As String
    Application.DisplayAlerts = False
    Set b = Workbooks.Add(xlWBATWorksheet)
    b.Worksheets(1).Name = "One"
    b.Worksheets.Add(After:=b.Worksheets(1)).Name = "Two"
    b.Names.Add Name:="Cell", RefersTo:="=Two!$A$2"
    b.Names.Add Name:="Block", RefersTo:="=Two!$A$1:$B$3*2"
    b.Names.Add Name:="Mixed", RefersTo:="=Two!$A$1+One!$A$1"
    b.Worksheets("Two").Delete
    s = b.Names("Cell").RefersTo & "|" & b.Names("Block").RefersTo & "|" & b.Names("Mixed").RefersTo
    b.SaveAs Filename:=folder & "\\Gone.xlsx", FileFormat:=xlOpenXMLWorkbook
    b.Close False
    Set b = Workbooks.Open(folder & "\\Gone.xlsx")
    s = s & "#" & b.Names("Cell").RefersTo & "|" & b.Names("Block").RefersTo & "|" & b.Names("Mixed").RefersTo
    b.Close False
    Application.DisplayAlerts = True
    DeletedSheet = s
End Function

Private Function StatusOf(ByVal folder As String, Optional ByVal file As String = "Dest.xlsx") As String
    Dim d As Workbook, v As Variant, i As Long, s As String
    Set d = Opened(folder, file)
    v = d.LinkSources(1)
    For i = LBound(v) To UBound(v)
        s = s & "|" & Mid(v(i), InStrRev(v(i), "\\") + 1) & " " & InfoTried(d, v(i), 3) & " " & InfoTried(d, v(i), 1)
    Next
    If file = "Dest.xlsx" Then s = s & "|A2 " & Shown(d.Worksheets("Dest").Range("A2").Value)
    d.Close False
    StatusOf = Mid(s, 2)
End Function

Public Function LinkStatuses(ByVal folder As String) As String
    Dim a As Workbook, s As String
    Application.DisplayAlerts = False
    s = "as saved:" & StatusOf(folder)
    ' NamedAgain's link lists a sheet Alpha, which Alpha.xlsx never had.
    s = s & "#phantom:" & StatusOf(folder, "NamedAgain.xlsx")
    Set a = Workbooks.Open(folder & "\\Alpha.xlsx")
    a.Worksheets("Sheet1").Range("A1").Value = 99
    a.Save
    a.Close False
    s = s & "#alpha changed:" & StatusOf(folder)
    Name folder & "\\Beta.xlsx" As folder & "\\Beta.moved"
    s = s & "#beta gone:" & StatusOf(folder)
    Name folder & "\\Beta.moved" As folder & "\\Beta.xlsx"
    Set a = Workbooks.Open(folder & "\\Alpha.xlsx")
    a.Worksheets("Other").Name = "Renamed"
    a.Save
    a.Close False
    s = s & "#sheet gone:" & StatusOf(folder)
    ' Plain's link lists Other though no formula reads it; Blank's reads an empty cell of it.
    s = s & "#unread sheet gone:" & StatusOf(folder, "Plain.xlsx") & "#blank cell sheet gone:" _
        & StatusOf(folder, "Blank.xlsx")
    Application.DisplayAlerts = True
    LinkStatuses = s
End Function
'''


def _scrubbed(data: bytes, folder: str) -> str:
    try:
        text = data.decode('utf-8')
    except UnicodeDecodeError:
        return 'base64:' + base64.b64encode(data).decode('ascii')
    for spelling in (folder, folder.replace('\\', '/'), 'file:///' + folder.replace('\\', '/')):
        text = text.replace(spelling, '<folder>')
    return text


def _parts(path: Path, folder: str) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {name: _scrubbed(package.read(name), folder) for name in package.namelist()}


def _deduplicated(files: dict[str, dict[str, str]]) -> dict[str, dict[str, str]]:
    """Each file's parts, a part the same as an earlier file's of its name written as same:<that file>."""
    out: dict[str, dict[str, str]] = {}
    for name, parts in files.items():
        out[name] = {}
        for part, text in parts.items():
            earlier = next((other for other, kept in files.items() if other in out and other != name
                            and kept.get(part) == text), None)
            out[name][part] = f'same:{earlier}' if earlier is not None else text
    return out


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module('LinkSink', SINK, kind='class')
        source = _module()
        record: dict[str, object] = {'measured_at': datetime.now(timezone.utc).isoformat(), 'dest': DEST,
                                     'read': [f'{sheet}!{cell}' for sheet, cell in READ],
                                     # The macros as Excel ran them, for the model to run the same.
                                     'module': source, 'sink': SINK}

        def run(proc: str, *args: str) -> str:
            result = excel.run_vba(source, proc, (folder, *args), timeout=300.0, module_name='LinksProbe')
            if result.outcome == 'runner-error':
                # Now and then the macro fails as it starts (VBA error 40040) with nothing run; once more does.
                print(f'== {proc}: {result.message}; running it again', flush=True)
                result = excel.run_vba(source, proc, (folder, *args), timeout=300.0, module_name='LinksProbe')
            assert result.ok and not result.dialogs, result
            text = str(result.value).replace(folder + '\\', '<folder>\\').replace(folder, '<folder>')
            print(f'== {proc} {" ".join(args)}\n{text}', flush=True)
            return text

        assert run('BuildFiles') == 'ok'
        files = {name: _parts(Path(folder) / name, folder) for name in ('Dest.xlsx', 'Alpha.xlsx', 'Beta.xlsx')}
        record['survey'] = run('SurveyLinks')
        record['events'] = run('BreakEvents')
        for which in ('Alpha', 'Beta', 'All'):
            record[f'break_{which.lower()}'] = run('BreakCase', which, f'Broke{which}.xlsx')
            files[f'Broke{which}.xlsx'] = _parts(Path(folder) / f'Broke{which}.xlsx', folder)
        record['errors'] = run('ErrorCases')
        record['settings'] = run('LinkSettings')
        record['more'] = run('MoreCases')
        for name in ('Never.xlsx', 'ValuesAgain.xlsx', 'Always.xlsx', 'AfterSave.xlsx', 'NoLinks.xlsx'):
            files[name] = _parts(Path(folder) / name, folder)
        assert run('ReSaves') == 'ok'
        # Each workbook as first saved, and of it saved again with its sources closed the parts that name its links.
        for kind in ('Plain', 'Named', 'Local', 'Mixed', 'Blank'):
            files[f'{kind}.xlsx'] = _parts(Path(folder) / f'{kind}.xlsx', folder)
            for name in (f'{kind}Again.xlsx',) + (('NamedThird.xlsx',) if kind == 'Named' else ()):
                files[name] = {part: text for part, text in _parts(Path(folder) / name, folder).items()
                               if 'externalLink' in part or part in ('xl/workbook.xml', 'xl/worksheets/sheet1.xml')}
        record['statuses'] = run('LinkStatuses')
        # A name a deleted sheet left reading #REF!, saved and read back as a broken link's is.
        record['deleted_sheet'] = run('DeletedSheet')
        files['Gone.xlsx'] = _parts(Path(folder) / 'Gone.xlsx', folder)
        record['files'] = _deduplicated(files)
    OUT.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
