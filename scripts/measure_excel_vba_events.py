"""Measure Excel events through WithEvents class sinks: a sheet's Change,
SelectionChange, Calculate, Activate and Deactivate, a workbook's SheetChange and
NewSheet, the Application's, EnableEvents, re-entry, and when Calculate fires
(a sheet, a range, an edit, manual mode).

The harness turns events off, so each probe turns them on and off again.
Volatile formulas recalculated as Calculation goes back to automatic are left
out: whether they are depends on what calculated before, which is not
modelled. Each probe is a module whose ``Function Probe() As String`` the
fixture records (see vba_probe_fixture.py); the sink classes are added to
every session.
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/excel_vba_events.json'

CLASSES = {
    'CSheetSink': '''Option Explicit
Public WithEvents Ws As Worksheet
Public Log As String
Private Sub Ws_Change(ByVal Target As Range)
    Log = Log & "chg:" & Target.Address & ";"
End Sub
Private Sub Ws_SelectionChange(ByVal Target As Range)
    Log = Log & "sel:" & Target.Address & ";"
End Sub
Private Sub Ws_Calculate()
    Log = Log & "calc;"
End Sub
Private Sub Ws_Activate()
    Log = Log & "act;"
End Sub
Private Sub Ws_Deactivate()
    Log = Log & "deact;"
End Sub
''',
    'CBookSink': '''Option Explicit
Public WithEvents Wb As Workbook
Public Log As String
Private Sub Wb_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "sheetchg:" & Target.Address & ";"
End Sub
Private Sub Wb_NewSheet(ByVal Sh As Object)
    Log = Log & "new:" & TypeName(Sh) & ";"
End Sub
Private Sub Wb_SheetActivate(ByVal Sh As Object)
    Log = Log & "sheetact;"
End Sub
''',
    'CAppSink': '''Option Explicit
Public WithEvents App As Application
Public Log As String
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Log = Log & "appchg:" & Target.Address & ";"
End Sub
Private Sub App_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)
    Log = Log & "appnew;"
End Sub
''',
    'CReentry': '''Option Explicit
Public WithEvents Ws As Worksheet
Public Count As Long
Private Sub Ws_Change(ByVal Target As Range)
    Count = Count + 1
    If Count < 5 Then Target.Offset(1, 0).Value = Count
End Sub
''',
}

CASES = {
    'change_single': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("B2").Value = 1\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_block': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1:C2").Value = 1\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_formula': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Formula = "=1+1"\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_clear': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").ClearContents\nws.Range("B1:B3").Clear\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_events_off': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nApplication.EnableEvents = False\nws.Range("A1").Value = 1\nApplication.EnableEvents = True\nws.Range("A2").Value = 1\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_select': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("C3").Select\nws.Range("D4:E5").Select\nProbe = s.Log\nApplication.EnableEvents = False',
    'calculate_event': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Formula = "=NOW()"\nws.Calculate\nProbe = s.Log\nApplication.EnableEvents = False',
    'activate_events': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim other As Worksheet\nSet other = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Activate\nother.Activate\nProbe = s.Log\nApplication.EnableEvents = False',
    'book_sheetchange': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim b As New CBookSink\nSet b.Wb = ThisWorkbook\nws.Range("A1").Value = 1\nWorksheets.Add\nProbe = b.Log\nApplication.EnableEvents = False',
    'app_sheetchange': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim a As New CAppSink\nSet a.App = Application\nws.Range("A1").Value = 1\nWorksheets.Add\nProbe = a.Log\nApplication.EnableEvents = False',
    'reentry': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim r As New CReentry\nSet r.Ws = ws\nws.Range("A1").Value = 0\nProbe = r.Count & "|" & ws.Range("A5").Value & "|" & ws.Range("A6").Value\nApplication.EnableEvents = False',
    'unhook': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 1\nSet s.Ws = Nothing\nws.Range("A2").Value = 1\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_insert_delete': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Rows(2).Insert\nws.Columns(1).Delete\nProbe = s.Log\nApplication.EnableEvents = False',
    'change_copy_paste': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Copy ws.Range("C1:C2")\nProbe = s.Log\nApplication.EnableEvents = False',
    'calc_no_formulas': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_plain_formula': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Formula = "=1+1"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_volatile': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Formula = "=RAND()"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Calculate\nws.Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_application': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Formula = "=RAND()"\nDim s As New CSheetSink\nSet s.Ws = ws\nApplication.Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_range': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Formula = "=RAND()"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_dependent_write': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nws.Range("B1").Formula = "=A1*2"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_independent_write': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("B1").Formula = "=1+1"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_volatile_write': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("B1").Formula = "=RAND()"\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'calc_empty_sheet': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Calculate\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'manual_dependent_write': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nws.Range("B1").Formula = "=A1*2"\nApplication.Calculation = xlCalculationManual\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nws.Calculate\nApplication.Calculation = xlCalculationAutomatic\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'manual_independent_write': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("B1").Formula = "=1+1"\nApplication.Calculation = xlCalculationManual\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nws.Calculate\nApplication.Calculation = xlCalculationAutomatic\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'manual_no_formulas': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nApplication.Calculation = xlCalculationManual\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nws.Calculate\nApplication.Calculation = xlCalculationAutomatic\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'manual_range_calc': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("B1").Formula = "=RAND()"\nApplication.Calculation = xlCalculationManual\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("B1").Calculate\nApplication.Calculation = xlCalculationAutomatic\nProbe = "[" & s.Log & "]"\nApplication.EnableEvents = False',
    'manual_back_to_auto': 'Application.EnableEvents = True\nDim ws As Worksheet\nSet ws = Worksheets.Add\nws.Range("A1").Value = 1\nws.Range("B1").Formula = "=A1*2"\nApplication.Calculation = xlCalculationManual\nDim s As New CSheetSink\nSet s.Ws = ws\nws.Range("A1").Value = 2\nApplication.Calculation = xlCalculationAutomatic\nProbe = "[" & s.Log & "]|" & ws.Range("B1").Value\nApplication.EnableEvents = False',
}


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT, CLASSES)


if __name__ == '__main__':
    main()
