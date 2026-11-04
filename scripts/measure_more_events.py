"""The events the model did not raise: BeforeDelete, NewChart, the window events, BeforePrint and FollowHyperlink.

Every handler writes to one log, so a scenario's answer gives the order in which sinks and the document modules
hear each event, among the sheet and workbook events the model already raises. Every print goes to a file in the TEMP
folder: nothing reaches a printer. Excel saves its window's state and size for its user as it exits, so every row
starts from a normal window and puts its size back.

    python scripts/measure_more_events.py

writes tests/fixtures/more_events.json, which tests/test_excel_more_events.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / "tests/fixtures/more_events.json"

#: A workbook event: its name, its other parameters, what the log says of them, and its name on an Application.
BOOK_EVENTS = [
    ("Activate", "", '""', "WorkbookActivate"),
    ("Deactivate", "", '""', "WorkbookDeactivate"),
    ("WindowActivate", "ByVal Wn As Window", "Wn.Caption", "WindowActivate"),
    ("WindowDeactivate", "ByVal Wn As Window", "Wn.Caption", "WindowDeactivate"),
    ("WindowResize", "ByVal Wn As Window", "Wn.Caption", "WindowResize"),
    ("NewChart", "ByVal Ch As Chart", 'Ch.Name & "/" & TypeName(Ch.Parent)', "WorkbookNewChart"),
    ("NewSheet", "ByVal Sh As Object", "Sh.Name", "WorkbookNewSheet"),
    ("SheetBeforeDelete", "ByVal Sh As Object", "Sh.Name", "SheetBeforeDelete"),
    ("SheetActivate", "ByVal Sh As Object", "Sh.Name", "SheetActivate"),
    ("SheetDeactivate", "ByVal Sh As Object", "Sh.Name", "SheetDeactivate"),
    ("SheetSelectionChange", "ByVal Sh As Object, ByVal Target As Range",
     'Sh.Name & "/" & Target.Address(False, False)', "SheetSelectionChange"),
    ("SheetFollowHyperlink", "ByVal Sh As Object, ByVal Target As Hyperlink",
     'Sh.Name & "/" & Target.SubAddress & "/" & ActiveCell.Address(False, False)', "SheetFollowHyperlink"),
    ("BeforePrint", "Cancel As Boolean", "ActiveSheet.Name", "WorkbookBeforePrint"),
    ("BeforeClose", "Cancel As Boolean", '""', "WorkbookBeforeClose"),
]
SHEET_EVENTS = [
    ("BeforeDelete", "", '""'),
    ("Activate", "", '""'),
    ("Deactivate", "", '""'),
    ("SelectionChange", "ByVal Target As Range", "Target.Address(False, False)"),
    ("FollowHyperlink", "ByVal Target As Hyperlink", "Target.SubAddress"),
]


def handler(procedure: str, signature: str, label: str, said: str) -> str:
    cancel = "Cancel = WantCancel\n" if "Cancel" in signature else ""
    return f'Private Sub {procedure}({signature})\nLog = Log & "{label}:" & {said} & "|"\n{cancel}End Sub\n'


SINK = "Public WithEvents Sheet As Worksheet\nPublic WithEvents Book As Workbook\nPublic WithEvents App As Application\n"
OWN_SHEET = ""
OWN_BOOK = ""
for event, signature, said in SHEET_EVENTS:
    SINK += handler(f"Sheet_{event}", signature, f"sheet-sink-{event}", said)
    OWN_SHEET += handler(f"Worksheet_{event}", signature, f"sheet-own-{event}", said)
for event, signature, said, on_app in BOOK_EVENTS:
    SINK += handler(f"Book_{event}", signature, f"book-sink-{event}", said)
    OWN_BOOK += handler(f"Workbook_{event}", signature, f"book-own-{event}", said)
    if on_app.startswith("Sheet"):
        SINK += handler(f"App_{on_app}", signature, f"app-sink-{on_app}", said)
    else:
        with_book = "ByVal Wb As Workbook" + (", " + signature if signature else "")
        SINK += handler(f"App_{on_app}", with_book, f"app-sink-{on_app}", f'Wb.Name & "/" & {said}')

SINK += handler("App_NewWorkbook", "ByVal Wb As Workbook", "app-sink-NewWorkbook", "Wb.Name")

TO_FILE = 'PrintToFile:=True, PrToFileName:=Environ("TEMP") & "\\pyopenvba_print_probe.prn"'
SECOND = 'Worksheets.Add After:=Sheets(1)\n'
CANCEL = 'WantCancel = True\n'
ACTIONS = {
    # Deleting a sheet.
    "delete_active": f'{SECOND}Sheets("Sheet1").Activate\nLog = ""\nSheets("Sheet1").Delete',
    "delete_inactive": f'{SECOND}Log = ""\nSheets("Sheet1").Delete',
    "delete_active_no_module": f'{SECOND}Log = ""\nActiveSheet.Delete',
    "delete_group": f'{SECOND}{SECOND}Log = ""\nSheets(Array("Sheet1", "Sheet2")).Delete',
    "delete_group_reversed": f'{SECOND}{SECOND}Log = ""\nSheets(Array("Sheet2", "Sheet1")).Delete',
    "delete_group_with_active": f'{SECOND}{SECOND}Log = ""\nSheets(Array("Sheet1", "Sheet3")).Delete',
    "delete_selected_group": f'{SECOND}{SECOND}Sheets(Array("Sheet1", "Sheet2")).Select\nLog = ""\n'
                             'ActiveWindow.SelectedSheets.Delete',
    "delete_only_sheet": 'Log = ""\nSheets(1).Delete',
    "delete_every_sheet": f'{SECOND}Log = ""\nSheets.Delete',
    "delete_protected_structure": f'{SECOND}ActiveWorkbook.Protect Structure:=True\nLog = ""\nSheets(1).Delete',
    "delete_events_off": f'{SECOND}Log = ""\nApplication.EnableEvents = False\nSheets("Sheet1").Delete',
    "delete_chart_sheet": 'Charts.Add\nLog = ""\nActiveChart.Delete',
    # Making a chart.
    "chart_sheet": 'Log = ""\nCharts.Add',
    "chart_sheet_2": 'Log = ""\nCharts.Add2',
    "chart_sheets_add": 'Log = ""\nSheets.Add Type:=xlChart',
    "chart_shape": 'Log = ""\nActiveSheet.Shapes.AddChart',
    "chart_shape_2": 'Log = ""\nActiveSheet.Shapes.AddChart2',
    "chart_object": 'Log = ""\nActiveSheet.ChartObjects.Add 10, 10, 200, 100',
    "chart_duplicate": 'ActiveSheet.ChartObjects.Add 10, 10, 200, 100\nLog = ""\nActiveSheet.ChartObjects(1).Duplicate',
    "chart_to_sheet": 'ActiveSheet.ChartObjects.Add 10, 10, 200, 100\nLog = ""\n'
                      'ActiveSheet.ChartObjects(1).Chart.Location xlLocationAsNewSheet',
    "chart_sheet_copy": 'ActiveSheet.ChartObjects.Add 10, 10, 200, 100\nLog = ""\nSheets(1).Copy After:=Sheets(1)',
    "chart_events_off": 'Log = ""\nApplication.EnableEvents = False\nCharts.Add',
    # Windows.
    "window_new_book": 'Log = ""\nWorkbooks.Add',
    "window_back": 'Set other = Workbooks.Add\nLog = ""\nhome.Activate',
    "window_by_windows": 'Set other = Workbooks.Add\nLog = ""\nWindows(home.Name).Activate',
    "window_next": 'Set other = Workbooks.Add\nLog = ""\nActiveWindow.ActivateNext',
    "window_close_active": 'Set other = Workbooks.Add\nLog = ""\nother.Close False',
    "window_close_behind": 'Set other = Workbooks.Add\nhome.Activate\nLog = ""\nother.Close False',
    "window_cancel_close": 'Set other = Workbooks.Add\nWantCancel = True\nLog = ""\nother.Close False',
    "window_close_by_window": 'Set other = Workbooks.Add\nLog = ""\nActiveWindow.Close False',
    "window_open_saved": 'Set other = Workbooks.Add\nother.SaveAs Environ("TEMP") & "\pyopenvba_open_probe.xlsx", 51\n'
                         'other.Close\nLog = ""\nWorkbooks.Open Environ("TEMP") & "\pyopenvba_open_probe.xlsx"',
    "window_sheet_copy_new_book": f'{SECOND}Log = ""\nSheets("Sheet2").Copy',
    "window_sheet_move_new_book": f'{SECOND}Log = ""\nSheets("Sheet2").Move',
    "window_same": 'Log = ""\nActiveWindow.Activate\nhome.Activate\nSheets(1).Activate',
    "window_sheet_only": f'{SECOND}Log = ""\nSheets(1).Activate',
    "window_new": 'Log = ""\nActiveWindow.NewWindow',
    "window_new_then_first": 'ActiveWindow.NewWindow\nLog = ""\nhome.Windows(1).Close',
    "window_events_off": 'Log = ""\nApplication.EnableEvents = False\nWorkbooks.Add',
    "resize_width": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.Width = 300',
    "resize_width_same": 'ActiveWindow.WindowState = xlNormal\nActiveWindow.Width = 300\nLog = ""\n'
                         'ActiveWindow.Width = 300',
    "resize_height": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.Height = 200',
    "resize_width_read_back": 'ActiveWindow.WindowState = xlNormal\nActiveWindow.Width = 300\nLog = ""\n'
                              'ActiveWindow.Width = ActiveWindow.Width',
    "resize_width_again": 'ActiveWindow.WindowState = xlNormal\nActiveWindow.Width = 300\nLog = ""\n'
                          'ActiveWindow.Width = 320',
    "resize_other_book": 'Set other = Workbooks.Add\nhome.Activate\nother.Windows(1).WindowState = xlNormal\nLog = ""\n'
                         'other.Windows(1).Width = 300',
    "state_same": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.WindowState = xlNormal',
    "move_left": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.Left = 40',
    "state_maximized": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.WindowState = xlMaximized',
    "state_minimized": 'ActiveWindow.WindowState = xlNormal\nLog = ""\nActiveWindow.WindowState = xlMinimized',
    # Printing.
    "print_sheet": f'{CANCEL}Range("A1").Value = 1\nLog = ""\nActiveSheet.PrintOut {TO_FILE}',
    "print_book": f'{CANCEL}Range("A1").Value = 1\nLog = ""\nActiveWorkbook.PrintOut {TO_FILE}',
    "print_range": f'{CANCEL}Range("A1").Value = 1\nLog = ""\nRange("A1").PrintOut {TO_FILE}',
    "print_other_sheet": f'{CANCEL}Range("A1").Value = 1\n{SECOND}Log = ""\nSheets("Sheet1").PrintOut {TO_FILE}',
    "print_chart": f'{CANCEL}Charts.Add\nLog = ""\nActiveChart.PrintOut {TO_FILE}',
    "print_through": f'Range("A1").Value = 1\nLog = ""\nActiveSheet.PrintOut {TO_FILE}',
    "print_through_other_sheet": f'Range("A1").Value = 1\n{SECOND}Log = ""\nSheets("Sheet1").PrintOut {TO_FILE}',
    "print_through_book": f'Range("A1").Value = 1\n{SECOND}Range("A1").Value = 1\nLog = ""\n'
                          f'ActiveWorkbook.PrintOut {TO_FILE}',
    "print_window": f'{CANCEL}Range("A1").Value = 1\nLog = ""\nActiveWindow.PrintOut {TO_FILE}',
    "print_empty": f'{CANCEL}Log = ""\nActiveSheet.PrintOut {TO_FILE}',
    "print_empty_through": f'Log = ""\nActiveSheet.PrintOut {TO_FILE}',
    "print_hidden": f'{CANCEL}Range("A1").Value = 1\n{SECOND}Sheets("Sheet1").Visible = False\nLog = ""\n'
                    f'Sheets("Sheet1").PrintOut {TO_FILE}',
    "print_other_book": f'{CANCEL}Range("A1").Value = 1\nSet other = Workbooks.Add\nLog = ""\n'
                        f'home.Sheets(1).PrintOut {TO_FILE}',
    # A link inside the workbook: nothing leaves the machine.
    "follow_inside": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "Sheet1!C3"\nLog = ""\n'
                     'ActiveSheet.Hyperlinks(1).Follow',
    "follow_range": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "Sheet1!C3"\nLog = ""\n'
                    'Range("A1").Hyperlinks(1).Follow NewWindow:=False',
    "follow_other_sheet": f'{SECOND}ActiveSheet.Hyperlinks.Add Range("A1"), "", "Sheet1!C3"\nLog = ""\n'
                          'ActiveSheet.Hyperlinks(1).Follow',
    "follow_nowhere": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "Nowhere!C3"\nLog = ""\n'
                      'ActiveSheet.Hyperlinks(1).Follow',
    "follow_unqualified": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "C3"\nLog = ""\nActiveSheet.Hyperlinks(1).Follow',
    "follow_name": 'ActiveWorkbook.Names.Add "Spot", "=Sheet1!$D$4"\nActiveSheet.Hyperlinks.Add Range("A1"), "", "Spot"\n'
                   'Log = ""\nActiveSheet.Hyperlinks(1).Follow',
    "follow_behind": f'Sheets(1).Hyperlinks.Add Sheets(1).Range("A1"), "", "C3"\n{SECOND}Log = ""\n'
                     'Sheets("Sheet1").Hyperlinks(1).Follow',
    "follow_same_cell": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "Sheet1!A1"\nLog = ""\nActiveSheet.Hyperlinks(1).Follow',
    "follow_events_off": 'ActiveSheet.Hyperlinks.Add Range("A1"), "", "Sheet1!C3"\nLog = ""\n'
                         'Application.EnableEvents = False\nActiveSheet.Hyperlinks(1).Follow',
}

PROBE = '''Function Probe() As String
Application.EnableEvents = False
Application.DisplayAlerts = False
Dim sink As New ExcelSink, home As Workbook, other As Workbook, book As Workbook, failure As String
Dim wide As Double, high As Double, across As Double, down As Double
Set home = ActiveWorkbook
Application.WindowState = xlNormal
wide = ActiveWindow.Width
high = ActiveWindow.Height
across = ActiveWindow.Left
down = ActiveWindow.Top
Set sink.Sheet = ActiveSheet
Set sink.Book = ActiveWorkbook
Set sink.App = Application
Application.EnableEvents = True
Log = ""
On Error GoTo Failed
{action}
GoTo Done
Failed:
failure = "!E" & Err.Number
Done:
On Error Resume Next
Probe = Log & failure & "#" & ActiveWorkbook.Name & "/" & ActiveSheet.Name & "/" & ActiveWindow.Caption & "/" & Windows.Count
Application.EnableEvents = False
For Each book In Workbooks
If book.Name <> home.Name Then book.Close False
Next
home.Activate
Application.WindowState = xlNormal
ActiveWindow.Width = wide
ActiveWindow.Height = high
ActiveWindow.Left = across
ActiveWindow.Top = down
If Dir(Environ("TEMP") & "\\pyopenvba_open_probe.xlsx") <> "" Then Kill Environ("TEMP") & "\\pyopenvba_open_probe.xlsx"
End Function
'''


LOG = "Public Log As String\nPublic WantCancel As Boolean\n"


def install_lines() -> str:
    lines = ['Function Install() As String', 'Dim text As String, part As Object']
    for owner, code in ((None, SINK), ('Worksheets(1).CodeName', OWN_SHEET), ('"ThisWorkbook"', OWN_BOOK)):
        if owner is None:
            lines += ['Set part = ActiveWorkbook.VBProject.VBComponents.Add(2)', 'part.Name = "ExcelSink"']
        else:
            lines += [f'Set part = ActiveWorkbook.VBProject.VBComponents({owner})']
        lines += ['text = ""']
        lines += ['text = text & "' + line.replace('"', '""') + '" & vbLf' for line in code.splitlines()]
        lines += ['part.CodeModule.AddFromString text']
    lines += ['Install = Application.Version & "|" & Application.Build', 'End Function']
    return "\n".join(lines)


def main() -> None:
    rows: list[dict[str, str]] = []
    version = ""
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        for name, action in ACTIONS.items():
            excel.new_document()
            excel.add_module("EventLog", LOG)
            installed = excel.run_vba(install_lines(), "Install", timeout=30.0)
            assert installed.ok, installed
            version = str(installed.value)
            result = excel.run_vba(PROBE.format(action=action), "Probe", timeout=60.0)
            print(name, result.value if result.ok else result, flush=True)
            rows.append({"name": name, "action": action, "result": str(result.value) if result.ok else f"failed {result}"})
    OUT.write_text(json.dumps({"measured_at": datetime.now(timezone.utc).isoformat(), "version_build": version,
                               "log": LOG, "sink": SINK, "sheet": OWN_SHEET, "book": OWN_BOOK, "probe": PROBE,
                               "rows": rows}, indent=1) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
