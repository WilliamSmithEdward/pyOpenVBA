"""Which events Excel raises, in what order and with what, and what a document module's code sees.

A fresh workbook gets a second sheet, event handlers in both sheets'
modules and in ThisWorkbook that note each event into a log in a
standard module, and a formula on the first sheet so that recalculation
shows. Each action then runs on its own, the log emptied before it, and
what the log holds after it is its answer. The handlers are written
through the VBA project's CodeModule, which needs "Trust access to the
VBA project object model".

The same workbook answers what code in a sheet's or the workbook's own
module reaches by an unqualified name -- its own Range, Cells and Name,
or the active sheet's -- and what outside code reaches through
Sheet2.Member.

    python scripts/measure_events.py

writes tests/fixtures/events.json, which tests/test_excel_events.py replays.
"""

from __future__ import annotations

import json
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "events.json"

LOG = """Public Events As String

Public Sub Note(ByVal text As String)
    Events = Events & text & "|"
End Sub
"""

#: The code each module gets: the first sheet's, the second sheet's and the workbook's.
SHEET1 = """Private Sub Worksheet_Change(ByVal Target As Range)
    Note "Change:" & Target.Address(False, False)
End Sub

Private Sub Worksheet_Calculate()
    Note "Calculate"
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Note "SelectionChange:" & Target.Address(False, False)
End Sub

Private Sub Worksheet_Activate()
    Note "Activate"
End Sub

Private Sub Worksheet_Deactivate()
    Note "Deactivate"
End Sub
"""

SHEET2 = """Public Counter As Long

Private Sub Worksheet_Change(ByVal Target As Range)
    Note "Sheet2.Change:" & Target.Address(False, False)
    If Target.Address(False, False) = "A1" And Counter < 3 Then
        Counter = Counter + 1
        Range("B1").Value = Target.Value * 2
    End If
End Sub

Private Sub Worksheet_Activate()
    Note "Sheet2.Activate"
End Sub

Private Sub Worksheet_Deactivate()
    Note "Sheet2.Deactivate"
End Sub

Public Function Where() As String
    Where = Name & "|" & Me.Name & "|" & TypeName(Me) & "|" & ActiveSheet.Name & "|" & _
        Range("A1").Parent.Name & "|" & Cells(1, 1).Parent.Name & "|" & Rows.Count
End Function

Public Sub WriteHere()
    Range("C1").Value = 7
End Sub

Private Function Secret() As String
    Secret = "secret"
End Function
"""

BOOK = """Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Note "SheetChange:" & Sh.Name & "!" & Target.Address(False, False)
End Sub

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    Note "SheetCalculate:" & Sh.Name
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Note "SheetSelectionChange:" & Sh.Name & "!" & Target.Address(False, False)
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Note "SheetActivate:" & Sh.Name
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    Note "SheetDeactivate:" & Sh.Name
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Note "NewSheet:" & Sh.Name
End Sub

Public Function Where() As String
    Where = Name & "|" & Me.Name & "|" & TypeName(Me) & "|" & Sheets.Count & "|" & Worksheets(1).Name
End Function
"""

#: Each action, run with Sheet1 active unless it changes that, the log emptied before it.
ACTIONS: list[tuple[str, str]] = [
    ("value", 'Sheet1.Range("A1").Value = 1'),
    ("same_value", 'Sheet1.Range("A1").Value = 1'),
    ("block", 'Sheet1.Range("A2:B3").Value = 5'),
    ("areas", 'Sheet1.Range("A2,B5").Value = 6'),
    ("formula", 'Sheet1.Range("A2").Formula = "=1+1"'),
    ("value2", 'Sheet1.Range("A1").Value2 = 2'),
    ("formula_r1c1", 'Sheet1.Range("A1").FormulaR1C1 = "=3"'),
    ("default_member", "Sheet1.Cells(1, 1) = 4"),
    ("clear_contents", 'Sheet1.Range("A2:B3").ClearContents'),
    ("clear_empty", 'Sheet1.Range("F1").ClearContents'),
    ("number_format", 'Sheet1.Range("A2").NumberFormat = "0.00"'),
    ("font", 'Sheet1.Range("A2").Font.Bold = True'),
    ("clear", 'Sheet1.Range("A2:B3").Clear'),
    ("insert_row", "Sheet1.Rows(10).Insert"),
    ("delete_row", "Sheet1.Rows(10).Delete"),
    ("insert_cells", 'Sheet1.Range("J1").Insert Shift:=xlDown'),
    ("delete_cells", 'Sheet1.Range("J1").Delete Shift:=xlUp'),
    ("copy", 'Sheet1.Range("A1").Copy Sheet1.Range("F1")'),
    ("autofill", 'Sheet1.Range("A1").AutoFill Sheet1.Range("A1:A3")'),
    ("fill_down", 'Sheet1.Range("A1:A3").FillDown'),
    ("replace", 'Sheet1.Range("A1:A3").Replace "4", "5"'),
    ("sort", 'Sheet1.Range("A1:A3").Sort Key1:=Sheet1.Range("A1")'),
    ("formula_array", 'Sheet1.Range("G1:G2").FormulaArray = "=A1:A2*2"'),
    ("merge", 'Sheet1.Range("H1:H2").Merge'),
    ("events_off", 'Application.EnableEvents = False: Sheet1.Range("A1").Value = 9: '
                   "Application.EnableEvents = True"),
    ("calculate", "Application.Calculate"),
    ("select", 'Sheet1.Range("A5").Select'),
    ("select_again", 'Sheet1.Range("A5").Select'),
    ("activate_other", "Sheet2.Activate"),
    ("activate_back", "Sheet1.Activate"),
    ("other_sheet", 'Sheet2.Range("A1").Value = 3'),
    ("add_sheet", "Worksheets.Add After:=Worksheets(Worksheets.Count)"),
    ("activate_first", "Sheet1.Activate"),
    ("add_sheet_plain", "Worksheets.Add"),
    ("activate_first_again", "Sheet1.Activate"),
]

#: What code reaches through a document module, read after the actions.
REACH: list[tuple[str, str]] = [
    ("sheet_where", "Sheet2.Where()"),
    ("book_where", "ThisWorkbook.Where()"),
    ("sheet_write", 'Sheet2.WriteHere: Reach = Sheet2.Range("C1").Value & "|" & Sheet1.Range("C1").Value'),
    ("sheet_variable", "Sheet2.Counter = 5: Reach = Sheet2.Counter"),
    ("run", 'Sheet2.Range("C1").ClearContents: Application.Run "Sheet2.WriteHere": Reach = Sheet2.Range("C1").Value'),
    ("private", "Dim late As Object: Set late = Sheet2: Reach = late.Secret()"),
    ("code_names", 'Reach = Worksheets(1).CodeName & "|" & Worksheets(2).CodeName & "|" & ThisWorkbook.CodeName'),
]


def _building(code: str) -> list[str]:
    """Statements that build ``code`` up in the variable ``text``, a line at a time: VBA takes no line past 1,023
    characters, so the whole code cannot be one literal."""
    return ['text = ""'] + [f'text = text & "{line.replace(chr(34), chr(34) * 2)}" & vbLf'
                            for line in code.splitlines()]


def install_module() -> str:
    lines = ["Public Function Install() As String", "Dim project As Object, text As String",
             "Worksheets.Add After:=Worksheets(1)", "Worksheets(1).Activate",
             'Worksheets(1).Range("D1").Formula = "=A1*2"', "Set project = ActiveWorkbook.VBProject"]
    for owner, code in (("Worksheets(1).CodeName", SHEET1), ("Worksheets(2).CodeName", SHEET2),
                        ('"ThisWorkbook"', BOOK)):
        lines += [*_building(code), f"project.VBComponents({owner}).CodeModule.AddFromString text"]
    lines += ['Install = Worksheets(1).CodeName & "|" & Worksheets(2).CodeName', "End Function"]
    return "\n".join(lines) + "\n"


def actions_module() -> str:
    # The harness runs with events off; the actions are about what Excel raises when they are on.
    lines = ["Public Function Probe() As String", "Dim out As String", "Application.EnableEvents = True"]
    lines += [f'out = out & "{name}=" & A{index}() & vbLf' for index, (name, _) in enumerate(ACTIONS)]
    lines += [f'out = out & "{name}=" & R{index}() & vbLf' for index, (name, _) in enumerate(REACH)]
    lines += ["Probe = out", "End Function"]
    for index, (_, action) in enumerate(ACTIONS):
        lines += [f"Private Function A{index}() As String", "On Error GoTo Bad", 'Events = ""', action,
                  f"A{index} = Events", "Exit Function", "Bad:", f'A{index} = Events & "!" & Err.Number',
                  "End Function"]
    for index, (_, reach) in enumerate(REACH):
        body = reach if "Reach =" in reach else f"Reach = {reach}"
        lines += [f"Private Function R{index}() As String", "Dim Reach As Variant", "On Error GoTo Bad", body,
                  f"R{index} = Reach", "Exit Function", "Bad:", f'R{index} = "!" & Err.Number', "End Function"]
    return "\n".join(lines) + "\n"


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.add_module("EventLog", LOG)
        installed = excel.run_vba(install_module(), "Install", timeout=120.0, module_name="Installer")
        assert installed.ok, f"{installed.outcome}: {installed.message} {installed.error}"
        result = excel.run_vba(actions_module(), "Probe", timeout=300.0, module_name="Actions")
        assert result.ok, f"{result.outcome}: {result.message} {result.error}"
    answers = dict(line.split("=", 1) for line in str(result.value).split("\n") if "=" in line)
    record = {"log": LOG, "sheet1": SHEET1, "sheet2": SHEET2, "book": BOOK, "code_names": str(installed.value),
              "actions": [{"name": name, "action": action, "events": answers[name]} for name, action in ACTIONS],
              "reach": [{"name": name, "code": code, "answer": answers[name]} for name, code in REACH]}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("code names", installed.value)
    for one in record["actions"]:
        print(f"{one['name']:16} {one['events']}")
    for one in record["reach"]:
        print(f"{one['name']:16} {one['answer']}")


if __name__ == "__main__":
    main()
