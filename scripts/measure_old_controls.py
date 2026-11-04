"""The sheet's old control collections, Buttons, CheckBoxes, OptionButtons, DropDowns and ListBoxes: what the collections and
their items answer, in live Excel. Each case makes a workbook of one sheet and runs its statements with o, a control of the
kind, added once ADD has run. Sizes are multiples of 1.5 points, whole pixels at both 96 and 144 DPI.

    python scripts/measure_old_controls.py

writes tests/fixtures/old_controls.json, which tests/test_excel_old_controls.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/old_controls.json'
KINDS = {'Buttons': 'Button', 'CheckBoxes': 'CheckBox', 'OptionButtons': 'OptionButton', 'DropDowns': 'DropDown', 'ListBoxes': 'ListBox'}
SHARED_READS = ('Name', 'Index', 'Left', 'Top', 'Width', 'Height', 'Enabled', 'Visible', 'Locked', 'Placement', 'PrintObject', 'OnAction', 'ZOrder', 'Caption', 'Text', 'LockedText', 'Value', 'LinkedCell', 'ListFillRange', 'ListCount', 'ListIndex', 'DropDownLines', 'MultiSelect', 'Display3DShading', 'Accelerator', 'AutoSize', 'AddIndent', 'AutoScaleFont', 'CancelButton', 'DefaultButton', 'DismissButton', 'HelpButton', 'HorizontalAlignment', 'VerticalAlignment', 'Orientation', 'ReadingOrder', 'Formula', 'PhoneticAccelerator')
SHARED_TYPES = ('Parent', 'Application', 'ShapeRange', 'TopLeftCell', 'BottomRightCell', 'Characters', 'Font', 'Border', 'Interior', 'List', 'Selected')
SETS = (('Name', '"mine"'), ('Left', '96'), ('Top', '144'), ('Width', '96'), ('Height', '48'), ('Enabled', 'False'), ('Visible', 'False'), ('Locked', 'False'), ('Placement', 'xlMove'), ('Placement', 'xlFreeFloating'), ('PrintObject', 'False'), ('OnAction', '"Macro1"'), ('OnAction', '""'), ('Caption', '"Hello"'), ('Text', '"Hello"'), ('LockedText', 'False'), ('Value', 'xlOn'), ('Value', '1'), ('Value', '0'), ('Value', '2'), ('LinkedCell', '"B2"'), ('ListFillRange', '"C1:C3"'), ('ListIndex', '2'), ('DropDownLines', '3'), ('MultiSelect', 'xlSimple'), ('Display3DShading', 'True'), ('Accelerator', '"H"'), ('AutoSize', 'True'), ('HorizontalAlignment', 'xlCenter'), ('VerticalAlignment', 'xlTop'), ('Orientation', 'xlUpward'))
CASES: list[str] = []
for many, one in KINDS.items():
    ADD = f'Set o = ws.{many}.Add(48, 96, 72, 24): '
    CASES += [f'r = ws.{many}.Count', f'r = TypeName(ws.{many}) & " " & TypeName(ws.{many}.Parent) & " " & TypeName(ws.{many}.Application)', f'r = TypeName(ws.{many}(1))', f'r = TypeName(ws.{many}("nope"))', f'{ADD}r = TypeName(o) & " " & o.Name & " " & ws.{many}.Count & " " & ws.Shapes.Count & " " & ws.Shapes(1).Name & " " & ws.DrawingObjects.Count',
              f'{ADD}r = ws.{many}(1).Name & " " & ws.{many}(o.Name).Index & " " & ws.{many}.Item(1).Name', f'{ADD}r = TypeName(ws.{many}(2))', f'{ADD}r = TypeName(ws.{many}("nope"))', f'{ADD}Set o = ws.{many}.Add(48, 144, 72, 24): r = o.Name & " " & o.Index & " " & ws.{many}.Count', f'{ADD}n = 0: For Each c In ws.{many}: n = n + 1: t = TypeName(c): Next: r = n & " " & t',
              f'{ADD}r = TypeName(ws.Shapes(1).OLEFormat) & " " & TypeName(ws.Shapes(1).OLEFormat.Object) & " " & ws.Shapes(1).Type & " " & ws.Shapes(1).FormControlType & " " & ws.Shapes(1).ControlFormat.Enabled', f'{ADD}r = TypeName(ws.DrawingObjects(1)) & " " & TypeName(ws.DrawingObjects)',
              f'Set o = ws.{many}.Add(48, 96, 72): r = o.Name', f'Set o = ws.{many}.Add(Left:=48, Top:=96, Width:=72, Height:=24): r = o.Left & " " & o.Top & " " & o.Width & " " & o.Height', f'Set o = ws.{many}.Add(-48, 96, 72, 24): r = o.Left', f'Set o = ws.{many}.Add(48, 96, 0, 0): r = o.Width & " " & o.Height', f'Set o = ws.{many}.Add("a", 96, 72, 24): r = o.Name']
    CASES += [f'{ADD}r = o.{name}' for name in SHARED_READS]
    CASES += [f'{ADD}r = TypeName(o.{name})' for name in SHARED_TYPES]
    CASES += [f'{ADD}o.{name} = {value}: r = o.{name}' for name, value in SETS]
    CASES += [f'{ADD}r = o.Characters.Text & "|" & o.Characters.Count', f'{ADD}o.Characters.Text = "Press me": r = o.Caption & "|" & o.Text & "|" & o.Characters.Text', f'{ADD}o.Caption = "Hello": r = o.Characters(1, 2).Text & "|" & o.Characters(2).Text', f'{ADD}o.Caption = "Hello": o.Characters(1, 1).Text = "J": r = o.Caption', f'{ADD}r = TypeName(o.Characters.Font) & " " & o.Characters.Font.Bold', f'{ADD}o.Font.Bold = True: r = o.Font.Bold & " " & o.Characters.Font.Bold', f'{ADD}o.Font.Name = "Arial": r = o.Font.Name', f'{ADD}o.Font.Size = 14: r = o.Font.Size', f'{ADD}o.Font.ColorIndex = 3: r = o.Font.ColorIndex',
              f'{ADD}r = o.Delete & " " & ws.{many}.Count', f'{ADD}ws.{many}.Delete: r = ws.{many}.Count & " " & ws.Shapes.Count', f'{ADD}r = o.Select & " " & TypeName(Selection) & " " & Selection.Name', f'{ADD}ws.{many}.Select: r = TypeName(Selection)', f'{ADD}Set o = o.Duplicate: r = TypeName(o) & " " & o.Name & " " & ws.{many}.Count', f'{ADD}r = o.BringToFront & " " & o.SendToBack & " " & o.ZOrder', f'{ADD}Set c = ws.{many}.Add(48, 144, 72, 24): o.BringToFront: r = o.ZOrder & " " & c.ZOrder & " " & o.Index & " " & c.Index', f'{ADD}r = TypeName(o.TopLeftCell) & " " & TypeName(o.ShapeRange) & " " & o.ShapeRange.Count & " " & o.ShapeRange.Name',
              f'{ADD}ws.Rows(1).Insert: r = o.Top', f'{ADD}ws.Shapes(1).Name = "viaShape": r = o.Name & " " & ws.{many}("viaShape").Name', f'{ADD}ws.Shapes(1).Delete: r = ws.{many}.Count', f'{ADD}ws.Shapes(1).OnAction = "Macro2": r = o.OnAction', f'{ADD}o.OnAction = "Macro3": r = ws.Shapes(1).OnAction', f'{ADD}o.Left = 96: r = ws.Shapes(1).Left', f'{ADD}ws.{many}.Left = 96: r = o.Left', f'{ADD}ws.{many}.Caption = "All": r = o.Caption', f'{ADD}ws.{many}.Enabled = False: r = o.Enabled', f'{ADD}ws.{many}.OnAction = "Macro4": r = o.OnAction', f'{ADD}ws.{many}.Visible = False: r = o.Visible', f'{ADD}r = ws.{many}.Left', f'{ADD}r = ws.{many}.Caption', f'{ADD}r = ws.{many}.Enabled', f'{ADD}Set c = ws.{many}(Array(1)): r = TypeName(c) & " " & c.Count', f'{ADD}Set c = ws.{many}.Group: r = TypeName(c)']
CASES += ['Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.AddItem "b": r = o.ListCount & " " & o.List(1) & " " & o.List(2)', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.AddItem "b", 1: r = o.List(1)', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.RemoveItem 1: r = o.ListCount', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.RemoveAllItems: r = o.ListCount', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.AddItem "b": o.ListIndex = 2: r = o.Value & " " & o.ListIndex & " " & o.Selected(2)', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.AddItem "a": o.AddItem "b": o.Selected(1) = True: r = o.ListIndex', 'Set o = ws.ListBoxes.Add(48, 96, 72, 24): o.List = Array("x", "y"): r = o.ListCount & " " & o.List(2)',
          'Set o = ws.DropDowns.Add(48, 96, 72, 24): o.AddItem "a": o.AddItem "b": o.Value = 2: r = o.ListIndex & " " & o.Text', 'Set o = ws.DropDowns.Add(48, 96, 72, 24): ws.Range("C1:C3").Value = 5: o.ListFillRange = "C1:C3": r = o.ListCount & " " & o.List(1)', 'Set o = ws.DropDowns.Add(48, 96, 72, 24): o.LinkedCell = "B2": o.AddItem "a": o.ListIndex = 1: r = ws.Range("B2").Value', 'Set o = ws.CheckBoxes.Add(48, 96, 72, 24): o.LinkedCell = "B2": o.Value = xlOn: r = ws.Range("B2").Value & " " & o.Value', 'Set o = ws.CheckBoxes.Add(48, 96, 72, 24): o.LinkedCell = "B2": ws.Range("B2").Value = True: r = o.Value', 'Set o = ws.OptionButtons.Add(48, 96, 72, 24): Set c = ws.OptionButtons.Add(48, 144, 72, 24): o.Value = xlOn: c.Value = xlOn: r = o.Value & " " & c.Value', 'Set o = ws.OptionButtons.Add(48, 96, 72, 24): o.LinkedCell = "B2": o.Value = xlOn: r = ws.Range("B2").Value',
          'Set o = ws.Buttons.Add(48, 96, 72, 24): Set c = ws.CheckBoxes.Add(48, 144, 72, 24): r = ws.Buttons.Count & " " & ws.CheckBoxes.Count & " " & ws.Shapes.Count & " " & o.Index & " " & c.Index & " " & c.ZOrder', 'Set c = ws.CheckBoxes.Add(48, 144, 72, 24): Set o = ws.Buttons.Add(48, 96, 72, 24): r = o.Index & " " & c.Name & " " & o.Name', 'Set o = ws.Buttons.Add(48, 96, 72, 24): ws.Protect: r = o.Caption', 'Set o = ws.Buttons.Add(48, 96, 72, 24): ws.Protect: o.Caption = "x": r = o.Caption', 'Set o = ws.Buttons.Add(48, 96, 72, 24): ws.Protect: o.Left = 96: r = o.Left', 'ws.Protect: Set o = ws.Buttons.Add(48, 96, 72, 24): r = o.Name']


def _probe() -> str:
    lines: list[str] = []
    for first in range(0, len(CASES), 90):
        lines += [f'Private Sub Act{first // 90}(ByVal which As Long, ws As Worksheet, r As Variant)', '    Dim o As Object, c As Object, n As Long, t As String', '    Select Case which']
        for index, statement in enumerate(CASES[first:first + 90], first):
            lines += [f'    Case {index}', f'        {statement}']
        lines += ['    End Select', 'End Sub', '']
    lines += ['Public Function Probe(ByVal which As Long) As String', '    Dim wb As Workbook, ws As Worksheet, s As String, v As String, r As Variant', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Activate', '    On Error GoTo Failed', '    Select Case which \\ 90']
    for block in range((len(CASES) + 89) // 90):
        lines += [f'    Case {block}: Act{block} which, ws, r']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next', '    Err.Clear', '    v = "": v = CStr(r)', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v', '    On Error GoTo 0', '    wb.Close False', '    Application.DisplayAlerts = True', '    Probe = s', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    answers: list[str] = []
    excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
    excel.__enter__()
    try:
        excel.new_document()
        for index in range(len(CASES)):
            result = excel.run_vba(source, 'Probe', (index,), timeout=60.0)
            if result.ok and not result.dialogs:
                answers.append(str(result.value))
                continue
            answers.append(f'STOPPED {result.outcome}')
            # A case that stopped Excel, or left it at a dialog, is that case's answer; the next gets a new Excel.
            excel.__exit__(None, None, None)
            excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
            excel.__enter__()
            excel.new_document()
    finally:
        excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source, 'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}, indent=1) + '\n', encoding='utf-8')
    for statement, answer in zip(CASES, answers, strict=True):
        print(statement, '=>', answer)


if __name__ == '__main__':
    main()
