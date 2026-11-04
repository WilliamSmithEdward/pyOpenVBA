"""The type each property answers on the objects macros reach most, typed and late. A fixture that keeps text sees 11 and 11 where one is a Long and the other a Double.

Rival rules: late answers as typed; or late is a Double where typed is a Long, as DropDown.ListCount and ControlFormat.Value are. Each read is one row: its type and value in
Excel, then in the model. Mail members and printing are not on these classes. Geometry goes by the display, so a run with StandardHeight other than 15 stops.

    python scripts/measure_late_types.py Font Shape ControlFormat TextFrame

writes tests/fixtures/typed_answers/late_types_<first class>.json, which tests/test_excel_late_types.py replays.
"""
from __future__ import annotations

import csv
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBAError  # noqa: E402

SETUP = r'''Public Sub Setup(ws As Object)
    Dim s As Object
    ws.Range("A1").Value = "text": ws.Range("A2").Value = 2: ws.Range("A3").Value = 3
    Set s = ws.Shapes.AddShape(1, 10, 10, 100, 40): s.Name = "Box": s.TextFrame.Characters.Text = "Boxed"
    Set s = ws.Shapes.AddTextbox(1, 120, 10, 100, 40): s.Name = "Note": s.TextFrame.Characters.Text = "Noted"
    Set s = ws.Shapes.AddFormControl(1, 10, 60, 90, 20): s.Name = "Tick": s.ControlFormat.Value = 1: s.ControlFormat.LinkedCell = "$C$1"
    Set s = ws.Shapes.AddFormControl(2, 10, 90, 90, 20): s.Name = "Drop": s.ControlFormat.ListFillRange = "$A$1:$A$3": s.ControlFormat.Value = 2
    Set s = ws.Shapes.AddFormControl(6, 10, 120, 90, 50): s.Name = "Pick": s.ControlFormat.ListFillRange = "$A$1:$A$3": s.ControlFormat.Value = 1
    Set s = ws.Shapes.AddFormControl(9, 110, 60, 20, 50): s.Name = "Spin": s.ControlFormat.Value = 3
End Sub
'''
#: Each class: the objects to read it on, as (label, expression on ws).
OBJECTS = {
    'Font': [('cell', 'ws.Range("A1").Font'), ('box', 'ws.Shapes("Box").TextFrame.Characters.Font')],
    'Shape': [('box', 'ws.Shapes("Box")'), ('note', 'ws.Shapes("Note")'), ('tick', 'ws.Shapes("Tick")')],
    'ControlFormat': [('tick', 'ws.Shapes("Tick").ControlFormat'), ('drop', 'ws.Shapes("Drop").ControlFormat'), ('pick', 'ws.Shapes("Pick").ControlFormat'), ('spin', 'ws.Shapes("Spin").ControlFormat')],
    'TextFrame': [('box', 'ws.Shapes("Box").TextFrame')],
    'FillFormat': [('box', 'ws.Shapes("Box").Fill')], 'LineFormat': [('box', 'ws.Shapes("Box").Line')], 'ColorFormat': [('fill', 'ws.Shapes("Box").Fill.ForeColor')],
    'ShadowFormat': [('box', 'ws.Shapes("Box").Shadow')], 'Characters': [('box', 'ws.Shapes("Box").TextFrame.Characters')], 'TextFrame2': [('box', 'ws.Shapes("Box").TextFrame2')],
    'TextRange2': [('box', 'ws.Shapes("Box").TextFrame2.TextRange')], 'Font2': [('box', 'ws.Shapes("Box").TextFrame2.TextRange.Font')],
}
OBJECTISH = re.compile(r'^(Application|Object|Range|Worksheet|Shape|Shapes|ShapeRange|Hyperlink|Chart|DrawingObject|Adjustments|[A-Z][A-Za-z0-9]*Format|[A-Z][A-Za-z0-9]*Frame2?|Font2?|Characters|TextRange2|Borders?|Interior|GroupShapes|Nodes|ShapeNodes|OLEFormat|ConnectorFormat|ColorFormat|GradientStops|ReflectionFormat|GlowFormat|SoftEdgeFormat|ThreeDFormat|PictureEffects|FreeformBuilder|SmartArt|Model3D|FileDialog|TickLabels)$')
SHOW = r'''Private Function Show(v As Variant) As String
    If IsObject(v) Then
        Show = "obj " & TypeName(v)
    ElseIf IsArray(v) Then
        Show = TypeName(v)
    ElseIf IsNull(v) Or IsEmpty(v) Then
        Show = TypeName(v)
    ElseIf IsError(v) Then
        Show = "Error " & CStr(v)
    Else
        Show = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def reads(wanted: list[str]) -> dict[str, list[str]]:
    found: dict[str, list[str]] = {name: [] for name in wanted}
    with (ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='') as file:
        for row in csv.DictReader(file):
            kind, signature = row['kind'], row['signature']
            if row['object'] not in found or not kind.startswith('property') or row['member'].startswith('_') or signature.startswith('('):
                continue
            found_type = re.search(r'\bAs (\w+)$', signature)
            returned = found_type.group(1) if found_type else ''
            if returned and not OBJECTISH.match(returned):
                found[row['object']].append(row['member'])
    return found


def module(wanted: list[str]) -> tuple[str, list[dict[str, object]]]:
    parts, cases = [SETUP, SHOW], []
    table = reads(wanted)
    for name in wanted:
        for label, get in OBJECTS[name]:
            for binding, declared in (('Typed', f'{"Office." if name in ("TextFrame2", "TextRange2", "Font2") else "Excel."}{name}'.replace('Office.TextFrame2', 'Excel.TextFrame2').replace('Office.TextRange2', 'Office.TextRange2').replace('Office.Font2', 'Office.Font2')), ('Late', 'Object')):
                title = f'{binding}{name}{label}'
                body = ''.join(f'    Case {at}\n        v = o.{member}\n' for at, member in enumerate(table[name]))
                parts.append(f'Public Function {title}(ws As Object, ByVal which As Long) As String\n    Dim o As {declared}, v As Variant\n    Set o = {get}\n    On Error Resume Next\n    Select Case which\n{body}'
                             f'    End Select\n    If Err.Number Then\n        {title} = "E" & Err.Number\n    Else\n        {title} = Show(v)\n    End If\nEnd Function\n\n')
                cases.extend({'class': name, 'on': label, 'binding': binding, 'member': member, 'routine': title, 'which': at} for at, member in enumerate(table[name]))
    return ''.join(parts), cases


DRIVER = r'''
Public Function Walk(ByVal list As String) As String
    Dim wb As Workbook, one As Variant, s As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Setup wb.Worksheets(1)
    s = wb.Worksheets(1).StandardHeight & "~~"
    For Each one In Split(list, "|")
        s = s & Application.Run("'" & ThisWorkbook.Name & "'!" & Split(one, ":")(0), wb.Worksheets(1), CLng(Split(one, ":")(1))) & "~~"
    Next
    wb.Close False
    Walk = s
End Function
'''
MODEL = r'''
Public Sub Prepared()
    Setup ActiveWorkbook.Worksheets(1)
End Sub

Public Function InModel(ByVal routine As String, ByVal which As Long) As String
    InModel = Application.Run(routine, ActiveWorkbook.Worksheets(1), which)
End Function
'''


def main() -> None:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')  # pyright: ignore[reportAttributeAccessIssue]
    wanted = sys.argv[1:]
    source, cases = module(wanted)
    groups: dict[str, list[dict[str, object]]] = {}
    for case in cases:
        groups.setdefault(str(case['routine']), []).append(case)
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    for routine, group in groups.items():
        list_text = '|'.join(f'{one["routine"]}:{one["which"]}' for one in group)
        done = session.run_vba(source + DRIVER, 'Walk', (list_text,), timeout=120.0, line_numbers=False)
        if not done.ok:
            print('NOT PASSED', routine, done.outcome, [one.message[:100] for one in done.dialogs], flush=True)
            session.close()
            session = ExcelSession(HarnessConfig(exclusive=False))
            session.new_document()
            for one in group:
                one['excel'] = f'failed {done.outcome}'
            continue
        answers = str(done.value).split('~~')
        assert answers[0] == '15', f'StandardHeight {answers[0]}: geometry would not be at 96 DPI'
        for one, answer in zip(group, answers[1:], strict=False):
            one['excel'] = answer
    session.close()
    app = ExcelApplication()
    app.add_workbook()
    app.add_module(source + MODEL, name='Probe')
    app.run('Probe.Prepared')
    for case in cases:
        try:
            case['model'] = str(app.run('Probe.InModel', case['routine'], case['which']))
        except VBAError as failed:
            case['model'] = f'reports {str(failed)[:50]}'
    (ROOT / 'tests/fixtures/typed_answers').mkdir(exist_ok=True)
    out = ROOT / f'tests/fixtures/typed_answers/late_types_{wanted[0]}.json'
    out.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source + MODEL, 'cases': cases}, indent=0) + '\n', encoding='utf-8')
    print(len(cases), 'cases')


if __name__ == '__main__':
    main()
