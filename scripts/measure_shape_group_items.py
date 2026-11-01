"""What a group's members answer and how a group moves, sizes and goes: ParentGroup, Regroup, ShapeRange.Ungroup,
GroupShapes.Parent and Range, Shapes(name) for a member, and a group's Left, Width and Delete.

Each case makes a workbook with three rectangles, A (100, 50, 120 by 60), B (250, 80, 60 by 40) and C (50, 200,
40 by 40), groups A and B, runs the case's statements and reads every top-level shape's name, id, type and box and
each group's members with their boxes, then saves it. The drawing is recorded with creation ids and anchor corners
cut out, as scripts/measure_shape_group.py records it.

Regroup after a member has moved is left out. Excel answers the members where they now are, but writes them at those
places inside the old group's child space, so the saved file draws them elsewhere. The model refuses it.

    python scripts/measure_shape_group_items.py

writes tests/fixtures/shape_group_items.json, which tests/test_excel_shape_group_items.py replays.
"""
from __future__ import annotations

import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from measure_shape_group import anchors
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_group_items.json'
CASES = [
    's = g.GroupItems("A").ParentGroup.Name',
    's = TypeName(ws.Shapes("C").ParentGroup)',
    's = TypeName(g.GroupItems.Parent) & "," & g.GroupItems.Parent.Name',
    's = TypeName(g.GroupItems.Range(Array("B"))) & "," & g.GroupItems.Range(2).Name',
    's = ws.Shapes("A").Name & "," & ws.Shapes("A").ParentGroup.Name',
    'g.Ungroup: s = ws.Shapes.Range(Array("A", "B")).Regroup.Name',
    's = TypeName(ws.Shapes.Range(Array("C")).Regroup)',
    's = TypeName(ws.Shapes.Range(Array("C")).Ungroup)',
    'g.Ungroup: Set g = ws.Shapes.Range(Array("A", "B", "C")).Group: g.GroupItems("A").Delete',
    'g.Width = 420: g.GroupItems("A").Left = 20',
    'ws.Shapes("A").Delete',
    's = TypeName(ws.Shapes.Range(Array("Group 4")).Ungroup)',
    'ws.Shapes.Range(Array("Group 4", "C")).Ungroup',
    'g.Left = 40',
    'g.Top = 20',
    'g.Width = 420',
    'g.Height = 35',
    'g.Delete',
    'g.GroupItems("A").Delete',
    'g.GroupItems("A").Left = 20',
    'g.GroupItems("B").Width = 100',
    'g.GroupItems("A").Name = "Z"',
    'g.Name = "G": s = ws.Shapes("G").ID',
]
READ = ('    For i = 1 To ws.Shapes.Count: s = s & "|" & ws.Shapes(i).Name & "," & ws.Shapes(i).ID & "," & '
        'ws.Shapes(i).Type & "," & ws.Shapes(i).Left & "," & ws.Shapes(i).Top & "," & ws.Shapes(i).Width & "," & '
        'ws.Shapes(i).Height: If ws.Shapes(i).Type = msoGroup Then For j = 1 To ws.Shapes(i).GroupItems.Count: '
        'Set m = ws.Shapes(i).GroupItems(j): s = s & ";" & m.Name & "," & m.Left & "," & m.Top & "," & m.Width & '
        '"," & m.Height: Next', '    Next')


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, g As Object, m As Object, s As String, i As Long, j As Long',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Shapes.AddShape(msoShapeRectangle, 100, 50, 120, 60).Name = "A"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 250, 80, 60, 40).Name = "B"',
             '    ws.Shapes.AddShape(msoShapeRectangle, 50, 200, 40, 40).Name = "C"',
             '    Set g = ws.Shapes.Range(Array("A", "B")).Group',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', *READ, '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        drawings = [anchors(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, drawings, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, _ in record['cases']:
        print(case[:50], '->', answer)


if __name__ == '__main__':
    main()
