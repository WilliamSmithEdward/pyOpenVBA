"""Rounds of a freeform's ShapeNodes (Shape.Nodes): Insert, Delete, SetPosition, SetEditingType and SetSegmentType, typed and
late, on an open polyline, a closed one, a curve and a freeform built of both; what each leaves of the nodes and the shape's box,
and the path the saved shape has. The model's answers are written down first, each case run alone, so a refusal answers that
case only.

Each case makes a workbook, adds the case's shape as sh, runs its statement on n, its Nodes, and reads NODES (every node's
SegmentType, EditingType and point) and READS, each as its TypeName and value or "e" and the error number, then saves the
workbook and records the shape's custGeom.

    python scripts/measure_shape_nodes_round.py predict [round]
    python scripts/measure_shape_nodes_round.py measure [round]

writes tests/fixtures/shape_nodes_round_<round>.json, which tests/test_excel_shape_nodes_round.py replays.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SEPARATOR = '%%'
#: Between what a case sets, before its book is saved and opened again, and what it does after.
AGAIN = ' || '
SHAPES = {'open': 'ws.Shapes.AddPolyline(Pts(10, 10, 110, 10, 60, 80))',
          'closed': 'ws.Shapes.AddPolyline(Pts(10, 10, 110, 10, 60, 80, 10, 10))',
          'curve': 'ws.Shapes.AddCurve(Pts(10, 50, 40, 100, 80, 100, 110, 50, 140, 0, 170, 60, 200, 50))',
          'mixed': 'MixedShape(ws)',
          'two': 'ws.Shapes.AddPolyline(Pts(10, 10, 110, 60))',
          'four': 'ws.Shapes.AddPolyline(Pts(10, 10, 110, 10, 110, 80, 10, 80))',
          'five': 'ws.Shapes.AddPolyline(Pts(20, 10, 120, 30, 90, 90, 40, 70, 10, 40))',
          'closed5': 'ws.Shapes.AddPolyline(Pts(20, 10, 120, 30, 90, 90, 40, 70, 10, 40, 20, 10))',
          'curve1': 'ws.Shapes.AddCurve(Pts(10, 50, 40, 100, 80, 100, 110, 50))',
          'curve3': 'ws.Shapes.AddCurve(Pts(10, 50, 40, 100, 80, 100, 110, 50, 140, 0, 170, 60, 200, 50, 230, 40, 260, 40, 290, 50))',
          'curveclosed': 'ws.Shapes.AddCurve(Pts(10, 50, 40, 100, 80, 100, 110, 50, 140, 0, 170, 60, 10, 50))'}
READS = ['TypeName(n)', 'TypeName(n.Parent)', 'n.Count', 'sh.Left', 'sh.Top', 'sh.Width', 'sh.Height', 'sh.Type',
         'sh.AutoShapeType', 'ws.Shapes.Count']
_LINE = 'msoSegmentLine'
_CURVE = 'msoSegmentCurve'
ROUNDS: dict[int, list[str]] = {
    # Round 1: each method alone and its errors, and what it leaves.
    1: ['', 'n.Delete 1', 'n.Delete 2', 'n.Delete n.Count', 'n.Delete 0', 'n.Delete 99', 'n.Delete 2: n.Delete 2',
        f'n.Insert 1, {_LINE}, msoEditingCorner, 50, 50', f'n.Insert 2, {_LINE}, msoEditingAuto, 50, 50',
        f'n.Insert n.Count, {_LINE}, msoEditingCorner, 250, 120', f'n.Insert 0, {_LINE}, msoEditingCorner, 50, 50',
        f'n.Insert 99, {_LINE}, msoEditingCorner, 50, 50', f'n.Insert 1, {_CURVE}, msoEditingCorner, 20, 90, 40, 90, 50, 50',
        f'n.Insert 2, {_CURVE}, msoEditingAuto, 50, 50', f'n.Insert 1, {_CURVE}, msoEditingCorner, 50, 50',
        f'n.Insert 1, 9, msoEditingCorner, 50, 50', f'n.Insert 1, {_LINE}, 9, 50, 50', f'n.Insert 1, {_LINE}, msoEditingSmooth, 50, 50',
        'n.SetPosition 1, 30, 30', 'n.SetPosition 2, 200, 5', 'n.SetPosition n.Count, 5, 150', 'n.SetPosition 2, 60.5, 20.25',
        'n.SetPosition 1, -20, -20', 'n.SetPosition 0, 30, 30', 'n.SetPosition 99, 30, 30', 'n.SetPosition 2, 2000, 2000',
        'n.SetEditingType 1, msoEditingCorner', 'n.SetEditingType 2, msoEditingSmooth', 'n.SetEditingType 2, msoEditingSymmetric',
        'n.SetEditingType 2, msoEditingAuto', 'n.SetEditingType 3, msoEditingSmooth', 'n.SetEditingType 4, msoEditingSmooth',
        'n.SetEditingType 2, 9', 'n.SetEditingType 0, msoEditingCorner',
        f'n.SetSegmentType 1, {_CURVE}', f'n.SetSegmentType 2, {_CURVE}', f'n.SetSegmentType 3, {_CURVE}',
        f'n.SetSegmentType 2, {_LINE}', f'n.SetSegmentType 4, {_LINE}', 'n.SetSegmentType 2, 9', f'n.SetSegmentType 0, {_LINE}',
        'TypeName(n.Application)', 'For Each o In n: s = s & "+": Next', 'n.Item(2).Points',
        f'n.SetPosition 2, 70, 30{AGAIN}', f'n.Delete 2{AGAIN}', f'n.Insert 1, {_LINE}, msoEditingCorner, 50, 50{AGAIN}',
        f'{AGAIN}n.SetPosition 2, 70, 30', f'{AGAIN}n.Delete 2'],
}
#: Round 2: what a new freeform's nodes read by how many there are, and line freeforms edited more than once.
_GROW = 'n.SetPosition 2, 200, 5'
ROUNDS[2] = ['', 'n.SetPosition 1, 15, 15', 'n.SetPosition 2, 100, 20', 'n.SetPosition 2, 105, 70', _GROW, f'{_GROW}: n.SetPosition 2, 300, 5',
             f'{_GROW}: n.SetPosition 2, 100, 20', f'{_GROW}: n.SetPosition 1, 300, 300', 'n.SetPosition 1, 15, 15: n.SetPosition 2, 200, 5',
             'n.SetPosition 1, 5, 5', 'n.SetPosition 1, 5, 5: n.SetPosition 2, 5, 200', 'n.SetPosition 2, 10, 10',
             f'n.Insert 1, {_LINE}, msoEditingCorner, 50, 40', f'n.Insert 2, {_LINE}, msoEditingCorner, 50, 40',
             f'n.Insert n.Count, {_LINE}, msoEditingCorner, 50, 40', f'n.Insert n.Count, {_LINE}, msoEditingCorner, 5, 5',
             f'n.Insert n.Count - 1, {_LINE}, msoEditingCorner, 250, 120', f'n.Insert 1, {_LINE}, msoEditingCorner, 250, 120',
             f'n.Insert 1, {_LINE}, msoEditingCorner, 50, 40: n.Insert 1, {_LINE}, msoEditingCorner, 60, 45',
             'n.Delete 1: n.Delete 1', 'n.Delete 3', 'n.Delete 1: n.SetPosition 1, 200, 5', f'{_GROW}: n.Delete 1',
             f'n.SetSegmentType 1, {_CURVE}: n.SetSegmentType 1, {_LINE}', f'n.SetSegmentType 2, {_CURVE}: n.Delete 1',
             f'n.SetSegmentType 3, {_CURVE}', f'n.SetSegmentType n.Count, {_CURVE}', f'{_GROW}: n.SetSegmentType 1, {_CURVE}',
             'n.SetEditingType 1, msoEditingCorner', 'n.SetEditingType n.Count, msoEditingCorner', 'n.SetPosition 1, 15, 15: n.SetEditingType 2, msoEditingCorner',
             f'{_GROW}{AGAIN}n.SetPosition 2, 300, 5', f'{AGAIN}{_GROW}', 'n.Delete n.Count: n.Delete n.Count',
             'n.Delete 1: n.Delete 1: n.Delete 1']
#: The shapes of each round where not all of them.
ROUND_SHAPES: dict[int, list[str]] = {1: ['open', 'closed', 'curve', 'mixed'], 2: ['open', 'closed', 'two', 'four', 'five', 'closed5', 'curve1', 'curve3', 'curveclosed']}
ROUND = 1


def out() -> Path:
    return ROOT / f'tests/fixtures/shape_nodes_round_{ROUND}.json'


def cases() -> list[list[Any]]:
    return [[shape, statement, typed] for typed in (False, True) for shape in ROUND_SHAPES.get(ROUND, list(SHAPES))
            for statement in ROUNDS[ROUND]]


def _case(case: list[Any], index: int) -> str:
    shape, statement, _typed = case
    lines = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    Set sh = {SHAPES[shape]}',
             '    sh.Name = "S"', '    Set n = sh.Nodes']
    if AGAIN in statement:
        before, statement = statement.split(AGAIN)
        if before:
            lines.append(f'    {before}')
        lines += [f'    wb.SaveAs folder & Application.PathSeparator & "a{index}.xlsx", xlOpenXMLWorkbook',
                  '    wb.Close SaveChanges:=False', f'    Set wb = Workbooks.Open(folder & Application.PathSeparator & "a{index}.xlsx")',
                  '    Set ws = wb.Worksheets(1)', '    Set sh = ws.Shapes("S")', '    Set n = sh.Nodes']
    if statement.startswith(('TypeName(', 'n.Item(')):
        statement = f's = Show({statement})' if statement.startswith('TypeName(') else f'v = {statement}: s = CStr(UBound(v, 1))'
    lines += ['    Err.Clear: s = ""', f'    {statement}', '    s = "e" & Err.Number & "/" & s', '    s = s & "|" & Listed(n)']
    for read in READS:
        lines += ['    Err.Clear: v = Empty', f'    v = {read}', '    s = s & "|" & Show(v)']
    lines += [f'    wb.SaveAs folder & Application.PathSeparator & "t{index}.xlsx", xlOpenXMLWorkbook', '    wb.Close SaveChanges:=False',
              '    Set n = Nothing: Set sh = Nothing', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & CStr(v)
End Function

Private Function Pts(ParamArray xy() As Variant) As Variant
    Dim a() As Single, i As Long
    ReDim a(1 To (UBound(xy) + 1) \\ 2, 1 To 2)
    For i = 0 To UBound(xy) Step 2
        a(i \\ 2 + 1, 1) = xy(i): a(i \\ 2 + 1, 2) = xy(i + 1)
    Next
    Pts = a
End Function

Private Function MixedShape(ByVal ws As Object) As Object
    Dim b As Object
    Set b = ws.Shapes.BuildFreeform(msoEditingCorner, 10, 10)
    b.AddNodes msoSegmentLine, msoEditingAuto, 110, 10
    b.AddNodes msoSegmentCurve, msoEditingCorner, 140, 40, 140, 80, 110, 110
    b.AddNodes msoSegmentLine, msoEditingAuto, 10, 110
    Set MixedShape = b.ConvertToShape
End Function

Private Function Listed(ByVal n As Object) As String
    Dim i As Long, c As Long, s As String, p As Variant, t As String
    On Error Resume Next
    Err.Clear
    c = n.Count
    If Err.Number Then Listed = "e" & Err.Number: Exit Function
    For i = 1 To c
        Err.Clear: t = CStr(n.Item(i).SegmentType)
        If Err.Number Then t = "e" & Err.Number
        s = s & ";" & t
        Err.Clear: t = CStr(n.Item(i).EditingType)
        If Err.Number Then t = "e" & Err.Number
        s = s & "/" & t
        Err.Clear: p = n.Item(i).Points
        If Err.Number Then
            s = s & "/e" & Err.Number
        Else
            s = s & "/" & CStr(p(1, 1)) & "," & CStr(p(1, 2))
        End If
    Next
    Listed = Mid(s, 2)
End Function
'''
PART = 8


#: The read of each node's EditingType, which a model run can leave out.
EDITING_READ = 'Err.Clear: t = CStr(n.Item(i).EditingType)'
#: What a node's EditingType reads where a model run left it out.
LEFT_OUT = 'refused'


def macro(chosen: list[tuple[int, list[Any]]], editing: bool = True) -> str:
    typed = chosen[0][1][2]
    helpers = HELPERS if editing else HELPERS.replace(EDITING_READ, f'Err.Clear: t = "{LEFT_OUT}"')
    declare = ('    Dim wb As Workbook, ws As Worksheet, sh As Shape, n As ShapeNodes, o As Object, v As Variant, s As String\n' if typed else
               '    Dim wb As Object, ws As Object, sh As Object, n As Object, o As Object, v As Variant, s As String\n')
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{k}(ByRef out As String, ByVal folder As String)\n{declare}    On Error Resume Next\n'
                         f'{"".join(_case(case, index) for index, case in part)}End Sub\n' for k, part in enumerate(parts))
    calls = ''.join(f'    Part{k} out, folder\n' for k in range(len(parts)))
    return (f'{helpers}{procedures}Function Probe(ByVal folder As String) As String\n    Dim out As String\n'
            '    Application.DisplayAlerts = False\n'
            f'{calls}    Application.DisplayAlerts = True\n    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')


def shape_xml(folder: Path, index: int) -> str:
    """The saved shape's xfrm and custGeom, or "" where the book or its drawing is not there."""
    try:
        with zipfile.ZipFile(folder / f't{index}.xlsx') as package:
            drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    except (FileNotFoundError, KeyError):
        return ''
    found = re.findall(r'<a:xfrm\b.*?</a:xfrm>|<a:custGeom>.*?</a:custGeom>', drawing, re.S)
    return ''.join(found[:2])


def _run(case: list[Any], index: int, folder: Path, editing: bool = True) -> list[str] | None:
    """The model's answer to the case, split, or None where it refuses."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    try:
        app.add_module(macro([(index, case)], editing), name='H')
        return str(app.run('H.Probe', str(folder))).split('|')
    except VBACompileError:
        return ['does not compile']
    except VBAUnsupportedError:
        return None


def alone(case: list[Any], index: int, folder: Path) -> list[str]:
    """The model's answer to one case run by itself, and the shape's saved markup; "refused" where it refuses."""
    whole = _run(case, index, folder)
    if whole is None:
        # Refused for its nodes' EditingType, the rest of it can still answer.
        whole = _run(case, index, folder, editing=False)
    if whole is None:
        return ['refused', '']
    return ['|'.join(whole), shape_xml(folder, index) if whole != ['does not compile'] else '']


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    every = cases()
    with tempfile.TemporaryDirectory() as folder:
        model = [alone(case, index, Path(folder)) for index, case in enumerate(every)]
    text = json.dumps({'reads': READS, 'cases': every, 'model': model}, indent=1) + '\n'
    out().write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


#: What a case Excel cannot run records in place of an answer.
BREAKS_EXCEL = 'not measured: breaks Excel'


def breaks_excel(case: list[Any]) -> bool:
    """Typed, SetEditingType on a curve's control point after a SetPosition raises 0x80020009 out of Excel and ends the session
    it runs in (round 2); late bound it answers."""
    return bool(case[2] and case[0].startswith('curve') and 'SetEditingType 2' in case[1] and 'SetPosition' in case[1])


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(out().read_text(encoding='utf-8'))
    every: list[list[Any]] = record['cases']
    answers: list[list[str]] = []
    height = 'Function H() As String\n    H = CStr(ActiveSheet.StandardHeight)\nEnd Function\n'
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel, tempfile.TemporaryDirectory() as folder:
        excel.new_document()
        heights = [str(excel.run_vba(height, 'H', line_numbers=False).value)]
        found: dict[int, list[str]] = {}
        for typed in (False, True):
            chosen = [(index, case) for index, case in enumerate(every) if case[2] is typed and not breaks_excel(case)]
            for at in range(0, len(chosen), 24):
                batch = chosen[at:at + 24]
                result = excel.run_vba(macro(batch), 'Probe', (folder,), timeout=600, line_numbers=False)
                assert result.ok, (typed, at, result.outcome, result.message)
                for (index, _case_), value in zip(batch, str(result.value).split(SEPARATOR)):
                    found[index] = [value, shape_xml(Path(folder), index)]
                print(typed, at + 24, flush=True)
        answers = [found.get(index, [BREAKS_EXCEL, '']) for index in range(len(every))]
        heights.append(str(excel.run_vba(height, 'H', line_numbers=False).value))
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': heights, 'excel': answers})
    out().write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model[0] != 'refused' for model, excel in zip(record['model'], answers))
    print('StandardHeight', heights, 'model wrong', wrong, 'refused', sum(model[0] == 'refused' for model in record['model']), 'of',
          len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
