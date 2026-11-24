"""Held-out rounds of a shape's ThreeDFormat (Shape.ThreeD): every member set and read, typed and late, on an AutoShape, an
oval, a text box and a line, and what each change writes to the shape's spPr (scene3d and sp3d). The model's answers are
written down first, each case run alone, so a refusal answers that case only; a read the model refuses answers "refused" in its
own place.

Each case makes a workbook, adds the case's shape as sh, runs its statement on t, its ThreeD, and reads READS, each as its
TypeName and value or "e" and the error number, then saves the workbook and records the shape's spPr.

    python scripts/measure_three_d_round.py predict [round]
    python scripts/measure_three_d_round.py measure [round]

writes tests/fixtures/three_d_round_<round>.json, which tests/test_excel_three_d_round.py replays.
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
SHAPES = {'rect': 'ws.Shapes.AddShape(msoShapeRectangle, 20, 20, 120, 60)', 'oval': 'ws.Shapes.AddShape(msoShapeOval, 20, 20, 120, 60)',
          'text': 'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 20, 120, 60)', 'line': 'ws.Shapes.AddLine(20, 20, 140, 80)'}
READS = ['TypeName(t)', 'TypeName(t.Parent)', 't.Visible', 't.BevelTopType', 't.BevelTopDepth', 't.BevelTopInset', 't.BevelBottomType',
         't.BevelBottomDepth', 't.BevelBottomInset', 't.Depth', 't.ContourWidth', 't.ContourColor.RGB', 't.ExtrusionColor.RGB',
         't.ExtrusionColorType', 't.FieldOfView', 't.LightAngle', 't.Perspective', 't.PresetCamera', 't.PresetExtrusionDirection',
         't.PresetLighting', 't.PresetLightingDirection', 't.PresetLightingSoftness', 't.PresetMaterial', 't.PresetThreeDFormat',
         't.ProjectText', 't.RotationX', 't.RotationY', 't.RotationZ', 't.Z', 'TypeName(t.ExtrusionColor)']
ROUNDS: dict[int, list[str]] = {
    # Round 1: each scalar member set alone, some out of range, and the rotation methods.
    1: ['', 't.Visible = msoTrue', 't.Visible = msoFalse', 't.BevelTopType = msoBevelCircle', 't.BevelTopDepth = 5', 't.BevelTopInset = 4',
        't.BevelBottomType = msoBevelAngle', 't.BevelBottomDepth = 3', 't.BevelBottomInset = 2', 't.Depth = 10', 't.ContourWidth = 2',
        't.ContourColor.RGB = RGB(255, 0, 0)', 't.ExtrusionColor.RGB = RGB(0, 0, 255)', 't.ExtrusionColorType = msoExtrusionColorCustom',
        't.FieldOfView = 30', 't.LightAngle = 20', 't.Perspective = msoTrue', 't.ProjectText = msoTrue', 't.RotationX = 30',
        't.RotationY = 20', 't.RotationZ = 10', 't.Z = 5', 't.IncrementRotationX 10', 't.IncrementRotationY 10', 't.IncrementRotationZ 10',
        't.IncrementRotationHorizontal 10', 't.IncrementRotationVertical 10', 't.RotationX = 30: t.ResetRotation', 't.Depth = -1',
        't.BevelTopDepth = 2000', 't.RotationX = 400', 't.FieldOfView = 200', 't.Depth = 10: t.Visible = msoFalse',
        't.BevelTopType = msoBevelCircle: t.BevelTopDepth = 5: t.BevelTopInset = 4'],
}
ROUND = 1


def out() -> Path:
    return ROOT / f'tests/fixtures/three_d_round_{ROUND}.json'


def cases() -> list[list[Any]]:
    return [[shape, statement, typed] for typed in (False, True) for shape in SHAPES for statement in ROUNDS[ROUND]]


#: The reads a model run makes, where it reads one of READS alone; None for all of them.
CHOSEN: list[str] | None = None


def _case(case: list[Any], index: int) -> str:
    shape, statement, _typed = case
    lines = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    Set sh = {SHAPES[shape]}', '    sh.Name = "S"',
             '    Set t = sh.ThreeD', '    Err.Clear', f'    {statement}', '    s = "e" & Err.Number']
    for read in READS if CHOSEN is None else CHOSEN:
        lines += ['    Err.Clear: v = Empty', f'    v = {read}', '    s = s & "|" & Show(v)']
    lines += [f'    wb.SaveAs folder & Application.PathSeparator & "t{index}.xlsx", xlOpenXMLWorkbook', '    wb.Close SaveChanges:=False',
              '    Set t = Nothing: Set sh = Nothing', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & CStr(v)
End Function
'''
PART = 8


def macro(chosen: list[tuple[int, list[Any]]]) -> str:
    typed = chosen[0][1][2]
    declare = ('    Dim wb As Workbook, ws As Worksheet, sh As Shape, t As ThreeDFormat, v As Variant, s As String\n' if typed else
               '    Dim wb As Object, ws As Object, sh As Object, t As Object, v As Variant, s As String\n')
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{n}(ByRef out As String, ByVal folder As String)\n{declare}    On Error Resume Next\n'
                         f'{"".join(_case(case, index) for index, case in part)}End Sub\n' for n, part in enumerate(parts))
    calls = ''.join(f'    Part{n} out, folder\n' for n in range(len(parts)))
    return (f'{HELPERS}{procedures}Function Probe(ByVal folder As String) As String\n    Dim out As String\n'
            '    Application.DisplayAlerts = False\n'
            f'{calls}    Application.DisplayAlerts = True\n    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')


def shape_xml(folder: Path, index: int) -> str:
    """The saved shape's spPr, or "" where the book or its drawing is not there."""
    try:
        with zipfile.ZipFile(folder / f't{index}.xlsx') as package:
            drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    except (FileNotFoundError, KeyError):
        return ''
    found = re.findall(r'<xdr:spPr>.*?</xdr:spPr>|<xdr:spPr/>', drawing, re.S)
    return found[0] if found else ''


def _run(case: list[Any], index: int, folder: Path, reads: list[str] | None) -> list[str] | None:
    """The model's answer to the case with these reads, split, or None where it refuses."""
    global CHOSEN
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

    CHOSEN = reads
    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    try:
        app.add_module(macro([(index, case)]), name='H')
        return str(app.run('H.Probe', str(folder))).split('|')
    except VBACompileError:
        return ['does not compile']
    except VBAUnsupportedError:
        return None
    finally:
        CHOSEN = None


def alone(case: list[Any], index: int, folder: Path) -> list[str]:
    """The model's answer to one case run by itself, and the shape's saved spPr; where a read is refused, each read alone."""
    whole = _run(case, index, folder, None)
    if whole is not None:
        return ['|'.join(whole), shape_xml(folder, index) if whole != ['does not compile'] else '']
    first = _run(case, index, folder, READS[:1])
    if first is None:
        return ['refused', '']
    xml = shape_xml(folder, index)
    reads = []
    for read in READS:
        got = _run(case, index, folder, [read])
        reads.append('refused' if got is None else got[1])
    return ['|'.join([first[0], *reads]), xml]


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    every = cases()
    with tempfile.TemporaryDirectory() as folder:
        model = [alone(case, index, Path(folder)) for index, case in enumerate(every)]
    text = json.dumps({'reads': READS, 'cases': every, 'model': model}, indent=1) + '\n'
    out().write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(out().read_text(encoding='utf-8'))
    every: list[list[Any]] = record['cases']
    answers: list[list[str]] = []
    height = 'Function H() As String\n    H = CStr(ActiveSheet.StandardHeight)\nEnd Function\n'
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel, tempfile.TemporaryDirectory() as folder:
        excel.new_document()
        heights = [str(excel.run_vba(height, 'H', line_numbers=False).value)]
        for typed in (False, True):
            chosen = [(index, case) for index, case in enumerate(every) if case[2] is typed]
            for at in range(0, len(chosen), 24):
                result = excel.run_vba(macro(chosen[at:at + 24]), 'Probe', (folder,), timeout=600, line_numbers=False)
                assert result.ok, (at, result.error, [one.message for one in result.dialogs])
                for (index, _case_), value in zip(chosen[at:at + 24], str(result.value).split(SEPARATOR)):
                    answers.append([value, shape_xml(Path(folder), index)])
                print(typed, at + 24, flush=True)
        heights.append(str(excel.run_vba(height, 'H', line_numbers=False).value))
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': heights, 'excel': answers})
    out().write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model[0] != 'refused' for model, excel in zip(record['model'], answers))
    print('StandardHeight', heights, 'model wrong', wrong, 'refused', sum(model[0] == 'refused' for model in record['model']), 'of', len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
