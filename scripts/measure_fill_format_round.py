"""A held-out round of FillFormat: the members the model has, on shapes and statements no earlier fill probe used, and a first
measurement of the texture members it lacks (TextureAlignment, TextureHorizontalScale, TextureVerticalScale, TextureOffsetX,
TextureOffsetY, TextureTile, TextureName, TextureType, PresetTexture, PictureEffects) on fills that hold no texture. The model's
answers are written down first, each case run alone, so a refusal answers that case only.

Each case makes a workbook, adds the case's shape as sh, runs its statement on f, its Fill, typed and late, and reads READS, each as its
value or "e" and the error number. It then saves the workbook, reopens it, reads READS again and records the shape's spPr.

    python scripts/measure_fill_format_round.py predict [round]
    python scripts/measure_fill_format_round.py measure [round]

writes tests/fixtures/fill_format_round_<round>.json, which tests/test_excel_fill_format_round.py replays.
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
SHAPES = {'oval': 'ws.Shapes.AddShape(msoShapeOval, 20, 20, 120, 60)', 'round': 'ws.Shapes.AddShape(msoShapeRoundedRectangle, 20, 20, 120, 60)',
          'text': 'ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 20, 120, 60)', 'line': 'ws.Shapes.AddLine(20, 20, 140, 80)'}
READS = ['f.Visible', 'f.Type', 'f.ForeColor.RGB', 'f.BackColor.RGB', 'f.Transparency', 'f.Pattern', 'f.GradientStyle', 'f.GradientVariant',
         'f.GradientColorType', 'f.GradientAngle', 'f.RotateWithObject', 'f.GradientStops.Count', 'f.TextureAlignment',
         'f.TextureHorizontalScale', 'f.TextureVerticalScale', 'f.TextureOffsetX', 'f.TextureOffsetY', 'f.TextureTile', 'f.TextureName',
         'f.TextureType', 'f.PresetTexture', 'f.PictureEffects.Count']
ROUNDS: dict[int, list[str]] = {
    1: ['', 'f.Solid', 'f.ForeColor.RGB = RGB(10, 200, 30)', 'f.Solid: f.BackColor.RGB = RGB(5, 6, 7)', 'f.Transparency = 0.25',
        'f.Visible = msoFalse', 'f.Visible = msoFalse: f.Visible = msoTrue', 'f.Patterned msoPatternDarkVertical',
        'f.TwoColorGradient msoGradientDiagonalUp, 2', 'f.OneColorGradient msoGradientHorizontal, 1, 0.3',
        'f.PresetGradient msoGradientVertical, 3, msoGradientFire', 'f.TwoColorGradient msoGradientHorizontal, 1: f.GradientAngle = 33',
        'f.TwoColorGradient msoGradientHorizontal, 1: f.RotateWithObject = msoFalse', 'f.TextureTile = msoTrue', 'f.TextureOffsetX = 5',
        'f.TextureHorizontalScale = 2', 'f.TextureAlignment = msoTextureCenter', 'f.Patterned 5: f.TextureTile = msoFalse'],
}
ROUND = 1
#: The reads a model run makes, where it reads some of READS only.
CHOSEN: list[str] = []


def out() -> Path:
    return ROOT / f'tests/fixtures/fill_format_round_{ROUND}.json'


def cases() -> list[list[Any]]:
    return [[shape, statement, typed] for typed in (False, True) for shape in SHAPES for statement in ROUNDS[ROUND]]


def _reads(prefix: str) -> list[str]:
    lines: list[str] = []
    for read in CHOSEN or READS:
        lines += ['    Err.Clear: v = Empty', f'    v = {read}', f'    s = s & "|{prefix}" & Show(v)']
    return lines


def _case(case: list[Any], index: int) -> str:
    shape, statement, _typed = case
    lines = ['    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', f'    Set sh = {SHAPES[shape]}', '    sh.Name = "S"',
             '    Set f = sh.Fill', '    Err.Clear', f'    {statement}', '    s = "e" & Err.Number', *_reads(''),
             f'    wb.SaveAs folder & Application.PathSeparator & "f{index}.xlsx", xlOpenXMLWorkbook', '    wb.Close SaveChanges:=False',
             f'    Set wb = Workbooks.Open(folder & Application.PathSeparator & "f{index}.xlsx")', '    Set f = wb.Worksheets(1).Shapes("S").Fill',
             *_reads('r'), '    Set f = Nothing: Set sh = Nothing', '    wb.Close SaveChanges:=False', f'    out = out & "{SEPARATOR}" & s']
    return '\n'.join(lines) + '\n'


HELPERS = '''Private Function Show(ByVal v As Variant) As String
    If Err.Number Then Show = "e" & Err.Number: Exit Function
    Show = TypeName(v) & ":" & CStr(v)
End Function
'''
PART = 8


def macro(chosen: list[tuple[int, list[Any]]]) -> str:
    typed = chosen[0][1][2]
    declare = ('    Dim wb As Workbook, ws As Worksheet, sh As Shape, f As FillFormat, v As Variant, s As String\n' if typed else
               '    Dim wb As Object, ws As Object, sh As Object, f As Object, v As Variant, s As String\n')
    parts = [chosen[at:at + PART] for at in range(0, len(chosen), PART)]
    procedures = ''.join(f'Private Sub Part{n}(ByRef out As String, ByVal folder As String)\n{declare}    On Error Resume Next\n'
                         f'{"".join(_case(case, index) for index, case in part)}End Sub\n' for n, part in enumerate(parts))
    calls = ''.join(f'    Part{n} out, folder\n' for n in range(len(parts)))
    return (f'{HELPERS}{procedures}Function Probe(ByVal folder As String) As String\n    Dim out As String\n    Application.DisplayAlerts = False\n'
            f'{calls}    Application.DisplayAlerts = True\n    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')


def shape_xml(folder: Path, index: int) -> str:
    with zipfile.ZipFile(folder / f'f{index}.xlsx') as package:
        drawing = package.read('xl/drawings/drawing1.xml').decode('utf-8')
    found = re.findall(r'<xdr:spPr>.*?</xdr:spPr>|<xdr:spPr/>', drawing, re.S)
    return found[0] if found else ''


def _run(case: list[Any], index: int, folder: Path, reads: list[str]) -> list[str] | None:
    """The model's answer to the case with these reads, its segments split, or None where it refuses."""
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
        CHOSEN = []


def alone(case: list[Any], index: int, folder: Path) -> list[str]:
    """The model's answer to one case run by itself, and the shape's saved spPr. A read the model refuses answers "refused" in its own
    place, so the others still answer: the reads go together where none is refused, else one at a time."""
    whole = _run(case, index, folder, READS)
    if whole is not None:
        return ['|'.join(whole), shape_xml(folder, index) if whole != ['does not compile'] else '']
    first = _run(case, index, folder, READS[:1])
    if first is None:
        return ['refused', '']
    xml = shape_xml(folder, index)
    before, after = [], []
    for read in READS:
        got = _run(case, index, folder, [read])
        before.append('refused' if got is None else got[1])
        after.append('refused' if got is None else got[2])
    return ['|'.join([first[0], *before, *after]), xml]


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
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel, tempfile.TemporaryDirectory() as folder:
        excel.new_document()
        for typed in (False, True):
            chosen = [(index, case) for index, case in enumerate(every) if case[2] is typed]
            for at in range(0, len(chosen), 24):
                result = excel.run_vba(macro(chosen[at:at + 24]), 'Probe', (folder,), timeout=600, line_numbers=False)
                assert result.ok, (at, result.error, [one.message for one in result.dialogs])
                for (index, _case_), value in zip(chosen[at:at + 24], str(result.value).split(SEPARATOR)):
                    answers.append([value, shape_xml(Path(folder), index)])
                print(typed, at + 24, flush=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    out().write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    wrong = sum(model != excel and model[0] != 'refused' for model, excel in zip(record['model'], answers))
    print('model wrong', wrong, 'refused', sum(model[0] == 'refused' for model in record['model']), 'of', len(answers))


if __name__ == '__main__':
    ROUND = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
