"""Fill and Line set on shapes a file holds. Each case opens tests/fixtures/object_roundtrip.xlsx, makes one change, and reads the shape: in
Excel, then from what Excel saves when Excel opens it again; in the model, then from what the model saves when Excel opens that. So `model_reopened`
against `reopened` is Excel's own reading of the model's file.

    python scripts/measure_loaded_fill_line.py

writes tests/fixtures/loaded_fill_line.json, which tests/test_excel_loaded_fill_line.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
from pyopenvba.apps.excel import ExcelApplication  # noqa: E402
from pyopenvba.exceptions import VBARuntimeError, VBAUnsupportedError  # noqa: E402

OUT = ROOT / 'tests/fixtures/loaded_fill_line.json'
BOOK = ROOT / 'tests/fixtures/object_roundtrip.xlsx'
READ = ['Fill.Visible', 'Fill.Type', 'Fill.ForeColor.RGB', 'Fill.ForeColor.ObjectThemeColor', 'Fill.ForeColor.Type', 'Fill.Transparency', 'Fill.Pattern', 'Line.Visible', 'Line.ForeColor.RGB', 'Line.ForeColor.ObjectThemeColor',
        'Line.Weight', 'Line.DashStyle', 'Line.Transparency', 'Line.BeginArrowheadStyle', 'Line.EndArrowheadStyle', 'Shadow.Visible']
CHANGES = ['Fill.ForeColor.RGB = 255', 'Fill.ForeColor.ObjectThemeColor = 6', 'Fill.Visible = msoFalse', 'Fill.Visible = msoTrue', 'Fill.Transparency = 0.25', 'Fill.Solid', 'Fill.Patterned 5',
           'Line.ForeColor.RGB = 65280', 'Line.ForeColor.ObjectThemeColor = 7', 'Line.Visible = msoFalse', 'Line.Visible = msoTrue', 'Line.Weight = 4', 'Line.DashStyle = 4', 'Line.Transparency = 0.5', 'Line.EndArrowheadStyle = 2']
TARGETS = ['Hidden', 'Painted', 'Box', 'Joined', 'Pic']
CASES = [(target, change) for target in TARGETS for change in CHANGES]

READER = ['Private Function One(s As Object, ByVal which As Long) As String', '    Dim v As Variant', '    On Error Resume Next', '    Select Case which']
READER += [f'    Case {index}\n        v = s.{read}' for index, read in enumerate(READ)]
READER += ['    End Select', '    If Err.Number Then', '        One = "E" & Err.Number', '    Else', '        One = CStr(v)', '    End If', 'End Function', '',
           'Public Function ReadShape(s As Object) As String', '    Dim i As Long, t As String', f'    For i = 0 To {len(READ) - 1}', '        t = t & "|" & One(s, i)', '    Next', '    ReadShape = Mid$(t, 2)', 'End Function', '',
           'Public Function ReadAt(ByVal path As String, ByVal target As String) As String', '    Dim wb As Workbook', '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Open(path, 0)', '    ReadAt = ReadShape(wb.Worksheets(1).Shapes(target))', '    wb.Close False', 'End Function', '']
APPLY = ['Private Sub Change(s As Object, ByVal which As Long)', '    Select Case which']
APPLY += [f'    Case {index}\n        s.{change}' for index, change in enumerate(CHANGES)]
APPLY += ['    End Select', 'End Sub', '']
DRIVER = '''Public Function Tried(ByVal path As String, ByVal target As String, ByVal which As Long, ByVal out As String) As String
    Dim wb As Workbook, t As String, s As Object
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Set s = wb.Worksheets(1).Shapes(target)
    On Error Resume Next
    Change s, which
    t = "E" & Err.Number & "@"
    Err.Clear
    On Error GoTo 0
    t = t & ReadShape(s)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
    Tried = t
End Function
'''
SOURCE = '\n'.join(READER + APPLY) + '\n' + DRIVER
MODEL = '''Public Function InModel(ByVal target As String, ByVal which As Long) As String
    Dim t As String, s As Object
    Set s = ActiveWorkbook.Worksheets(1).Shapes(target)
    On Error Resume Next
    Change s, which
    t = "E" & Err.Number & "@"
    Err.Clear
    On Error GoTo 0
    InModel = t & ReadShape(s)
End Function
'''


def main() -> None:
    answers: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for index, (target, change) in enumerate(CASES):
            which = CHANGES.index(change)
            own = Path(folder) / f'own{index}.xlsx'
            done = excel.run_vba(SOURCE, 'Tried', (str(BOOK), target, which, str(own)), timeout=40.0, line_numbers=False)
            answer = {'target': target, 'change': change, 'excel': str(done.value) if done.ok else f'failed {done.outcome}', 'reopened': '', 'model_reopened': ''}
            if done.ok:
                answer['reopened'] = str(excel.run_vba(SOURCE, 'ReadAt', (str(own), target), timeout=40.0, line_numbers=False).value)
            mine = Path(folder) / f'model{index}.xlsx'
            app = ExcelApplication.open(BOOK, with_vba=False)
            app.add_module(SOURCE.split('Public Function Tried')[0] + MODEL, name='Probe')
            try:
                app.run('Probe.InModel', target, which)
                app.save(mine)
            except VBAUnsupportedError as refused:
                answer['model_reopened'] = f'reports {str(refused)[:80]}'
            except VBARuntimeError as failed:
                answer['model_reopened'] = f'failed {failed}'
            if not answer['model_reopened']:
                read = excel.run_vba(SOURCE, 'ReadAt', (str(mine), target), timeout=40.0, line_numbers=False)
                answer['model_reopened'] = str(read.value) if read.ok else f'not opened {read.outcome} {[one.message[:100] for one in read.dialogs]}'
            answers.append(answer)
            mark = 'same' if answer['model_reopened'] == answer['reopened'] else 'DIFF'
            print(mark, target, change, answer['excel'][:60], '|', answer['reopened'][:50], '|', answer['model_reopened'][:50], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READ, 'changes': CHANGES, 'source': SOURCE.split('Public Function Tried')[0] + MODEL, 'cases': answers}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
