"""A note's box: Locked, Placement and its fill colour, on notes a file holds and on a new one. Each case opens tests/fixtures/object_roundtrip.xlsx, makes one
change, and reads the box: in Excel, then from what Excel saves when Excel opens it again; in the model, then from what the model saves when Excel opens that. Rival
rules for Placement in the VML: MoveWithCells and SizeWithCells each mean the box does not follow, so 3 is both, 2 only SizeWithCells and 1 neither; or the reverse.

    python scripts/measure_note_shapes.py

writes tests/fixtures/note_shapes.json, which tests/test_excel_note_shapes.py replays.
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

OUT = ROOT / 'tests/fixtures/note_shapes.json'
BOOK = ROOT / 'tests/fixtures/object_roundtrip.xlsx'
CHANGES = ['Locked = False', 'Locked = True', 'Placement = 1', 'Placement = 2', 'Placement = 3', 'Placement = 99', 'Fill.ForeColor.RGB = 255', 'Fill.ForeColor.RGB = 14811135']
TARGETS = ['A10', 'A11', 'A12', 'J30']
READER = '''Private Function One(s As Object, ByVal which As Long) As String
    Dim v As Variant
    On Error Resume Next
    Select Case which
    Case 0
        v = s.Locked
    Case 1
        v = s.Placement
    Case 2
        v = s.Fill.ForeColor.RGB
    End Select
    If Err.Number Then
        One = "E" & Err.Number
    Else
        One = CStr(v)
    End If
End Function

Public Function ReadBox(c As Object) As String
    ReadBox = One(c.Comment.Shape, 0) & "|" & One(c.Comment.Shape, 1) & "|" & One(c.Comment.Shape, 2)
End Function

Public Function ReadAt(ByVal path As String, ByVal target As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    ReadAt = ReadBox(wb.Worksheets(1).Range(target))
    wb.Close False
End Function

Private Sub Change(s As Object, ByVal which As Long)
    Select Case which
''' + ''.join(f'    Case {index}\n        s.{change}\n' for index, change in enumerate(CHANGES)) + '''    End Select
End Sub

Public Function Applied(ByVal wb As Object, ByVal target As String, ByVal which As Long) As String
    Dim c As Object, t As String
    Set c = wb.Worksheets(1).Range(target)
    If c.Comment Is Nothing Then c.AddComment "new"
    On Error Resume Next
    Change c.Comment.Shape, which
    t = "E" & Err.Number & "@"
    Err.Clear
    On Error GoTo 0
    Applied = t & ReadBox(c)
End Function

Public Function InModel(ByVal target As String, ByVal which As Long) As String
    InModel = Applied(ActiveWorkbook, target, which)
End Function
'''
DRIVER = '''
Public Function Tried(ByVal path As String, ByVal target As String, ByVal which As Long, ByVal out As String) As String
    Dim wb As Workbook
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, 0)
    Tried = Applied(wb, target, which)
    wb.SaveAs out, xlOpenXMLWorkbook
    wb.Close False
End Function
'''


def main() -> None:
    answers: list[dict[str, str]] = []
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as folder, ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        for target in TARGETS:
            for which, change in enumerate(CHANGES):
                own, mine = Path(folder) / f'own{target}{which}.xlsx', Path(folder) / f'model{target}{which}.xlsx'
                done = excel.run_vba(READER + DRIVER, 'Tried', (str(BOOK), target, which, str(own)), timeout=40.0, line_numbers=False)
                answer = {'target': target, 'change': change, 'excel': str(done.value) if done.ok else f'failed {done.outcome}', 'reopened': '', 'model_reopened': ''}
                if done.ok:
                    answer['reopened'] = str(excel.run_vba(READER + DRIVER, 'ReadAt', (str(own), target), timeout=40.0, line_numbers=False).value)
                app = ExcelApplication.open(BOOK, with_vba=False)
                app.add_module(READER, name='Probe')
                try:
                    app.run('Probe.InModel', target, which)
                    app.save(mine)
                except VBAUnsupportedError as refused:
                    answer['model_reopened'] = f'reports {str(refused)[:80]}'
                except VBARuntimeError as failed:
                    answer['model_reopened'] = f'failed {failed}'
                if not answer['model_reopened']:
                    read = excel.run_vba(READER + DRIVER, 'ReadAt', (str(mine), target), timeout=40.0, line_numbers=False)
                    answer['model_reopened'] = str(read.value) if read.ok else f'not opened {read.outcome} {[one.message[:100] for one in read.dialogs]}'
                answers.append(answer)
                mark = 'same' if answer['model_reopened'] == answer['reopened'] else 'DIFF'
                print(mark, target, change, answer['excel'], '|', answer['reopened'], '|', answer['model_reopened'][:60], flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'changes': CHANGES, 'source': READER, 'cases': answers}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
