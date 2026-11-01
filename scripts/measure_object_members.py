"""Small members of pivot, shape, pane and conditional-format objects: Application, Parent, PTCondition and the
pivot objects' default members; with Validation.IMEMode, Font.Background and Interior.InvertIfNegative on cells, and
the dataValidation a save writes for IMEMode.

    python scripts/measure_object_members.py

writes tests/fixtures/object_members.json, which tests/test_excel_object_members.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/object_members.json'
#: Each read, on a sheet with a pivot table pt over A1:B5, a rectangle shape, a split window, conditional formats
#: on D1:D5 and validation on F1.
READS = [
    'TypeName(pt.Application) & " " & TypeName(pt.PivotFields(1).Application) & " " & '
    'TypeName(pt.PivotFields(1).PivotItems(1).Application) & " " & TypeName(pt.PivotCache.Application)',
    'pt.Value & " " & pt & " " & TypeName(pt.Value)',
    'pt.PivotFields(1).Value & " " & pt.PivotFields(1) & " " & pt.PivotFields(1).PivotItems(1)',
    'TypeName(ws.Shapes(1).Application) & " " & TypeName(ws.Shapes(1).Parent) & " " & '
    'TypeName(ws.Shapes.Application) & " " & TypeName(ws.Shapes.Parent) & " " & (ws.Shapes.Parent Is ws)',
    'TypeName(ActiveWindow.Panes.Application) & " " & TypeName(ActiveWindow.Panes.Parent) & " " & '
    'ActiveWindow.Panes.Count',
    'TypeName(fc.Application) & " " & fc.PTCondition',
    'TypeName(t10.Application) & " " & t10.PTCondition & " " & t10.CalcFor',
    'TypeName(av.Application) & " " & av.PTCondition & " " & av.CalcFor',
    'TypeName(cs.Application) & " " & cs.PTCondition',
    'TypeName(db.Application) & " " & db.PTCondition',
    'TypeName(ic.Application) & " " & ic.PTCondition',
    'TypeName(ws.Range("D1:D5").FormatConditions.AddUniqueValues.Application) & " " & '
    'ws.Range("D1:D5").FormatConditions(7).PTCondition',
    'TypeName(ic.IconCriteria.Application)',
    'TypeName(ic.IconSet.Application)',
    'TypeName(ic.IconCriteria(1).Application)',
    'ws.Range("F1").Validation.IMEMode',
    'ws.Range("F1").Font.Background',
    'ws.Range("F1").Interior.InvertIfNegative',
    'TypeName(ws.Range("F1").Font.Background)',
]
#: Statements run after the reads, each with what it reads, and the file saved, if any.
EDITS = [
    (['ws.Range("F1").Validation.IMEMode = xlIMEModeOn'], 'ws.Range("F1").Validation.IMEMode', 'ime.xlsx'),
    (['ws.Range("F1").Validation.IMEMode = xlIMEModeHiragana'], 'ws.Range("F1").Validation.IMEMode', 'hiragana.xlsx'),
    (['ws.Range("F1").Validation.IMEMode = 11'], 'ws.Range("F1").Validation.IMEMode', ''),
    (['ws.Range("F1").Validation.IMEMode = -1'], 'ws.Range("F1").Validation.IMEMode', ''),
    (['ws.Range("F1").Validation.IMEMode = xlIMEModeOn', 'ws.Range("F1").Validation.IMEMode = xlIMEModeNoControl'],
     'ws.Range("F1").Validation.IMEMode', 'back.xlsx'),
    (['ws.Range("F1").Validation.Delete'], 'ws.Range("F1").Validation.IMEMode', ''),
    ([], 'ws.Range("G1").Validation.IMEMode', ''),
    (['ws.Range("F1").Font.Background = xlBackgroundOpaque'], 'TypeName(ws.Range("F1").Font.Background)', ''),
    (['ws.Range("F1").Interior.InvertIfNegative = True'], 'ws.Range("F1").Interior.InvertIfNegative', ''),
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, fc As Object, t10 As Object, av As Object',
             '    Dim cs As Object, db As Object, ic As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:B5").Value = Application.Transpose(Application.Transpose(Array(Array("k", "v"), '
             'Array("a", 1), Array("b", 2), Array("a", 3), Array("c", 4))))',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B5")).CreatePivotTable(ws.Range("I1"), "PT")',
             '    pt.PivotFields("k").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum',
             '    ws.Shapes.AddShape msoShapeRectangle, 10, 10, 50, 20',
             '    ws.Range("D1:D5").Value = Application.Transpose(Array(1, 2, 3, 4, 5))',
             '    Set fc = ws.Range("D1:D5").FormatConditions.Add(xlCellValue, xlGreater, "=2")',
             '    Set t10 = ws.Range("D1:D5").FormatConditions.AddTop10',
             '    Set av = ws.Range("D1:D5").FormatConditions.AddAboveAverage',
             '    Set cs = ws.Range("D1:D5").FormatConditions.AddColorScale(3)',
             '    Set db = ws.Range("D1:D5").FormatConditions.AddDatabar',
             '    Set ic = ws.Range("D1:D5").FormatConditions.AddIconSetCondition',
             '    ws.Range("F1").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, "1", "5"',
             '    On Error GoTo Failed', '    Select Case which']
    for index, read in enumerate(READS):
        lines += [f'    Case {index}', f'        Tried = {read}']
    for index, (statements, read, saved) in enumerate(EDITS, start=len(READS)):
        lines.append(f'    Case {index}')
        lines += [f'        {statement}' for statement in statements]
        lines.append(f'        Tried = {read}')
        if saved:
            lines.append(f'        wb.SaveAs folder & "\\{saved}", xlOpenXMLWorkbook')
    lines += ['    End Select', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim s As String, i As Long',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(READS) + len(EDITS) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=300.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        parts = {}
        for _, _, saved in EDITS:
            if saved:
                with zipfile.ZipFile(Path(folder) / saved) as package:
                    sheet = package.read('xl/worksheets/sheet1.xml').decode('utf-8')
                found = re.search(r'<dataValidation\b[^>]*>', sheet)
                parts[saved] = found.group(0) if found else ''
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'reads': [[read, answer] for read, answer in zip(READS, answers[:len(READS)], strict=True)],
              'edits': [[statements, read, saved, answer]
                        for (statements, read, saved), answer in zip(EDITS, answers[len(READS):], strict=True)],
              'parts': parts}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['reads']:
        print(read[:70], '->', answer)
    for statements, read, saved, answer in record['edits']:
        print(statements[-1:] or [read], '->', answer, '|', parts.get(saved, ''))


if __name__ == '__main__':
    main()
