"""Small Range reads on plain cells, a dynamic spill and an array formula: HasRichDataType, LinkedDataTypeState,
SavedAsArray, Summary, MDX, CommentThreaded, QueryTable and ServerActions.

    python scripts/measure_range_small_reads.py

writes tests/fixtures/range_small_reads.json, which tests/test_excel_range_small_reads.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/range_small_reads.json'
READS = ['HasRichDataType', 'LinkedDataTypeState', 'SavedAsArray', 'Summary', 'MDX', 'CommentThreaded', 'QueryTable',
         'ServerActions']
#: A number in A1, an empty B1, A1:B2, a spill from C1 and an array formula over D1:D2.
TARGETS = ['A1', 'B1', 'A1:B2', 'C1', 'C2', 'D1', 'D2', 'F1:F3']


def _probe() -> str:
    lines = ['Private Function GV(ByVal o As Object, ByVal n As String) As String', '    Dim v As Variant',
             '    On Error GoTo Failed', '    v = CallByName(o, n, VbGet)', '    If IsArray(v) Then',
             '        GV = "arr"', '    ElseIf IsNull(v) Then', '        GV = "Null"', '    Else',
             '        GV = TypeName(v) & ":" & CStr(v)', '    End If', '    Exit Function', 'Failed:',
             '    GV = "E" & Err.Number', 'End Function', '',
             'Private Function GO(ByVal o As Object, ByVal n As String) As String', '    Dim x As Object',
             '    On Error GoTo Failed', '    Set x = CallByName(o, n, VbGet)',
             '    If x Is Nothing Then GO = "Nothing" Else GO = "obj:" & TypeName(x)', '    Exit Function', 'Failed:',
             '    GO = "E" & Err.Number', 'End Function', '',
             'Private Function G(ByVal o As Object, ByVal n As String) As String', '    G = GV(o, n)',
             '    If G = "E450" Or G = "E438" Or G = "E13" Or G = "E91" Then G = G & "/" & GO(o, n)', 'End Function',
             '', 'Public Function Probe() As String', '    Dim wb As Workbook, ws As Worksheet, s As String, n As Variant',
             '    Dim t As Variant', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1").Value = 5', '    ws.Range("C1").Formula2 = "=SEQUENCE(2)"',
             '    ws.Range("D1:D2").FormulaArray = "=ROW(1:2)"']
    listed = ', '.join(f'"{name}"' for name in READS)
    targets = ', '.join(f'"{target}"' for target in TARGETS)
    lines += [f'    For Each t In Array({targets})', f'        For Each n In Array({listed})',
              '            s = s & "~" & G(ws.Range(t), CStr(n))', '        Next', '    Next', '    wb.Close False',
              '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('~')
    pairs = [[target, name] for target in TARGETS for name in READS]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'reads': [[*pair, answer] for pair, answer in zip(pairs, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for target, name, answer in record['reads']:
        print(target, name, answer)


if __name__ == '__main__':
    main()
