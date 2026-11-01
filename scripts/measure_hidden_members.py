"""The hidden members _Evaluate and _CodeName, reached through bracketed names as a macro reaches them: what each
answers next to its visible twin, and the errors. A code name itself is left out: a workbook a macro has just made
reads an empty one until its project is compiled.

Each case makes a workbook with 4 in A1 and a chart on its sheet, runs the case's statement and answers what it set.

    python scripts/measure_hidden_members.py

writes tests/fixtures/hidden_members.json, which tests/test_excel_hidden_members.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/hidden_members.json'
CASES = [
    's = ws.[_Evaluate]("A1*2") & "," & ws.Evaluate("A1*2")',
    's = TypeName(ws.[_Evaluate]("A1"))',
    's = ws.[_Evaluate]("SUM(1,2)")',
    's = ws.[_CodeName] & ""',
    's = CStr(ws.[_CodeName]) & "|" & CStr(ws.[_Evaluate])',
    's = Application.[_Evaluate]("2*3") & "," & Application.Evaluate("2*3")',
    's = ch.[_Evaluate]("1+1")',
    's = ch.[_CodeName]',
    'ws.[_CodeName] = "Data": s = ws.CodeName',
    's = TypeName(ws.[_Evaluate]("1/0"))',
    's = TypeName(ws.[_CodeName])',
    's = TypeName(ws.[_Evaluate])',
    's = TypeName(Application.[_Evaluate]("A1")) & "," & Application.[_Evaluate]("SUM(1,2)")',
    's = TypeName(ch.[_CodeName])',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long) As String',
             '    Dim wb As Workbook, ws As Object, ch As Object, s As String',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Range("A1").Value = 4',
             '    Set ch = ws.ChartObjects.Add(100, 10, 200, 100).Chart',
             '    On Error GoTo Failed', '    Select Case which']
    for index, case in enumerate(CASES):
        lines += [f'    Case {index}', f'        {case}']
    lines += ['    End Select', '    Tried = "ok " & s', 'Done:', '    wb.Close False', '    Exit Function', 'Failed:',
              '    Tried = "Err " & Err.Number', '    Resume Done', 'End Function', '',
              'Public Function Probe() As String', '    Dim i As Long, s As String',
              f'    For i = 0 To {len(CASES) - 1}', '        s = s & "@@@" & Tried(i)', '    Next',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer in record['cases']:
        print(case, '->', answer)


if __name__ == '__main__':
    main()
