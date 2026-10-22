"""Application.ReferenceStyle: what it reads, what it takes, what it changes and how a workbook keeps it.

The runtime probe reads and sets the style, then reads a formula back
through each of Range's formula properties and Evaluates R1C1 references
in both styles. The file probe saves a workbook in each style, records
the calcPr element Excel wrote, and opens the R1C1 one with the style at
A1 to see whether the workbook brings its style back. The style is put
back to A1 after each.

    python scripts/measure_reference_style.py

writes tests/fixtures/reference_style.json, which
tests/test_excel_reference_style.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/reference_style.json'

RUNTIME = '''Function Probe() As String
Dim ws As Worksheet, s As String, v As Variant
Set ws = Worksheets.Add
On Error Resume Next
s = Application.ReferenceStyle
Application.ReferenceStyle = xlR1C1
s = s & "|" & Err.Number & ":" & Application.ReferenceStyle: Err.Clear
ws.Range("A1").Value = 5: ws.Range("B2").Value = 7
ws.Range("C3").Formula = "=A1*2"
s = s & "|" & ws.Range("C3").Formula & "|" & ws.Range("C3").FormulaR1C1 & "|" & ws.Range("C3").FormulaLocal
s = s & "|" & ws.Range("C3").FormulaR1C1Local & "|" & ws.Range("A1").Address & "|" & ws.Range("C3").Value
ws.Range("D4").Formula = "=R1C1+1"
s = s & "|" & Err.Number & ":" & ws.Range("D4").Formula: Err.Clear
ws.Range("D5").FormulaLocal = "=R1C1+1"
s = s & "|" & Err.Number & ":" & ws.Range("D5").Formula: Err.Clear
For Each v In Array("=R1C1", "R1C1", "RC", "R[1]C[1]", "SUM(R1C1:R2C2)", "=A1")
s = s & "|" & ProbeEval(ws, CStr(v))
Next
For Each v In Array(3, 0, -1, 1, xlR1C1)
Application.ReferenceStyle = v
s = s & "|" & v & ">" & Err.Number & ":" & Application.ReferenceStyle: Err.Clear
Next
Application.ReferenceStyle = xlA1
For Each v In Array("=R1C1", "R1C1", "A1")
s = s & "|" & ProbeEval(ws, CStr(v))
Next
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
Probe = s
End Function
Function ProbeEval(ws As Worksheet, text As String) As String
Dim v As Variant
On Error Resume Next
v = ws.Evaluate(text)
If Err.Number <> 0 Then
ProbeEval = "e" & Err.Number
ElseIf IsError(v) Then
ProbeEval = CStr(v)
ElseIf IsObject(v) Then
ProbeEval = TypeName(v)
Else
ProbeEval = TypeName(v) & ":" & CStr(v)
End If
End Function
'''


def file_probe(folder: Path) -> str:
    paths = {style: str(folder / f'style-{style}.xlsx').replace('"', '""') for style in ('a1', 'r1c1')}
    return f'''Function Probe() As String
Dim wb As Workbook, s As String
Application.DisplayAlerts = False
Set wb = Workbooks.Add(xlWBATWorksheet)
wb.Worksheets(1).Range("A1").Value = 1
Application.ReferenceStyle = xlR1C1
wb.SaveAs Filename:="{paths['r1c1']}", FileFormat:=xlOpenXMLWorkbook
Application.ReferenceStyle = xlA1
wb.SaveAs Filename:="{paths['a1']}", FileFormat:=xlOpenXMLWorkbook
wb.Close False
s = Application.ReferenceStyle
Set wb = Workbooks.Open("{paths['r1c1']}")
s = s & "|" & Application.ReferenceStyle
wb.Close False
s = s & "|" & Application.ReferenceStyle
Application.ReferenceStyle = xlA1
Application.DisplayAlerts = True
Probe = s
End Function
'''


def main() -> None:
    record: dict[str, object] = {'measured_at': datetime.now(timezone.utc).isoformat()}
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function V() As String\nV = Application.Version & "|" & Application.Build\nEnd Function',
                                'V', line_numbers=False)
        record['excel_version_build'] = version.value
        runtime = excel.run_vba(RUNTIME, 'Probe', timeout=60.0, line_numbers=False)
        assert runtime.ok and not runtime.dialogs, runtime
        record['runtime'] = {'macro': RUNTIME, 'result': runtime.value}
        saved = excel.run_vba(file_probe(Path(folder)), 'Probe', timeout=60.0, line_numbers=False)
        assert saved.ok and not saved.dialogs, saved
        files = {}
        for style in ('a1', 'r1c1'):
            with zipfile.ZipFile(Path(folder) / f'style-{style}.xlsx') as package:
                workbook = package.read('xl/workbook.xml').decode('utf-8')
            found = re.search(r'<calcPr\b[^>]*/?>', workbook)
            files[style] = found.group(0) if found else ''
        record['file'] = {'opened': saved.value, 'calcPr': files}
    OUT.write_text(json.dumps(record, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(record, indent=2)[-1500:])


if __name__ == '__main__':
    main()
