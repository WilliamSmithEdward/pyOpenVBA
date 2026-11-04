"""Every AutoShape type's adjustments: its default values, and each one's name and scale in the file, from
Adjustments(k) = default + 0.1 and the saved shape.

    python scripts/measure_shape_adjustment_table.py

writes tests/fixtures/shape_adjustment_table.json, from which src/pyopenvba/shapes/_adjustments.py was made and which
tests/test_excel_shape_adjustments.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/shape_adjustment_table.json'
SOURCE = '''
Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String
    Dim t As Long, k As Long, wb As Workbook, ws As Worksheet, sh As Object, s As String, n As Long, d As Double
    Application.DisplayAlerts = False
    For t = first To last
        Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
        On Error Resume Next
        Err.Clear
        Set sh = Nothing
        Set sh = ws.Shapes.AddShape(t, 10, 10, 100, 40)
        If Err.Number <> 0 Or sh Is Nothing Then
            s = s & "@@@" & t & "|E" & Err.Number
            Err.Clear
        Else
            n = sh.Adjustments.Count
            s = s & "@@@" & t & "|" & sh.AutoShapeType & "|" & sh.Name & "|" & n
            For k = 1 To n
                s = s & "|" & sh.Adjustments(k)
            Next
            For k = 1 To n
                d = sh.Adjustments(k)
                sh.Adjustments(k) = d + 0.1
                If Err.Number <> 0 Then s = s & "|S" & k & "E" & Err.Number: Err.Clear
                s = s & "|R" & k & "=" & sh.Adjustments(k)
            Next
            s = s & "|T=" & sh.Type
            wb.SaveAs folder & "\\t" & t & ".xlsx", xlOpenXMLWorkbook
        End If
        On Error GoTo 0
        wb.Close False
    Next
    Application.DisplayAlerts = True
    Probe = Mid(s, 4)
End Function
'''


def main() -> None:
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(SOURCE, 'Probe', (folder, 1, 190), timeout=500.0)
        assert result.ok and not result.dialogs, result
        rows = []
        for line in str(result.value).split('@@@'):
            path = Path(folder) / f"t{line.split('|')[0]}.xlsx"
            geometry = ''
            if path.exists():
                with zipfile.ZipFile(path) as package:
                    xml = package.read('xl/drawings/drawing1.xml').decode('utf-8')
                found = re.search(r'<a:prstGeom.*?</a:prstGeom>|<a:custGeom', xml, re.DOTALL)
                geometry = found.group(0) if found else ''
                body = re.search(r'<xdr:txBody>.*?</xdr:txBody>', xml, re.DOTALL)
                text = body.group(0) if body else ''
            else:
                text = ''
            rows.append([line, geometry, text])
    OUT.write_text(json.dumps({'source': SOURCE, 'rows': rows}, indent=0) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
