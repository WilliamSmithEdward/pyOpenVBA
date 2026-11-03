"""Range.Cut through a table: where the table goes, what its cells hold and what the table part says.

Each case makes a workbook with a table T on A1:C5 of Sheet1, a formula on H10 and a value on H1, runs the case's statements,
reads READS, each as its value or "E" and an error number, and saves the workbook.

    python scripts/measure_table_cut.py

writes tests/fixtures/table_cut.json, which tests/test_excel_table_cut.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/table_cut.json'
CASES = [
    '',
    # All of it.
    'ws.Range("A1:C5").Cut ws.Range("D3")', 'ws.Range("A1:C5").Cut ws.Range("B2")', 'ws.Range("A1:C5").Cut wb.Worksheets(2).Range("B2")', 'o.Range.Cut ws.Range("J1")',
    'o.ShowTotals = True: o.Range.Cut ws.Range("J1")', 'ws.Range("A1:C5").Cut: ws.Paste ws.Range("J1")', 'ws.Range("A1:C5").Cut: ws.Range("J1").Select: ActiveSheet.Paste', 'ws.Range("A1:D6").Cut ws.Range("J1")', 'ws.Range("A1:C6").Cut ws.Range("J1")', 'ws.Range("A1:D5").Cut ws.Range("J1")', 'ws.Rows("1:5").Cut ws.Rows(11)', 'ws.Columns("A:C").Cut ws.Columns("J")',
    'o.Range.AutoFilter 1, ">1": o.Range.Cut ws.Range("J1")', 'ws.Range("A1:C5").Cut ws.Range("A1")',
    # Part of it.
    'ws.Range("A2:C5").Cut ws.Range("J1")', 'o.DataBodyRange.Cut ws.Range("J1")', 'ws.Range("A3:C3").Cut ws.Range("J1")', 'ws.Range("B1:B5").Cut ws.Range("J1")', 'ws.Range("B2:B5").Cut ws.Range("J1")', 'ws.Range("B3").Cut ws.Range("J1")', 'ws.Range("A1:C1").Cut ws.Range("J1")', 'ws.Range("A1").Cut ws.Range("J1")', 'ws.Range("A1:B2").Cut ws.Range("J1")',
    # Within it.
    'ws.Range("A2").Cut ws.Range("B3")', 'ws.Range("A2:A3").Cut ws.Range("C4")', 'ws.Range("A1").Cut ws.Range("B1")', 'ws.Range("A2").Cut ws.Range("A1")', 'ws.Range("A1:A5").Cut ws.Range("C1")', 'ws.Range("A3:C3").Cut ws.Range("A5")',
    # Onto it.
    'ws.Range("H1").Cut ws.Range("B3")', 'ws.Range("H1").Cut ws.Range("B1")', 'ws.Range("H1").Cut ws.Range("D2")', 'ws.Range("H1").Cut ws.Range("B6")', 'ws.Range("H1:I2").Cut ws.Range("C5")', 'ws.Range("H1:I2").Cut ws.Range("A1")', 'wb.Worksheets(2).Range("A1").Cut ws.Range("B3")',
]
READS = ['ws.ListObjects.Count', 'wb.Worksheets(2).ListObjects.Count', 'o.Range.Parent.Name', 'o.Range.Address', 'o.DataBodyRange.Address', 'o.HeaderRowRange.Address', 'o.ListColumns(1).Name', 'o.ListColumns(2).Name', 'ws.Range("A1").Value', 'ws.Range("B3").Value', 'ws.Range("D3").Value',
         'ws.Range("J1").Value', 'ws.Range("H10").Formula', 'ws.Range("H10").Value', 'ws.AutoFilterMode', 'o.ShowAutoFilter', 'o.AutoFilter.Range.Address', 'Application.CutCopyMode', 'ws.Range("A1").ListObject Is Nothing', 'ws.Range("J1").ListObject Is Nothing', 'ws.Range("A2").Interior.Color']


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Object, o As Object, s As String, v As String', '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    wb.Worksheets.Add After:=wb.Worksheets(1)', '    Set ws = wb.Worksheets(1)', '    ws.Activate',
             '    wb.Worksheets(2).Range("A1").Value = "z"', '    ws.Range("A1:C1").Value = Array("a", "b", "c")', '    ws.Range("A2:C2").Value = Array(1, 2, 3)',
             '    ws.Range("A3:C3").Value = Array(4, 5, 6)', '    ws.Range("A4:C4").Value = Array(7, 8, 9)', '    ws.Range("A5:C5").Value = Array(10, 11, 12)',
             '    Set o = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C5"), , xlYes)', '    o.Name = "T"', '    ws.Range("H1").Value = "h"', '    ws.Range("I1").Value = "i"', '    ws.Range("H10").Formula = "=SUM(T[a])"', '    On Error GoTo Failed', '    Select Case which']
    for index, statement in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statement}']
    lines += ['    End Select', '    s = "ok "', 'Answered:', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number', '    s = s & "|" & v']
    lines += ['    On Error GoTo 0', '    Tried = s', '    Application.CutCopyMode = False', '    Application.DisplayAlerts = False', '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook',
              '    wb.Close False', '    Application.DisplayAlerts = True', '    Exit Function', 'Failed:', '    s = "Err " & Err.Number & " "', '    Resume Answered', 'End Function', '',
              'Public Function Probe(ByVal folder As String, ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String',
              '    For i = first To last', '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def parts_of(path: Path) -> dict[str, str]:
    found: dict[str, str] = {}
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if re.fullmatch(r'xl/worksheets/sheet\d\.xml|xl/tables/table\d+\.xml|xl/styles\.xml', name):
                found[name] = package.read(name).decode('utf-8')
    return found


def main() -> None:
    source = _probe()
    answers: list[str] = []
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(CASES), 20):
            last = min(first + 19, len(CASES) - 1)
            result = excel.run_vba(source, 'Probe', (folder, first, last), timeout=280.0)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
        parts = [parts_of(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'reads': READS, 'source': source,
              'cases': [[case, answer, part] for case, answer, part in zip(CASES, answers, parts, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    print(READS)
    for case, answer, _ in record['cases']:
        print(case, '=>', answer)


if __name__ == '__main__':
    main()
