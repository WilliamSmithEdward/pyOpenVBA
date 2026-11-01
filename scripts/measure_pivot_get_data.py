"""PivotTable.GetData: what it answers for names picking data fields and items, quoted and not, in either order and
any case, with field[item], Grand Total, names the table lacks, and items whose cell shows no value.

    python scripts/measure_pivot_get_data.py

writes tests/fixtures/pivot_get_data.json, which tests/test_excel_pivot_get_data.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_get_data.json'
#: Names asked of a table with k on rows, p on columns, and the data fields Sum of v and Count of q.
NAMES = ['Sum of v', "'Sum of v'", "'Sum of v' a", 'Sum of v a', "'Sum of v' k[a]", "'Sum of v' k['a']", 'a',
         "'Sum of v' 'a'", "'Sum of v' x", "'Sum of v' a x", "'Sum of v' 'Grand Total'", "'Sum of v' zz", '',
         "'Count of q'", "'Count of q' b", "'Sum of v' A", "'sum of v' a", 'Sum of v[a]', "'Sum of v' a y",
         "'Sum of v' k[a] p[x]", "'Sum of v' x a", "'Sum of v' b x", "'Sum of v' x y", "'Sum of v' a a",
         "'Sum of v' k[zz]", "'Sum of v' zz[a]", "'Sum of v' 'a", "'Sum of v' k[a", 'v', "v a",
         "'Sum of v' Grand Total", "'Sum of v' grand total x", "Count of q Sum of v"]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal pt As Object, ByVal n As String) As String', '    Dim v As Variant',
             '    On Error GoTo Failed', '    v = pt.GetData(n)', '    Tried = TypeName(v) & ":" & CStr(v)',
             '    Exit Function', 'Failed:', '    Tried = "E" & Err.Number', 'End Function', '',
             'Public Function Probe() As String', '    Dim wb As Workbook, ws As Worksheet, pt As Object, s As String',
             '    Dim n As Variant', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)',
             '    ws.Range("A1:D5").Value = Application.Transpose(Application.Transpose(Array(Array("k", "p", "v", "q"), '
             'Array("a", "x", 1, "m"), Array("b", "y", 2, "n"), Array("a", "y", 3, "m"), Array("c", "x", 4, "n"))))',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:D5")).CreatePivotTable(ws.Range("I3"), "PT")',
             '    pt.PivotFields("k").Orientation = xlRowField', '    pt.PivotFields("p").Orientation = xlColumnField',
             '    pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum',
             '    pt.AddDataField pt.PivotFields("q"), "Count of q", xlCount']
    listed = ', '.join('"' + name.replace('"', '""') + '"' for name in NAMES)
    lines += [f'    For Each n In Array({listed})', '        s = s & "~" & Tried(pt, CStr(n))', '    Next',
              '    wb.Close False', '    Application.DisplayAlerts = True', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (), timeout=300.0)
        assert result.ok and not result.dialogs, result
    answers = str(result.value).split('~')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'names': [[name, answer] for name, answer in zip(NAMES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer in record['names']:
        print(repr(name), '->', answer)


if __name__ == '__main__':
    main()
