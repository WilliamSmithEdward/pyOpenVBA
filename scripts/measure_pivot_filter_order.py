"""PivotFilter.Order and PivotItems.Add: what each filter's Order reads before and after a change, the errors, each
filter's evalOrder as the saved pivot table definition holds it, and the Region field's items as the table and the
cache save them.

Each case makes a workbook with a list of regions, products and amounts on its first sheet, builds a pivot table on a
second sheet with Region and Product as row fields and the sum of Amount as data, runs the case's statements with pt
as the table, then reads every filter of the table with its field and Order, and saves it.

    python scripts/measure_pivot_filter_order.py

writes tests/fixtures/pivot_filter_order.json, which tests/test_excel_pivot_filter_order.py replays.
"""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/pivot_filter_order.json'
REGION = 'pt.PivotFields("Region").PivotFilters.Add2 xlCaptionBeginsWith, Value1:="E"'
PRODUCT = 'pt.PivotFields("Product").PivotFilters.Add2 xlCaptionContains, Value1:="n"'
CASES = [
    REGION,
    f'{REGION}: {PRODUCT}',
    f'{REGION}: {PRODUCT}: pt.PivotFields("Product").PivotFilters(1).Order = 1',
    f'{REGION}: {PRODUCT}: pt.PivotFields("Region").PivotFilters(1).Order = 2',
    f'{REGION}: {PRODUCT}: pt.PivotFields("Region").PivotFilters(1).Order = 3',
    f'{REGION}: {PRODUCT}: pt.PivotFields("Region").PivotFilters(1).Order = 0',
    f'{REGION}: pt.PivotFields("Region").PivotFilters(1).Order = 1',
    f'{REGION}: {PRODUCT}: pt.PivotFields("Region").ClearAllFilters',
    f'{PRODUCT}: {REGION}',
    's = TypeName(pt.PivotFields("Region").PivotItems.Add("North"))',
    'pt.PivotFields("Region").PivotItems.Add "North": s = pt.PivotFields("Region").PivotItems.Count & "," & '
    'pt.TableRange1.Rows.Count',
]


def _probe() -> str:
    lines = ['Private Function Tried(ByVal which As Long, ByVal folder As String) As String',
             '    Dim wb As Workbook, ws As Worksheet, pt As Object, s As String, f As Object, flt As Object',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    Set ws = wb.Worksheets(1)', '    ws.Name = "data"',
             '    ws.Range("A1:C1").Value = Array("Region", "Product", "Amount")',
             '    ws.Range("A2:C2").Value = Array("East", "Pen", 10)',
             '    ws.Range("A3:C3").Value = Array("East", "Ink", 20)',
             '    ws.Range("A4:C4").Value = Array("West", "Pen", 30)',
             '    Set ws = wb.Worksheets.Add(After:=ws)', '    ws.Name = "report"',
             '    Set pt = wb.PivotCaches.Create(xlDatabase, "data!R1C1:R4C3").CreatePivotTable(ws.Range("A3"), "P")',
             '    pt.PivotFields("Region").Orientation = xlRowField',
             '    pt.PivotFields("Product").Orientation = xlRowField',
             '    pt.AddDataField pt.PivotFields("Amount"), "Sum of Amount", xlSum',
             '    On Error GoTo Failed', '    Select Case which']
    for index, statements in enumerate(CASES):
        lines += [f'    Case {index}', f'        {statements}']
    lines += ['    End Select',
              '    For Each f In pt.RowFields',
              '        For Each flt In f.PivotFilters: s = s & "|" & f.Name & ":" & flt.FilterType & ":" & flt.Order: Next',
              '    Next',
              '    Tried = "ok " & s', 'Saved:',
              '    wb.SaveAs folder & "\\c" & which & ".xlsx", xlOpenXMLWorkbook', '    wb.Close False',
              '    Exit Function', 'Failed:', '    Tried = "Err " & Err.Number', '    Resume Saved', 'End Function', '',
              'Public Function Probe(ByVal folder As String) As String', '    Dim i As Long, s As String',
              '    Application.DisplayAlerts = False', f'    For i = 0 To {len(CASES) - 1}',
              '        s = s & "@@@" & Tried(i, folder)', '    Next', '    Application.DisplayAlerts = True',
              '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def filters(path: Path) -> list[str]:
    """Each saved filter's field, type, evalOrder and id, in the order the definition holds them; then the Region
    field's items as the table lists them and its shared items as the cache keeps them."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        table = next(one for one in names if re.fullmatch(r'xl/pivotTables/pivotTable\d+\.xml', one))
        xml = package.read(table).decode('utf-8')
        cache = package.read(next(one for one in names if re.fullmatch(
            r'xl/pivotCache/pivotCacheDefinition\d+\.xml', one))).decode('utf-8')
    found = [' '.join(re.findall(r'\b(?:fld|type|evalOrder|id)="[^"]*"', one))
             for one in re.findall(r'<filter\b[^>]*>', xml)]
    region = re.search(r'<pivotField\b[^>]*>(.*?)</pivotField>', xml, re.S)
    shared = re.search(r'<cacheField name="Region"[^>]*>(.*?)</cacheField>', cache, re.S)
    return found + [region.group(1) if region else '', shared.group(1) if shared else '']


def main() -> None:
    source = _probe()
    with tempfile.TemporaryDirectory() as folder, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', (folder,), timeout=280.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
        saved = [filters(Path(folder) / f'c{index}.xlsx') for index in range(len(CASES))]
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[case, answer, found] for case, answer, found in zip(CASES, answers, saved, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for case, answer, found in record['cases']:
        print(case[-60:], '->', answer, found)


if __name__ == '__main__':
    main()
