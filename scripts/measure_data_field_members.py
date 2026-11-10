"""PivotField members read on a pivot table's data field, beside the same on its row field and its first field. PivotField round
58101 had Excel raise 1004 for ParentItems, ChildItems and CalculatedItems on the data field and 5 for AllItemsVisible, where the
model answered. Rules written before this probe:

A. The data field answers as any field (what the model did).
B. ParentItems, ChildItems and CalculatedItems are 1004 and AllItemsVisible 5 on the data field; the rest answer as on any field.
C. As B, and every collection the data field holds (PivotItems, VisibleItems, HiddenItems) is 1004 too.

Each read is its own macro on a new book and on tests/fixtures/dxf_table/made.xlsx, late bound and typed, reading the type and Err
straight after; by live Excel and the model, written to tests/fixtures/data_field_members.json.

    python scripts/measure_data_field_members.py
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as rounds  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'data_field_members.json'
READS = ['AllItemsVisible', 'CalculatedItems.Count', 'ChildField.Name', 'ChildItems.Count', 'DataType', 'GroupLevel', 'HiddenItems.Count',
         'IsCalculated', 'LayoutPageBreak', 'ParentField.Name', 'ParentItems.Count', 'PivotItems.Count', 'ShowingInAxis', 'SourceCaption',
         'SourceName', 'TotalLevels', 'VisibleItems.Count', 'Orientation', 'Name']
FIELDS = {'data': 'DataFields(1)', 'row': 'RowFields(1)', 'first': 'PivotFields(1)'}
BOOKS = {
    'new': ('    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)\n'
            '    ws.Range("A1:B3").Value = Array("k", "v"): ws.Range("A2:B2").Value = Array("a", 1): ws.Range("A3:B3").Value = Array("b", 2)\n'
            '    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B3")).CreatePivotTable(ws.Range("E3"), "PT")\n'
            '    pt.PivotFields("k").Orientation = xlRowField: pt.AddDataField pt.PivotFields("v"), "Sum of v", xlSum\n'),
    'file': '    Set wb = Workbooks.Open(path & "\\t.xlsx"): Set pt = wb.Worksheets("pivot").PivotTables(1)\n',
}


def cases() -> list[dict[str, str]]:
    return [{'book': book, 'field': field, 'read': read} for book in BOOKS for field in FIELDS for read in READS]


def source(case: dict[str, str], typed: bool) -> str:
    kind = 'PivotField' if typed else 'Object'
    return ('Public Function One(ByVal path As String) As String\n'
            f'    Dim wb As Object, ws As Object, pt As Object, o As {kind}, v As Variant, s As String\n    Application.DisplayAlerts = False\n'
            + BOOKS[case['book']] + f'    Set o = pt.{FIELDS[case["field"]]}\n    On Error Resume Next\n'
            f'    v = o.{case["read"]}\n    s = TypeName(v) & ":" & v & " E" & Err.Number\n    On Error GoTo 0\n'
            '    wb.Close False\n    One = s\nEnd Function\n')


def model(case: dict[str, str], typed: bool, folder: str) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case, typed), name='Fields')
    try:
        return str(app.run('Fields.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[union-attr]
    made = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        rounds.books(Path(held))
        lock = _lock()
        owner = rounds.take_lock(lock, 'Verification 02')
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in made:
                    for tag, typed in (('late', False), ('typed', True)):
                        done = excel.run_vba(source(case, typed), 'One', (held,), timeout=60.0, line_numbers=False)
                        case[tag] = str(done.value) if done.ok else f'no answer: {str(done.error)[:200]}'
        finally:
            rounds.release_lock(lock, owner)
        for case in made:
            for tag, typed in (('late', False), ('typed', True)):
                case[f'model_{tag}'] = model(case, typed, held)
    OUT.write_text(json.dumps({'cases': made}, indent=1) + '\n', encoding='utf-8')
    for case in made:
        for tag in ('late', 'typed'):
            if case[tag] != case[f'model_{tag}']:
                print('DIFF', tag, case['book'], case['field'], case['read'], '| Excel', case[tag], '| model', case[f'model_{tag}'])
    print('cases', len(made))


if __name__ == '__main__':
    main()
