"""Workbook.TableStyles and the built-in TableStyle objects: what a style of each family reads, which galleries it
shows in, what Delete and Duplicate do to a built-in one, and the collection's count and order.

Each read is its value (TypeName for an object) or "E" and the error number, on a new workbook.

    python scripts/measure_table_styles.py

writes tests/fixtures/table_styles.json, which tests/test_excel_table_styles.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/table_styles.json'
STYLES = ['TableStyleMedium2', 'TableStyleLight1', 'TableStyleDark11', 'PivotStyleLight16', 'PivotStyleDark28',
          'SlicerStyleLight1', 'SlicerStyleOther2', 'TimeSlicerStyleLight1', 'TimeSlicerStyleDark6']
MEMBERS = ['Name', 'NameLocal', 'BuiltIn', 'ShowAsAvailableTableStyle', 'ShowAsAvailablePivotTableStyle',
           'ShowAsAvailableSlicerStyle', 'ShowAsAvailableTimelineStyle']
READS = [f'wb.TableStyles("{style}").{name}' for style in STYLES for name in MEMBERS] + [
    'TypeName(wb.TableStyles("TableStyleMedium2").Parent)', 'TypeName(wb.TableStyles("TableStyleMedium2").Application)',
    'TypeName(wb.TableStyles("TableStyleMedium2").TableStyleElements)', 'wb.TableStyles.Count',
    'TypeName(wb.TableStyles)', 'TypeName(wb.TableStyles.Parent)', 'TypeName(wb.TableStyles.Application)',
    'wb.TableStyles(1).Name', 'wb.TableStyles(2).Name', 'wb.TableStyles(61).Name', 'wb.TableStyles(62).Name',
    'wb.TableStyles(145).Name', 'wb.TableStyles(146).Name', 'wb.TableStyles(159).Name', 'wb.TableStyles(160).Name',
    'wb.TableStyles(wb.TableStyles.Count).Name', 'wb.TableStyles(0).Name', 'wb.TableStyles(9999).Name',
    'wb.TableStyles("NoSuchStyle").Name', 'wb.TableStyles("tablestylemedium2").Name',
]
STATEMENTS = ['wb.TableStyles("TableStyleMedium2").Delete',
              'wb.TableStyles("TableStyleMedium2").ShowAsAvailableTableStyle = False']


def _probe() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Workbook, s As String, v As String',
             '    Set wb = Workbooks.Add', '    On Error Resume Next']
    for read in READS:
        lines += ['    Err.Clear', '    v = ""', f'    v = {read}', '    If Err.Number Then v = "E" & Err.Number',
                  '    s = s & "@@@" & v']
    lines += ['    Dim one As Object, all As String', '    For Each one In wb.TableStyles',
              '        all = all & "|" & one.Name & "=" & one.NameLocal', '    Next', '    s = s & "@@@" & Mid(all, 2)']
    for statement in STATEMENTS:
        lines += ['    Err.Clear', f'    {statement}', '    s = s & "@@@" & "E" & Err.Number & "/" & '
                  'wb.TableStyles("TableStyleMedium2").ShowAsAvailableTableStyle & "/" & wb.TableStyles.Count']
    lines += ['    On Error GoTo 0', '    wb.Close False', '    Probe = Mid(s, 4)', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    source = _probe()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Probe', timeout=200.0)
        assert result.ok and not result.dialogs, result
        answers = str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [[read, answer] for read, answer in zip([*READS, 'every Name=NameLocal', *STATEMENTS], answers,
                                                               strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for read, answer in record['cases']:
        print(read, '->', answer)


if __name__ == '__main__':
    main()
