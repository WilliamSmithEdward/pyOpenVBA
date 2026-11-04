"""What a copied sheet is called when its name is taken, for names already ending in a number in brackets.

    python scripts/measure_sheet_copy_names.py

writes tests/fixtures/sheet_copy_names.json, which tests/test_excel_sheet_copy_names.py replays. Each case names the sheets
of a new workbook and copies the last one after itself.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/sheet_copy_names.json'
CASES = (('S',), ('S (2)',), ('S', 'S (2)'), ('S (5)',), ('S (2)', 'S (5)'), ('S (3)', 'S (2)'), ('S (1)',), ('S (0)',), ('S (02)',), ('S(2)',), ('(2)',), ('S (x)',), ('S (-2)',),
         ('S  (2)',), ('s (2)', 'S (3)'), ('A' * 27 + ' (2)',), ('A' * 31,), ('S (99999)',))


def module() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Workbook, s As String', '    Application.DisplayAlerts = False']
    for names in CASES:
        lines += ['    Set wb = Workbooks.Add(xlWBATWorksheet)', f'    wb.Worksheets(1).Name = "{names[0]}"']
        lines += [f'    wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).Name = "{name}"' for name in names[1:]]
        lines += ['    On Error Resume Next', '    Err.Clear', '    wb.Worksheets(wb.Worksheets.Count).Copy After:=wb.Worksheets(wb.Worksheets.Count)',
                  '    s = s & wb.Worksheets.Count & "|" & wb.Worksheets(wb.Worksheets.Count).Name & "|" & Err.Number & vbLf', '    On Error GoTo 0', '    wb.Close False']
    lines += ['    Application.DisplayAlerts = True', '    Probe = s', 'End Function']
    return '\n'.join(lines) + '\n'


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=600.0)) as excel:
        excel.new_document()
        result = excel.run_vba(module(), 'Probe', (), timeout=120.0)
    answers = str(result.value).strip().split('\n')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': [[list(names), answer] for names, answer in zip(CASES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for names, answer in record['cases']:
        print(names, '=>', answer)


if __name__ == '__main__':
    main()
