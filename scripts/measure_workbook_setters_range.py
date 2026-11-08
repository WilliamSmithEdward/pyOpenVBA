"""A Range given to Workbook settings in live Excel, typed and late, with the model's answers written down first: one cell holding True,
one holding 5, an empty one, and one area of two cells. Each case on a new workbook: the set's error and the setting read after.

    python scripts/measure_workbook_setters_range.py predict
    python scripts/measure_workbook_setters_range.py measure

writes tests/fixtures/workbook_setters_range.json.
"""
from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/workbook_setters_range.json'
SEPARATOR = '%%'
SETTINGS = ('CheckCompatibility', 'DisplayInkComments', 'EnableAutoRecover', 'ForceFullCalculation', 'InactiveListBorderVisible',
            'ShowPivotChartActiveFields', 'ShowPivotTableFieldList', 'UpdateLinks', 'Date1904', 'PrecisionAsDisplayed', 'SaveLinkValues',
            'DisplayDrawingObjects')
RANGES = ('ws.Range("A1")', 'ws.Range("B1")', 'ws.Range("C1")', 'ws.Range("A1:B1")')


def macro(declared: str) -> str:
    cases = ''.join(f'''    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = True: ws.Range("B1").Value = 5
    Err.Clear
    wb.{setting} = {value}
    s = "e" & Err.Number
    Err.Clear: v = Empty
    v = wb.{setting}
    out = out & "{SEPARATOR}" & s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
    wb.Close SaveChanges:=False
''' for setting in SETTINGS for value in RANGES)
    return f'''Function Probe() As String
    Dim wb As {declared}, ws As Object, v As Variant, s As String, out As String
    Application.DisplayAlerts = False
    On Error Resume Next
{cases}    On Error GoTo 0
    Application.DisplayAlerts = True
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def predict() -> None:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication

    answers: dict[str, list[str]] = {}
    for declared in ('Workbook', 'Object'):
        app = ExcelApplication()
        app.add_workbook()
        app.add_module(macro(declared), name='H')
        answers[declared] = str(app.run('H.Probe')).split(SEPARATOR)
    text = json.dumps({'settings': SETTINGS, 'ranges': RANGES, 'predicted': answers}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print('predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    answers: dict[str, list[str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
        for declared in ('Workbook', 'Object'):
            result = excel.run_vba(macro(declared), 'Probe', (), timeout=300, line_numbers=False)
            assert result.ok, (result.error, [one.message for one in result.dialogs])
            answers[declared] = str(result.value).split(SEPARATOR)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    cases = [(setting, value) for setting in SETTINGS for value in RANGES]
    for declared in answers:
        for (setting, value), mine, theirs in zip(cases, record['predicted'][declared], answers[declared]):
            if mine != theirs:
                print(declared, setting, value, 'excel', theirs, 'model', mine)
    print('misses', sum(a != b for d in answers for a, b in zip(record['predicted'][d], answers[d])), 'of', 2 * len(cases))


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
