"""What a call that does not fit the member answers bound late, for each form of member the type library has.

The lines are those of scripts/measure_vba_compile_kinds.py whose arguments do not fit, run through variables As Object under
On Error Resume Next. Such a call fails at the object's Invoke before the member runs, so nothing changes. Application is left
out: reading its members with arguments has shown a modal dialog of Excel's own.

    python scripts/measure_vba_late_errors.py

writes tests/fixtures/vba_late_errors.json, keeps what it holds and asks only the rest.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))

from measure_vba_compile_kinds import asked  # noqa: E402

OUT = ROOT / 'tests/fixtures/vba_late_errors.json'
FAMILIES = ('too_many', 'too_few', 'args_on_no_args', 'args_on_function', 'bad_name')
SETUP = '''    Dim r As Object, ws As Object, wb As Object, f As Object, sh As Object, lo As Object, ch As Object, ss As Object, n As Object, v As Variant, o As Object, s As String
    Dim se As Object, wi As Object, pt As Object, cm As Object, hl As Object, lc As Object, co As Object, it As Object, ax As Object, va As Object, ps As Object, wbs As Object
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("H1", "H2"): ws.Range("A2:B4").Value = 2
    Set r = ws.Range("C7"): Set f = r.Font: Set it = r.Interior
    Set sh = ws.Shapes.AddShape(1, 300, 100, 40, 40)
    Set lo = ws.ListObjects.Add(1, ws.Range("A1:B4"), , 1): Set lc = lo.ListColumns(1)
    Set co = ws.ChartObjects.Add(300, 200, 200, 150): Set ch = co.Chart
    ch.ChartType = xlColumnClustered: ch.SetSourceData ws.Range("A1:B4")
    Set se = ch.SeriesCollection(1): Set ax = ch.Axes(1)
    Set ss = wb.Worksheets: Set n = wb.Names.Add("Nm", "=1"): Set wi = wb.Windows(1)
    Set cm = ws.Range("E9").AddComment("x"): Set hl = ws.Hyperlinks.Add(ws.Range("E8"), "http://x.test/")
    Set va = ws.Range("E7").Validation: Set ps = ws.PageSetup: Set wbs = Workbooks
    Set pt = wb.PivotCaches.Create(xlDatabase, ws.Range("A1:B4")).CreatePivotTable(ws.Range("H20"))
    On Error Resume Next
'''
LEAVE = '''    On Error GoTo 0
    Application.DisplayAlerts = False
    wb.Close False
    Application.DisplayAlerts = True
    Probe = s
'''
SIZE = 30


def save(answers: dict[str, str]) -> None:
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=0) + '\n', encoding='utf-8')


def lines() -> dict[str, str]:
    chosen: dict[str, str] = {}
    for key, line in asked().items():
        head, rest = key.split(': ', 1)
        family = rest.split(' ', 1)[0]
        if head != 'a' and family in FAMILIES and '\n' not in line and line.startswith('v = ') and 'Select' not in line:
            chosen[key] = line
    return chosen


def main() -> None:
    size = SIZE
    answers: dict[str, str] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    wanted = {key: line for key, line in lines().items() if key not in answers}
    print(len(wanted), 'left of', len(lines()), flush=True)
    pending = list(wanted.items())
    while pending:
        batch, pending = pending[:size], pending[size:]
        body = ''.join(f'    Err.Clear\n    {line}\n    s = s & Err.Number & ";"\n' for _, line in batch)
        source = 'Function Probe() As String\n' + SETUP + body + LEAVE + 'End Function\n'
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(source, 'Probe', timeout=120.0, line_numbers=False)
        if not result.ok or result.dialogs:
            if len(batch) == 1:
                answers[batch[0][0]] = 'stopped'
            else:
                middle = len(batch) // 2
                print('split', batch[0][1], result.message, [d.message[:50] for d in result.dialogs], flush=True)
                pending = batch + pending
                size = max(1, middle)
            save(answers)
            continue
        for (key, _), answer in zip(batch, str(result.value).split(';'), strict=False):
            answers[key] = answer
        save(answers)
        print(len(answers), flush=True)


if __name__ == '__main__':
    main()
