"""The error from assigning 1 to every scalar read-only property of the objects scripts/measure_object_trio.py reaches.

Nothing in the type library says which error a member raises, so each is measured. A member that answers differently
by the way it is reached is left out: the error there follows the object's state, as a whole sheet's Range.Count
overflows before the assignment fails.

    python scripts/measure_read_only_lets.py

writes tests/fixtures/read_only_lets.json, which tests/test_excel_object_trio.py replays. Run measure_object_trio.py
first, and against the checklist the rows are read from.
"""
from __future__ import annotations

import csv
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import measure_object_trio as trio
from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/read_only_lets.json'
CHUNK = 60
SCALARS = {'String', 'Long', 'Boolean', 'Double', 'Single', 'Integer', 'Date', 'Variant', 'Currency'}
#: Worksheet.FilterMode takes the assignment and does nothing, which the model does not do. The other two fail in the
#: read the model does first: International with no index and a chart's DownBars.Name. DisplayFormat is 450 late bound
#: but 424 bound early, which is how a macro meets it (tests/fixtures/conditional/display.json).
LEFT_OUT = {('Worksheet', 'FilterMode'), ('Application', 'International'), ('DownBars', 'Name')}


def cases() -> list[tuple[str, str, int]]:
    """Every (class, member, way in) with a scalar read-only property and no arguments."""
    rows = list(csv.DictReader((ROOT / 'docs/excel_checklist.csv').open(encoding='utf-8', newline='')))
    found = []
    for index, (owner, _) in enumerate(trio.REACHES):
        for row in rows:
            kind = row['signature'].removeprefix('As ')
            if (row['object'] == owner and row['kind'] == 'property, read-only' and row['signature'].startswith('As ')
                    and row['status'] != 'MISSING' and not row['member'].startswith('_')
                    and row['member'] not in ('Application', 'Creator', 'Parent')
                    and (kind in SCALARS or kind.startswith(('Xl', 'Mso')))):
                found.append((owner, row['member'], index))
    return found


def _probe(chosen: list[tuple[str, str, int]], log: str) -> str:
    source = trio.source()
    body = source[:source.index('Public Function Probe()')]
    calls = []
    for number, (_, member, index) in enumerate(chosen):
        body += (f'\nPrivate Function L{number}() As Long\n    Dim o As Object\n    On Error Resume Next\n'
                 f'    Set o = Reached({index})\n    o.{member} = 1\n    L{number} = Err.Number\nEnd Function\n')
        calls.append(f'    Open "{log}" For Append As #1: Print #1, {number}: Close #1\n'
                     f'    out = out & "@" & L{number}()\n')
    return (body + 'Public Function Probe() As String\n    Dim out As String\n    Setup\n' + ''.join(calls)
            + '    Probe = Mid(out, 2)\n    wb.Close False\nEnd Function\n')


def main() -> None:
    pending = cases()
    answers: dict[tuple[str, str, int], int] = {}
    excel = None
    with tempfile.TemporaryDirectory() as folder:
        log = str(Path(folder) / 'progress.txt')
        while pending:
            chosen = pending[:CHUNK]
            Path(log).unlink(missing_ok=True)
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False))
                excel.__enter__()
                excel.new_document()
            result = excel.run_vba(_probe(chosen, log), 'Probe', (), timeout=120.0)
            if result.ok and not result.dialogs:
                answers.update(zip(chosen, map(int, str(result.value).split('@')), strict=True))
                pending = pending[len(chosen):]
                continue
            seen = Path(log).read_text().split() if Path(log).exists() else ['0']
            print('stopped at', chosen[int(seen[-1])], result.outcome)
            pending.remove(chosen[int(seen[-1])])
            excel.__exit__(None, None, None)
            excel = None
    if excel is not None:
        excel.__exit__(None, None, None)
    by_member: dict[tuple[str, str], set[int]] = {}
    for (owner, member, _), number in answers.items():
        by_member.setdefault((owner, member), set()).add(number)
    kept = [[owner, member, index, number] for (owner, member, index), number in answers.items()
            if len(by_member[(owner, member)]) == 1 and (owner, member) not in LEFT_OUT and owner != 'DisplayFormat']
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': kept}, indent=0) + '\n',
                   encoding='utf-8')
    print(len(kept), 'kept of', len(answers))


if __name__ == '__main__':
    main()
