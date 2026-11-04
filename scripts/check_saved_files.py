"""The save gate for an audit's fixture: the model makes every case and saves, and live Excel opens that file and reads each
case back. The answers should be what Excel read after its own save. A refused file is the worst result this finds, and
it found a huge number in plain digits doing that.

    python scripts/check_saved_files.py [format_properties member_audit method_audit range_values]

prints the differences by member, and writes nothing. Cases the fixture's "early" holds are left out, and so are those the
model reports as gaps. Run it after changing what a save writes.
"""
from __future__ import annotations

import collections
import json
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.exceptions import VBAUnsupportedError

FIXTURES = Path(__file__).resolve().parents[1] / 'tests/fixtures'
#: Each fixture: the file its probe saves, where a case holds what Excel read after its own save, and its member's place.
AUDITS = {'format_properties': ('formats.xlsx', 5, 2), 'member_audit': ('audit.xlsx', 5, 2),
          'method_audit': ('methods.xlsx', 4, 1), 'range_values': ('values.xlsx', 5, 0), 'range_more': ('more.xlsx', 5, 0),
          'range_more_blocks': ('more.xlsx', 5, 0)}


def _probe(source: str, count: int) -> str:
    """The probe's helpers and reads, with a Probe opening the file and reading every case."""
    first = source.index('Public Function S0()')
    head = source[:first]
    # R functions, where a probe reads through one per case, sit after S0's start in the source's order.
    reads = ''.join(found.group(0) for found in re.finditer(r'Private Function R\d+\(ws As Object\) As String\n.*?End Function\n',
                                                            source[first:], re.S))
    again = ''.join(re.search(rf'Public Function A{number}\(\) As String\n.*?End Function\n', source, re.S).group(0)  # type: ignore[union-attr]
                    for number in range(count))
    parts = [list(range(at, min(at + 200, count))) for at in range(0, count, 200)]
    runner = ''
    for number, part in enumerate(parts):
        lines = ''.join(f'    out = out & "@@" & A{case}()\n' for case in part)
        runner += f'\nPrivate Function G{number}() As String\n    Dim out As String\n{lines}    G{number} = Mid(out, 3)\nEnd Function\n'
    joined = ' & "@@" & '.join(f'G{number}()' for number in range(len(parts)))
    return (head + reads + again + runner + '\nPublic Function Probe(ByVal path As String) As String\n'
            f'    Set wb = Workbooks.Open(path)\n    Probe = {joined}\n    wb.Close False\nEnd Function\n')


def check(name: str, excel: ExcelSession) -> None:
    filename, after, member = AUDITS[name]
    record = json.loads((FIXTURES / f'{name}.json').read_text(encoding='utf-8'))
    folder = tempfile.mkdtemp()
    try:
        app = ExcelApplication(sandbox=False)
        app.add_workbook()
        app.add_module(record['source'].replace('folder & Application.PathSeparator', f'folder & "{os.sep}"'), name='Saved')
        app.run('Saved.Setup')
        gaps = set()
        for number in range(len(record['cases'])):
            try:
                app.run(f'Saved.S{number}')
            except VBAUnsupportedError:
                gaps.add(number)
        app.run('Saved.Reopen', folder)
        result = excel.run_vba(_probe(record['source'], len(record['cases'])), 'Probe', (str(Path(folder) / filename),), timeout=400.0)
        if not result.ok:
            print(name, 'REFUSED OR FAILED:', result.outcome, result.error, result.message[:120])
            return
        answers = re.split(r'@@(?=\|)', str(result.value))
        differ: collections.Counter[str] = collections.Counter()
        for number, case in enumerate(record['cases']):
            if number not in gaps and str(number) not in record.get('early', {}) and answers[number] != case[after]:
                differ[f'{case[0]}.{case[member]}'] += 1
        print(name, 'compared', len(record['cases']) - len(gaps) - len(record.get('early', {})), 'differ', dict(differ))
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def main() -> None:
    with ExcelSession(HarnessConfig(lock_wait_s=600.0, exclusive=False)) as excel:
        excel.new_document()
        for name in sys.argv[1:] or AUDITS:
            check(name, excel)


if __name__ == '__main__':
    main()
