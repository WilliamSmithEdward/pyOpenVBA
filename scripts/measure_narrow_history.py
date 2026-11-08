"""Whether what a narrow column shows depends on the widths visited before it, into tests/fixtures/narrow_history.json. Needs Excel at 96 DPI, checked at both ends of each call.

Two rival rules, written before Excel was asked. CURRENT: the same column in pixels shows the same whatever came before. HISTORY: it does not. Each sequence runs twice, in a different order the second time, so a difference left over from an earlier call shows too.
"""
from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

FIXTURE = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'narrow_history.json'
SEQUENCES = ('0.09,0.25,0.5,1,1.5,2,3,5', '0.25,0.5,1,1.5,2,3,5', '0.5,1,1.5,2,3,5', '1,1.5,2,3,5', '2,3,5', '3', '1.5', '0.5', '0.25', '0.09', '5,3,2,1.5,1,0.5,0.25,0.09', '0.09,5,0.5', '0.09,0.5', '0.17,0.5', '0.25,0.09,0.5', '0.5,0.09,0.5', '0,0.5', '0,5,0.5')
CELLS = {'typed number': ('A1.Value = 1234567.891', 'General'), 'typed date': ('A1.Value = 45123', 'yyyy-mm-dd'), 'typed logical': ('A1.Value = True', 'General'), 'typed small': ('A1.Value = 0.000123456', 'General'), 'formula': ('A1.Formula = "=1234567.891"', 'General'), 'text with a fill': ('A1.Value = "abc"', '@*-')}
SOURCE = r'''Public Function Go(ByVal widths As String, ByVal font As String, ByVal size As Double, ByVal format As String) As String
    Dim ws As Object, r As String, one As Variant, before As String
    Set ws = ActiveSheet
    before = ws.StandardHeight
    ws.Cells.Clear
    ws.Columns(1).UseStandardWidth = True
    ws.Range("A1").NumberFormat = format
    ws.Range("%SET%
    ws.Range("A1").Font.Name = font
    ws.Range("A1").Font.Size = size
    For Each one In Split(widths, ",")
        ws.Columns(1).ColumnWidth = Val(one)
        r = r & "|" & CLng(ws.Columns(1).Width * 4 / 3) & "=" & ws.Range("A1").Text
    Next
    Go = before & "~" & ws.StandardHeight & r
End Function
'''


def main() -> None:
    cases = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, (assign, shape) in CELLS.items():
            code = SOURCE.replace('%SET%', assign.replace('A1.', 'A1").', 1))
            for font, size in (('Arial', 18.0), ('Calibri', 18.0), ('Arial', 9.0)):
                for repeat, order in enumerate((SEQUENCES, tuple(reversed(SEQUENCES)))):
                    for widths in order:
                        result = excel.run_vba(code, 'Go', (widths, font, size, shape), timeout=60.0, line_numbers=False)
                        assert result.ok and not result.dialogs, (name, widths, result)
                        first, second = str(result.value).split('|', 1)[0].split('~')
                        assert (first, second) == ('15', '15'), f'Excel is not at 96 DPI: {first}, {second}'
                        cases.append({'cell': name, 'font': font, 'size': size, 'pass': repeat, 'widths': widths, 'shown': str(result.value).split('|')[1:]})
    FIXTURE.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': '15', 'source': SOURCE, 'cells': CELLS, 'cases': cases}, indent=0), encoding='utf-8', newline='\n')
    seen: dict[tuple[str, str, float, str], set[str]] = defaultdict(set)
    where: dict[tuple[str, str, float, str, str], list[str]] = defaultdict(list)
    for case in cases:
        for place, one in enumerate(case['shown']):
            pixels, shown = one.split('=', 1)
            seen[(case['cell'], case['font'], case['size'], pixels)].add(shown)
            where[(case['cell'], case['font'], case['size'], pixels, shown)].append(f"{case['widths']}@{place}#{case['pass']}")
    varied = {key: value for key, value in seen.items() if len(value) > 1}
    print(len(seen), 'cell, font and width pairs,', len(varied), 'show more than one thing')
    for key, value in sorted(varied.items(), key=str):
        print(key, {shown: where[(*key, shown)][:4] for shown in value})
    for case in cases:
        if case['cell'] == 'typed number' and case['font'] == 'Arial' and case['size'] == 18.0 and case['pass'] == 0:
            print(case['widths'], case['shown'])


if __name__ == '__main__':
    main()
