"""Every character of the Basic Multilingual Plane (surrogates aside), tried in a name by Names.Add on live Excel and the model:
after a first part that can make no address ("q_" & c), and first ("" & c & "q"). The characters either refuses are written to
tests/fixtures/name_characters_bmp.json with the rule each side follows.

Rules written before the sweep: tests/fixtures/name_later_characters.json had Excel take every character from U+0080 up except
U+0085, U+00A0, U+00AB and U+00A9 after a first letter, and the model only letters, digits and Latin accents. Prediction for
the sweep: the refusals beyond ASCII are a short list, all in U+0080-U+00FF, and the first character refuses the same ones.

    python scripts/measure_name_characters_bmp.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'name_characters_bmp.json'
SPANS = [(0x80, 0xD7FF), (0xE000, 0xFFFF), (0xD800, 0xDFFF)]
STEP = 4096

SOURCE = '''
Public Function Sweep(ByVal lo As Long, ByVal hi As Long) As String
    Dim wb As Object, p As Long, s As String, t As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error Resume Next
    For p = lo To hi
        wb.Names.Add "q_" & ChrW(p), "=1"
        If Err.Number <> 0 Then
            s = s & " " & p
            Err.Clear
        Else
            wb.Names("q_" & ChrW(p)).Delete
            Err.Clear
        End If
        wb.Names.Add ChrW(p) & "q", "=1"
        If Err.Number <> 0 Then
            t = t & " " & p
            Err.Clear
        Else
            wb.Names(ChrW(p) & "q").Delete
            Err.Clear
        End If
    Next
    On Error GoTo 0
    wb.Close False
    Sweep = Mid(s, 2) & "|" & Mid(t, 2)
End Function
'''


def ranges() -> list[tuple[int, int]]:
    return [(lo, min(lo + STEP - 1, end)) for start, end in SPANS for lo in range(start, end + 1, STEP)]


def _points(text: str) -> list[int]:
    return [int(part) for part in text.split()]


def model_refusals() -> tuple[list[int], list[int]]:
    from pyopenvba.apps.excel._names import valid

    later = [p for start, end in SPANS for p in range(start, end + 1) if not valid('q_' + chr(p))]
    first = [p for start, end in SPANS for p in range(start, end + 1) if not valid(chr(p) + 'q')]
    return later, first


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'done': [], 'later': [], 'first': []}
    lock = _lock()
    for lo, hi in ranges():
        if [lo, hi] in record['done']:
            continue
        while True:
            try:
                os.mkdir(lock)
                break
            except FileExistsError:
                time.sleep(10)
        try:
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True), ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                done = excel.run_vba(SOURCE, 'Sweep', (lo, hi), timeout=500.0, line_numbers=False)
        finally:
            os.rmdir(lock)
        if not done.ok:
            print('no answer for', hex(lo), hex(hi), str(done)[:300])
            return
        later, first = str(done.value).split('|')
        record['later'] += _points(later)
        record['first'] += _points(first)
        record['done'].append([lo, hi])
        OUT.write_text(json.dumps(record) + '\n', encoding='utf-8')
        print('swept', hex(lo), hex(hi), flush=True)
    model_later, model_first = model_refusals()
    record['model_later'], record['model_first'] = model_later, model_first
    OUT.write_text(json.dumps(record) + '\n', encoding='utf-8')
    print('Excel refuses after the first', len(record['later']), 'and first', len(record['first']))
    print('model refuses after the first', len(model_later), 'and first', len(model_first))
    print('later, Excel only:', [hex(p) for p in sorted(set(record['later']) - set(model_later))][:60])
    print('later, model only:', len(set(model_later) - set(record['later'])))
    print('first, Excel only:', [hex(p) for p in sorted(set(record['first']) - set(model_first))][:60])
    print('first, model only:', [hex(p) for p in sorted(set(model_first) - set(record['first']))][:60])


if __name__ == '__main__':
    main()
