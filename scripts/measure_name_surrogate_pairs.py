"""Characters beyond the Basic Multilingual Plane in a name. tests/fixtures/name_characters_bmp.json has Excel take every lone high
surrogate unit first and every surrogate unit after the first. The rule written before this probe: Excel reads a name unit by
unit, so a pair is taken wherever its units are, which is anywhere. Each pair is tried after "q_" and first before "q", on live
Excel and the model, and written to tests/fixtures/name_surrogate_pairs.json.

    python scripts/measure_name_surrogate_pairs.py
"""
from __future__ import annotations

import json
import os
import subprocess
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'name_surrogate_pairs.json'
#: Emoji, CJK Extension B, mathematical letters, Gothic, Linear B, a musical symbol, a tag, private use and unassigned planes.
POINTS = [0x1F600, 0x1F44D, 0x20000, 0x2A6D6, 0x1D400, 0x1D7CE, 0x10330, 0x10000, 0x1D11E, 0xE0041, 0xF0000, 0x10FFFD, 0x3FFFF, 0x1FFFE, 0x16FE0]


def _pair(point: int) -> str:
    high, low = 0xD800 + ((point - 0x10000) >> 10), 0xDC00 + ((point - 0x10000) & 0x3FF)
    return f'ChrW({high}) & ChrW({low})'


def source() -> str:
    lines = ['Public Function Probe() As String', '    Dim wb As Object, s As String', '    Application.DisplayAlerts = False',
             '    Set wb = Workbooks.Add(xlWBATWorksheet)', '    On Error Resume Next']
    for point in POINTS:
        lines += [f'    wb.Names.Add "q_" & {_pair(point)}, "=1"', '    s = s & " " & Err.Number: Err.Clear',
                  f'    wb.Names.Add {_pair(point)} & "q", "=1"', '    s = s & " " & Err.Number: Err.Clear']
    lines += ['    s = s & " " & wb.Names.Count', '    On Error GoTo 0', '    wb.Close False', '    Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


def model() -> str:
    from pyopenvba.apps.excel import ExcelApplication

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(), name='Pairs')
    return str(app.run('Pairs.Probe'))


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    modelled = model().split()
    lock = _lock()
    while True:
        try:
            os.mkdir(lock)
            break
        except FileExistsError:
            time.sleep(10)
    try:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True), ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
            excel.new_document()
            done = excel.run_vba(source(), 'Probe', (), timeout=120.0, line_numbers=False)
            assert done.ok, done
            measured = str(done.value).split()
    finally:
        os.rmdir(lock)
    cases = [{'point': f'U+{point:05X}', 'later_excel': int(measured[2 * at]), 'first_excel': int(measured[2 * at + 1]),
              'later_model': int(modelled[2 * at]), 'first_model': int(modelled[2 * at + 1]), 'rule': 0} for at, point in enumerate(POINTS)]
    OUT.write_text(json.dumps({'excel_count': int(measured[-1]), 'model_count': int(modelled[-1]), 'cases': cases}, indent=1) + '\n', encoding='utf-8')
    for case in cases:
        print(case)


if __name__ == '__main__':
    main()
