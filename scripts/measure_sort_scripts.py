"""Whether Excel sorts, filters, removes duplicates and keeps names in Hebrew, Arabic, Thai and Hangul by the Windows
word-sort weights scripts/bake_sort_weights.py bakes for them (src/pyopenvba/apps/excel/_sort_weights_scripts.py).
The rule written before this probe:

W. Text in these scripts orders and compares by its word-sort key under en-US, as Latin does: Range.Sort orders it so,
   ignoring case or matching it; RemoveDuplicates keeps the first of the text with one key; a filter's = keeps the
   text with the criterion's key and its > the text whose key is greater; Names are listed in key order, and
   Names(...) finds a name by text with its key.

Its prediction is the model's answer, written beside each case before Excel answers. Strings come from
tests/fixtures/sort_keys_scripts.json and from characters the baked table weighs alike; VBA builds each with ChrW, so
the module stays ASCII. Written to tests/fixtures/sort_scripts.json, live Excel beside the model.

    python scripts/measure_sort_scripts.py
"""
from __future__ import annotations

import json
import os
import random
import subprocess
import sys
import tempfile
import time
from collections import defaultdict
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))
OUT = ROOT / 'tests' / 'fixtures' / 'sort_scripts.json'
KEYS = ROOT / 'tests' / 'fixtures' / 'sort_keys_scripts.json'
SEED = 20261107


def vba(text: str) -> str:
    """A VBA expression for ``text`` that keeps the module ASCII."""
    parts = [f'ChrW({ord(character)})' if ord(character) > 126 or character == '"' else f'"{character}"' for character in text]
    return ' & '.join(parts) if parts else '""'


def _orderable(text: str) -> bool:
    from pyopenvba.apps.excel._sort import ordering_key
    from pyopenvba.exceptions import VBAUnsupportedError

    try:
        ordering_key(text, False)
        ordering_key(text, True)
    except VBAUnsupportedError:
        return False
    return text[:1] not in "'=+-@ " and text.strip() == text


def _alike() -> list[list[str]]:
    """Groups of new-script characters, two or more each, the baked table gives one key, case ignored."""
    from pyopenvba.apps.excel._sort import text_key
    from pyopenvba.apps.excel._sort_weights_scripts import SCRIPTS

    groups: dict[bytes, list[str]] = defaultdict(list)
    for code in SCRIPTS:
        character = chr(code)
        if _orderable('a' + character):
            groups[text_key('a' + character, False)].append(character)
    return [group for group in groups.values() if len(group) > 1]


def cases() -> list[dict[str, Any]]:
    pick = random.Random(SEED)
    texts = [case[0] for case in json.loads(KEYS.read_text(encoding='utf-8'))['cases'] if _orderable(case[0])]
    found: list[dict[str, Any]] = []
    for group in range(4):
        chosen = pick.sample(texts, 40)
        for matching in (False, True):
            found.append({'kind': 'sort', 'group': group, 'matching': matching, 'texts': chosen})
    alike = _alike()
    for group in range(4):
        rows: list[str] = []
        for _ in range(10):
            letters = pick.choice(alike)
            stem = pick.choice(texts)[:3]
            rows += [stem + pick.choice(letters) for _ in range(pick.randint(1, 3))]
        found.append({'kind': 'duplicates', 'group': group, 'texts': rows})
    for group in range(4):
        letters = pick.choice(alike)
        stem = pick.choice(texts)[:3]
        rows = [stem + letter for letter in letters] + pick.sample(texts, 12)
        for operator in ('=', '>'):
            found.append({'kind': 'filter', 'group': group, 'operator': operator, 'criterion': stem + letters[0],
                          'texts': rows})
    from pyopenvba.apps.excel._names import valid

    names = [text for text in texts if valid(text) and len(text) > 1]
    for group in range(3):
        chosen = pick.sample(names, 15)
        found.append({'kind': 'names', 'group': group, 'texts': chosen,
                      'lookups': [name[:-1] + other for name in chosen[:5] for letters in alike[:400]
                                  if name[-1] in letters for other in letters if other != name[-1]][:6]})
    return found


def source(case: dict[str, Any]) -> str:
    texts = case['texts']
    lines = ['Public Function One() As String', '    Dim wb As Object, ws As Object, s As String, i As Long, n As Object',
             '    Application.DisplayAlerts = False', '    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)']
    if case['kind'] == 'names':
        for at, text in enumerate(texts):
            lines.append(f'    wb.Names.Add Name:={vba(text)}, RefersTo:="={at + 1}"')
        lines += ['    For i = 1 To wb.Names.Count', '        s = s & Mid(wb.Names(i).RefersTo, 2) & ","', '    Next', '    s = s & "|"',
                  '    On Error Resume Next']
        for text in case['lookups']:
            lines += ['    Set n = Nothing: Err.Clear', f'    Set n = wb.Names({vba(text)})',
                      '    If n Is Nothing Then s = s & "E" & Err.Number & "," Else s = s & Mid(n.RefersTo, 2) & ","']
        lines.append('    On Error GoTo 0')
    else:
        for at, text in enumerate(texts):
            lines.append(f'    ws.Cells({at + 1}, 1).Value = {vba(text)}: ws.Cells({at + 1}, 2).Value = {at + 1}')
        last = len(texts)
        if case['kind'] == 'sort':
            lines += [f'    ws.Range("A1:B{last}").Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlNo, '
                      f'MatchCase:={case["matching"]}',
                      f'    For i = 1 To {last}: s = s & ws.Cells(i, 2).Value & ",": Next']
        elif case['kind'] == 'duplicates':
            lines += [f'    ws.Range("A1:B{last}").RemoveDuplicates Columns:=1, Header:=xlNo',
                      f'    For i = 1 To {last}: s = s & ws.Cells(i, 2).Value & ",": Next']
        else:
            lines += ['    ws.Rows(1).Insert: ws.Range("A1:B1").Value = Array("t", "n")',
                      f'    ws.Range("A1:B{last + 1}").AutoFilter 1, "{case["operator"]}" & {vba(case["criterion"])}',
                      f'    For i = 2 To {last + 1}: If Not ws.Rows(i).Hidden Then s = s & ws.Cells(i, 2).Value & ","',
                      '    Next']
    lines += ['    wb.Close False', '    One = s', 'End Function']
    return '\n'.join(lines) + '\n'


def model(case: dict[str, Any]) -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case), name='Scripts')
    try:
        return str(app.run('Scripts.One'))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:160]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    found = cases()
    for case in found:
        case['rule'] = model(case)
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True):
        lock = _lock()
        while True:
            try:
                os.mkdir(lock)
                break
            except FileExistsError:
                time.sleep(10)
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in found:
                    done = excel.run_vba(source(case), 'One', (), timeout=60.0, line_numbers=False)
                    case['excel'] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            os.rmdir(lock)
    OUT.write_text(json.dumps({'cases': found}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    print('rule W agrees on', sum(case['excel'] == case['rule'] for case in found), 'of', len(found))
    for case in found:
        if case['excel'] != case['rule']:
            print(case['kind'], case['group'], case.get('matching', case.get('operator', '')), '| Excel', case['excel'][:200],
                  '| W', case['rule'][:200])


if __name__ == '__main__':
    main()
