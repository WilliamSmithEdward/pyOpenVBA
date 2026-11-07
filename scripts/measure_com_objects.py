"""What CreateObject and GetObject answer in live Excel, into an Object and into a Variant, read line by line.

Only names that do not exist and Scripting.Dictionary are created. GetObject attaches only to the Excel the probe runs in and reads whether it is
Application. No other program starts, no file is named and no server is named.

    python scripts/measure_com_objects.py

writes tests/fixtures/com_objects.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/com_objects.json'
SEPARATOR = '|'
LINES: list[tuple[str, str]] = [
    ('Set o = CreateObject("No.Such.Class")', 'TypeName(o)'), ('Set o = CreateObject("")', 'TypeName(o)'), ('Set o = CreateObject("NoDots")', 'TypeName(o)'),
    ('Set o = CreateObject(5)', 'TypeName(o)'), ('Set o = CreateObject(Null)', 'TypeName(o)'), ('Set o = CreateObject(Empty)', 'TypeName(o)'),
    ('Set o = CreateObject("Scripting.Dictionary")', 'TypeName(o)'), ('Set o = CreateObject("SCRIPTING.DICTIONARY")', 'TypeName(o)'),
    ('Set o = CreateObject("Scripting.Dictionary", "")', 'TypeName(o)'), ('Set o = CreateObject(" Scripting.Dictionary")', 'TypeName(o)'),
    ('Set o = CreateObject("Scripting.Dictionary ")', 'TypeName(o)'),
    ('o = CreateObject("Scripting.Dictionary")', 'TypeName(o)'), ('v = CreateObject("Scripting.Dictionary")', 'TypeName(v)'),
    ('n = CreateObject("Scripting.Dictionary")', 'n'), ('Set c = CreateObject("Scripting.Dictionary")', 'TypeName(c)'),
    ('Set o = CreateObject("Scripting.Dictionary"): o.NoSuchMember', 'TypeName(o)'), ('Set o = CreateObject("Scripting.Dictionary"): o.Add 1', 'o.Count'),
    ('Set o = CreateObject("Scripting.Dictionary"): o.Add "k", 1: Set o = Nothing', 'TypeName(o)'),
    ('Set o = GetObject(, "No.Such.Class")', 'TypeName(o)'), ('Set o = GetObject(, "Scripting.Dictionary")', 'TypeName(o)'),
    ('Set o = GetObject("", "Scripting.Dictionary")', 'TypeName(o)'), ('Set o = GetObject("", "No.Such.Class")', 'TypeName(o)'),
    ('Set o = GetObject(, "")', 'TypeName(o)'), ('Set o = GetObject(, "Excel.Application")', 'TypeName(o) & " " & (o Is Application)'),
    ('Set o = GetObject(, "excel.application")', 'TypeName(o) & " " & (o Is Application)'),
]
TEMPLATE = '''Function Probe() As String
    Dim o As {O}, c As Collection, v As Variant, n As Long, e As Long, r As String
{BODY}
    Probe = r
End Function
'''


def macro(declared: str) -> str:
    body = ''.join('    Set o = Nothing: Set c = Nothing: v = Empty: n = 0\n    On Error Resume Next\n    ' + one.replace(': ', '\n    ')
                   + f'\n    e = Err.Number: Err.Clear\n    r = r & e & "=" & {two} & "{SEPARATOR}"\n    On Error GoTo 0\n' for one, two in LINES)
    return TEMPLATE.replace('{O}', declared).replace('{BODY}', body)


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, str] = {}
    for declared in ('Object', 'Variant'):
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(macro(declared), 'Probe', timeout=60.0, line_numbers=False)
        if not result.ok or result.dialogs:
            raise SystemExit(f'{declared} stopped: {result.message}')
        answers[declared] = str(result.value)
        print(declared, 'done', flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8', newline='\n')


if __name__ == '__main__':
    main()
