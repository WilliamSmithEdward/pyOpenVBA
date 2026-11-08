"""ThemeColorScheme's default member in live Excel, on a new workbook, typed and late: an index straight on the scheme, as
`wb.Theme.ThemeColorScheme(5).RGB` writes it, and what the same index answers through Colors.

    python scripts/measure_theme_scheme_default.py

writes tests/fixtures/theme_scheme_default.json.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/theme_scheme_default.json'
CASES = ['s(5).RGB', 'TypeName(s(5))', 's(5)', 's.Colors(5).RGB', 's(1).RGB', 's(12).RGB', 's(0).RGB', 's(13).RGB', 's("x").RGB', 's("5").RGB',
         's(5.6).RGB', 'TypeName(s)']
SEPARATOR = '%%'


def macro(declared: str) -> str:
    body = ''.join(f'''    Err.Clear: v = Empty
    v = {case}
    out = out & "{SEPARATOR}" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
''' for case in CASES)
    return f'''Function Probe() As String
    Dim wb As Workbook, s As {declared}, v As Variant, out As String
    Set wb = Workbooks.Add
    Set s = wb.Theme.ThemeColorScheme
    On Error Resume Next
{body}    On Error GoTo 0
    wb.Close SaveChanges:=False
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    answers: dict[str, list[str]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
        for declared in ('ThemeColorScheme', 'Object'):
            result = excel.run_vba(macro(declared), 'Probe', (), timeout=120, line_numbers=False)
            assert result.ok, (result.error, [d.message for d in result.dialogs])
            answers[declared] = str(result.value).split(SEPARATOR)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'cases': CASES, 'answers': answers}, indent=1) + '\n',
                   encoding='utf-8', newline='\n')
    for case, typed, late in zip(CASES, answers['ThemeColorScheme'], answers['Object']):
        print(case.ljust(18), typed, '' if typed == late else f'LATE {late}')


if __name__ == '__main__':
    main()
