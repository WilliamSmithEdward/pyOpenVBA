"""Each worksheet function's parameters the type library declares other than Variant, given cells, several cells, an array,
Null and text, on a WorksheetFunction typed and on one held As Object. The rest of a call is one Excel answered in
tests/fixtures/worksheet_functions.json or worksheet_functions_more.json, whose sheet it uses.

    python scripts/measure_worksheet_function_types.py

writes tests/fixtures/worksheet_function_types.json, which tests/test_excel_worksheet_function_types.py replays both ways.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from pyvbaharness import ExcelSession, HarnessConfig  # noqa: E402

from pyopenvba.interpreter._type_library import find_class  # noqa: E402

OUT = ROOT / 'tests/fixtures/worksheet_function_types.json'
FIXTURES = ROOT / 'tests/fixtures'
#: Each value held in a Variant, and the same range written in the call itself, which VBA treats as an object it has.
HELD = ['Set x = ws.Range("A2")', 'Set x = ws.Range("B2")', 'Set x = ws.Range("A1:A3")', 'x = Array(1, 2)', 'x = Null', 'x = "x"', 'x = "2"', 'x = True']
DIRECT = ['ws.Range("A2")', 'ws.Range("A1:A3")']
HELPER = '''Private Function Note(e As Long, v As Variant) As String
    If e <> 0 Then
        Note = "e" & e
    ElseIf IsObject(v) Then
        Note = "Object " & TypeName(v)
    ElseIf IsNull(v) Then
        Note = "Null"
    ElseIf IsEmpty(v) Then
        Note = "Empty"
    ElseIf IsArray(v) Then
        Note = TypeName(v)
    ElseIf IsError(v) Then
        Note = CStr(v)
    ElseIf Len(CStr(v)) > 40 Then
        Note = TypeName(v) & " " & Len(CStr(v))
    Else
        Note = TypeName(v) & " " & CStr(v)
    End If
End Function
'''


def split_arguments(text: str) -> list[str]:
    """A call's arguments, split at the commas outside quotes and brackets."""
    parts: list[str] = []
    depth, quoted, start = 0, False, 0
    for at, char in enumerate(text):
        if char == '"':
            quoted = not quoted
        elif not quoted and char in '({':
            depth += 1
        elif not quoted and char in ')}':
            depth -= 1
        elif not quoted and not depth and char == ',':
            parts.append(text[start:at].strip())
            start = at + 1
    parts.append(text[start:].strip())
    return parts


def cases() -> list[tuple[str, str, str]]:
    """Each (function, the parameter's declared type, the statements of the call)."""
    library = find_class('excel', 'worksheetfunction')
    assert library is not None
    calls: dict[str, list[list[str]]] = {}
    for name in ('worksheet_functions.json', 'worksheet_functions_more.json'):
        for case in json.loads((FIXTURES / name).read_text(encoding='utf-8'))['cases']:
            calls.setdefault(case['function'], []).append(split_arguments(case['arguments']))
    found: list[tuple[str, str, str]] = []
    for name in sorted(calls):
        forms = library.members.get(name.lower())
        if not forms:
            continue
        declared = [one.split(':')[-1] for one in forms[0][4]]
        longest = max(calls[name], key=len)
        for index, kind in enumerate(declared[:len(longest)]):
            if kind == 'Variant':
                continue
            for held in HELD:
                arguments = [*longest[:index], 'x', *longest[index + 1:]]
                found.append((name, kind, f'{held}: v = f.{name}({", ".join(arguments)})'))
            for direct in DIRECT:
                arguments = [*longest[:index], direct, *longest[index + 1:]]
                found.append((name, kind, f'v = f.{name}({", ".join(arguments)})'))
    return found


def source(typed: bool) -> str:
    every = cases()
    fill = json.loads((FIXTURES / 'worksheet_functions.json').read_text(encoding='utf-8'))['setup']
    kind = 'WorksheetFunction' if typed else 'Object'
    body = HELPER + fill.split('Private Function Item')[0]
    parts = [every[at:at + 200] for at in range(0, len(every), 200)]
    for number, part in enumerate(parts):
        lines = [f'Private Function Part{number}(f As {kind}, ws As Object) As String', '    Dim v As Variant, x As Variant, s As String', '    On Error Resume Next']
        lines += [f'    v = Empty: x = Empty: {statement}: s = s & "@@" & Note(Err.Number, v): Err.Clear' for _, _, statement in part]
        lines += [f'    Part{number} = Mid(s, 3)', 'End Function', '']
        body += '\n'.join(lines) + '\n'
    joined = ' & "@@" & '.join(f'Part{number}(f, ws)' for number in range(len(parts)))
    return body + (f'Public Function Probe() As String\n    Dim wb As Object, ws As Object, f As {kind}, s As String\n    Set wb = Workbooks.Add(xlWBATWorksheet)\n'
                   f'    Set ws = wb.Worksheets(1)\n    Fill ws\n    Set f = Application.WorksheetFunction\n    s = {joined}\n    wb.Close False\n    Probe = s\nEnd Function\n')


def main() -> None:
    every = cases()
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
        excel.new_document()
        late = excel.run_vba(source(False), 'Probe', (), timeout=270.0)
        assert late.ok and not late.dialogs, late
        typed = excel.run_vba(source(True), 'Probe', (), timeout=270.0)
        assert typed.ok and not typed.dialogs, typed
    answers, bound = str(late.value).split('@@'), str(typed.value).split('@@')
    assert len(answers) == len(bound) == len(every), (len(answers), len(bound), len(every))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'late_source': source(False), 'typed_source': source(True),
              'cases': [[name, kind, statement, one, two] for (name, kind, statement), one, two in zip(every, answers, bound, strict=True)]}
    OUT.write_text(json.dumps(record, indent=0) + '\n', encoding='utf-8')
    print(len(every), 'cases', sum(case[3] != case[4] for case in record['cases']), 'differ by binding')


if __name__ == '__main__':
    main()
