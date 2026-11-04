"""When Excel's compiler refuses, and what a macro or a caller sees of it.

A module holds a good procedure and one with a line that does not compile, in the same module or another, and the good one
is run. Compile On Demand is the editor's setting (Tools, Options, General), held in the registry and read here, not changed.
Each case states what each rule predicts, so the answer splits them:

    procedure  refused only when the bad procedure itself is entered
    module     refused when anything in the bad procedure's module runs
    project    refused whenever anything in the project runs

    python scripts/measure_vba_compile_timing.py

writes tests/fixtures/vba_compile_timing.json.
"""
from __future__ import annotations

import json
import re
import winreg
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_compile_timing.json'
#: Each family of refusal the binding work would add, and two the engine has long had, as the line that is refused.
BAD_LINES = {
    'read_only': 'ws.Index = 1',
    'argument_count': 'v = r.Offset(1, 2, 3)',
    'named_argument': 'v = r.Offset(Foo:=1)',
    'sub_as_value': 'v = ws.Activate',
    'unknown_member': 'v = ws.Nope',
    'undeclared_explicit': 'v = NotDeclared',
    'project_function_count': 'v = Good(1, 2)',
}
HEAD = 'Option Explicit\n'
DECLARE = '    Dim ws As Worksheet, r As Range, v As Variant\n    Set ws = ActiveSheet: Set r = ws.Range("A1")\n'
GOOD = 'Function Good() As String\n    Good = "ran"\nEnd Function\n'


def bad(line: str, name: str = 'Bad') -> str:
    return f'Function {name}() As String\n{DECLARE}    {line}\n    {name} = "bad ran"\nEnd Function\n'


def main() -> None:
    answers: dict[str, object] = {}
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\VBA\7.1\Common') as key:
        # Not written while it stands at its default, which is on.
        for name in ('CompileOnDemand', 'BackGroundCompile'):
            try:
                answers[name] = winreg.QueryValueEx(key, name)[0]
            except OSError:
                answers[name] = None

    def ask(label: str, entry: str, modules: list[tuple[str, str]]) -> None:
        # The first module is the harness's own: it is the one run_vba adds. The others are added first.
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            for name, source in modules[1:]:
                excel.add_module(name, source)
            result = excel.run_vba(modules[0][1], entry, timeout=30.0, line_numbers=False)
        dialogs = [{'kind': dialog.classification, 'text': dialog.message[:90]} for dialog in result.dialogs]
        found = re.search(r'VBA error (-?\d+)', str(result.error))
        answers[label] = {'ok': result.ok, 'value': result.value if result.ok else None, 'dialogs': dialogs,
                          'error': found.group(1) if found else (str(result.error)[:120] if result.error else None), 'message': str(result.message)[:120] if result.message else None}
        print(label, answers[label], flush=True)

    # Rule tests, with one refusal family: procedure and module and project predict differently across these.
    line = BAD_LINES['read_only']
    ask('same_module_other_procedure_run', 'Good', [('Probe', HEAD + GOOD + bad(line))])
    ask('bad_procedure_run', 'Bad', [('Probe', HEAD + GOOD + bad(line))])
    ask('other_module_not_called', 'Good', [('Probe', HEAD + GOOD), ('Other', HEAD + bad(line))])
    ask('other_module_called', 'CallsBad', [('Probe', HEAD + GOOD + 'Function CallsBad() As String\n    CallsBad = Bad()\nEnd Function\n'), ('Other', HEAD + bad(line))])
    ask('reached_after_work', 'Stages', [('Probe', HEAD + 'Function Stages() As String\n    ActiveSheet.Range("C3").Value = "before"\n    Stages = Bad()\nEnd Function\n' + bad(line))])
    # On Error in the caller, the bad procedure entered from it.
    ask('on_error_in_caller', 'Guarded', [('Probe', HEAD + 'Function Guarded() As String\n    On Error Resume Next\n    Guarded = Bad()\n    Guarded = "went on " & Err.Number\nEnd Function\n' + bad(line))])
    ask('on_error_other_module', 'Guarded', [('Probe', HEAD + 'Function Guarded() As String\n    On Error Resume Next\n    Guarded = Bad()\n    Guarded = "went on " & Err.Number\nEnd Function\n'), ('Other', HEAD + bad(line))])
    # Each family against the same split: is it refused when only its neighbour runs, and when it is run.
    for family, text in BAD_LINES.items():
        extra = 'Function Good() As String\n    Good = "ran"\nEnd Function\n'
        ask(f'{family} neighbour_run', 'Good', [('Probe', HEAD + extra + bad(text))])
        ask(f'{family} itself_run', 'Bad', [('Probe', HEAD + extra + bad(text))])
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
