"""What VBA's compiler says to a variable passed ByRef to a parameter declared as an object type, or passed an object variable.

Each case declares a variable as one type and passes it, under `If False`, to a procedure whose parameter is another type, in one of
several call forms. Only the compile answers: the project is compiled, and either compiles or shows VBA's compile error. The model's
answers are written first.

    python scripts/measure_byref_objects.py predict
    python scripts/measure_byref_objects.py measure

writes tests/fixtures/vba_byref_objects.json, which tests/test_vba_byref_objects.py replays.
"""
from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/vba_byref_objects.json'
#: A class of the project's own, so a user class stands beside the library's.
THING = 'Public Value As Long\n'
ARGUMENTS = ['Variant', 'Object', 'Range', 'Worksheet', 'Collection', 'Thing', 'Long', 'String']
PARAMETERS = ['Variant', 'Object', 'Range', 'Worksheet', 'Collection', 'Thing', 'Long']
#: The call forms, each on a Variant and an object variable passed to Object and Range parameters.
FORMS = {'call': 'Call Take(x)', 'parenthesised': 'Take (x)', 'named': 'Take p:=x', 'function': 'y = Give(x)', 'byval': 'TakeByVal x',
         'array element': 'Take xs(0)'}


def source(argument: str, parameter: str, form: str) -> str:
    element = 'Dim xs(1) As ' + argument + '\n    ' if form == 'array element' else ''
    return (f'Sub Take(p As {parameter})\nEnd Sub\nFunction Give(p As {parameter}) As Variant\nEnd Function\n'
            f'Sub TakeByVal(ByVal p As {parameter})\nEnd Sub\n'
            f'Function Probe() As String\n    Dim x As {argument}, y As Variant\n    {element}If False Then\n        {FORMS[form] if form != "plain" else "Take x"}\n'
            f'    End If\n    Probe = "compiled"\nEnd Function\n')


def cases() -> list[dict[str, Any]]:
    every = [{'argument': argument, 'parameter': parameter, 'form': 'plain'} for argument in ARGUMENTS for parameter in PARAMETERS]
    every += [{'argument': argument, 'parameter': parameter, 'form': form} for form in FORMS for argument in ('Variant', 'Object', 'Range')
              for parameter in ('Object', 'Range')]
    for case in every:
        case['source'] = source(case['argument'], case['parameter'], case['form'])
    return every


def model(case: dict[str, Any]) -> str:
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError

    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(THING, name='Thing', kind='class')
        app.add_module(case['source'], name='Harness')
        return str(app.run('Harness.Probe'))
    except VBACompileError as caught:
        return 'compile error: ' + caught.message


def predict() -> None:
    every = cases()
    record = {'cases': every, 'model': [model(case) for case in every]}
    text = json.dumps(record, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def measure() -> None:
    from pyvbaharness import ExcelSession

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    answers: list[str] = []
    with ExcelSession() as excel:
        for case in record['cases']:
            excel.new_document()
            excel.add_module('Thing', THING, kind='class', line_numbers=False)
            excel.add_module('Harness', case['source'], line_numbers=False)
            compiled = excel.compile_project(watch_seconds=20)
            assert compiled.outcome in ('accepted', 'rejected'), (case, compiled)
            if compiled.ok:
                answers.append(str(excel.run_raw('Harness.Probe', timeout=20)))
            else:
                assert compiled.dialog is not None
                answers.append('compile error: ' + compiled.dialog.message.replace('Compile error:', '').strip())
            print(len(answers), case['argument'], case['parameter'], case['form'], answers[-1], flush=True)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    print('model wrong', sum(mine != theirs for mine, theirs in zip(record['model'], answers)), 'of', len(answers))


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
