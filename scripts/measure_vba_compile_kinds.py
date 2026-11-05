"""What VBA's compiler refuses on a member call bound early, for each form of member the type library has.

Each line sits under `If False` in a procedure that is run, so that only the compile answers. The lines are chosen from the
table: for each class, each kind of return with and without optional arguments, so that rival rules give different answers.

    python scripts/measure_vba_compile_kinds.py

writes tests/fixtures/vba_compile_kinds.json, keeps what it holds and asks only the rest, so run it again after a time limit.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

from pyopenvba.interpreter._type_library import LibraryClass, find_class  # noqa: E402

OUT = ROOT / 'tests/fixtures/vba_compile_kinds.json'
VARIABLES = {'r': 'Range', 'ws': 'Worksheet', 'wb': 'Workbook', 'f': 'Font', 'sh': 'Shape', 'lo': 'ListObject', 'a': 'Application', 'ch': 'Chart', 'ss': 'Sheets', 'n': 'Name'}
#: Measured after the rule was written, every sixth line, as a check on it.
HELD_OUT = {'se': 'Series', 'wi': 'Window', 'pt': 'PivotTable', 'cm': 'Comment', 'hl': 'Hyperlink', 'lc': 'ListColumn', 'co': 'ChartObject', 'it': 'Interior', 'ax': 'Axis', 'va': 'Validation', 'ps': 'PageSetup', 'wbs': 'Workbooks'}
BODY = 'Function Probe() As String\n    Dim r As Range, ws As Worksheet, wb As Workbook, f As Font, sh As Shape, lo As ListObject, a As Application, ch As Chart, ss As Sheets, n As Name, se As Series, wi As Window, pt As PivotTable, cm As Comment, hl As Hyperlink, lc As ListColumn, co As ChartObject, it As Interior, ax As Axis, va As Validation, ps As PageSetup, wbs As Workbooks, v As Variant, o As Object\n    If False Then\n        @@LINE@@\n    End If\n    Probe = "compiled"\nEnd Function\n'
FILLERS = {'String': '"x"', 'Variant': '1', 'Long': '1', 'Double': '1', 'Boolean': 'True', 'Object': 'o', 'Integer': '1'}


#: Arguments on a function that takes none, by what it returns.
NO_PARAMETER_FUNCTIONS = ['v = wb.pivotcaches(1)', 'v = sh.duplicate(1)', 'v = a.nextletter(1)', 'v = wb.newwindow(1)', 'v = wb.getworkflowtasks(1)', 'v = r.calculate(1)', 'v = r.select(1)', 'v = r.clear(1)', 'v = wb.cancheckin(1)', 'v = a.findfile(1)', 'v = wb.exclusiveaccess(1)', 'v = wb.pivotcaches(Nope:=1)', 'v = r.select(Nope:=1)', 'r.select 1', 'r.select(1)']


def category(returned: str) -> str:
    if returned in ('', 'Variant', 'Object', 'String', 'Long', 'Boolean', 'Double', 'Integer', 'Single', 'Date', 'Currency', 'Byte', 'LongLong'):
        return returned or 'nothing'
    return 'class'


#: No live probe calls or reads these, even under If False: they could send real mail.
NEVER_ASKED = frozenset({'sendmail', 'sendforreview', 'replywithchanges', 'endreview', 'route', 'routeworkbook', 'routingslip', 'hasroutingslip', 'hasmailer', 'mailer',
                         'mailenvelope', 'sendfaxoverinternet', 'maillogon', 'maillogoff', 'mailsession', 'mailsystem', 'followhyperlink', 'follow',
                         'printout', 'printpreview'})


def forms(found: LibraryClass) -> dict[str, list[list[object]]]:
    return {name: entries for name, entries in sorted(found.members.items()) if not name.startswith('_') and name not in NEVER_ASKED}


def lines_for(variable: str, found: LibraryClass) -> dict[str, str]:
    lines: dict[str, str] = {}
    seen: dict[tuple[str, ...], int] = {}

    def take(family: str, shape: tuple[str, ...], line: str, limit: int = 2) -> None:
        key = (family, *shape)
        if seen.get(key, 0) < limit:
            seen[key] = seen.get(key, 0) + 1
            lines[f'{family} {line}'] = line

    for name, entries in forms(found).items():
        kinds = {entry[0] for entry in entries}
        reading = next((entry for entry in entries if entry[0] in (1, 2)), None)
        if reading is None:
            continue
        returned, required, count, varargs = category(str(reading[3])), int(str(reading[1])), len(reading[4]), bool(reading[2])  # type: ignore[arg-type]
        optional = 'optional' if count > required else 'none'
        if 2 in kinds and not kinds & {4, 8} and required == 0:
            take('read_only', (returned, optional), f'{variable}.{name} = 1')
        if 1 in kinds:
            take('assign_method', (returned, optional), f'{variable}.{name} = 1', 1)
        if 2 in kinds and count == 0:
            take('args_on_no_args', (returned,), f'v = {variable}.{name}(1)')
        if count and not varargs:
            take('too_many', (returned, 'method' if 1 in kinds else 'get'), f'v = {variable}.{name}({", ".join(["1"] * (count + 1))})')
        if required:
            take('too_few', (returned, 'method' if 1 in kinds else 'get'), f'v = {variable}.{name}', 1)
        if not returned or returned == 'nothing':
            take('nothing_as_value', ('x',), f'v = {variable}.{name}', 2)
        take('with_result', (returned,), f'With {variable}.{name}{"(" + ", ".join(["1"] * required) + ")" if required else ""}\nEnd With', 1)
        take('bad_name', ('get' if 2 in kinds else 'method',), f'v = {variable}.{name}(Nope:=1)', 1)
    return lines


def asked() -> dict[str, str]:
    lines: dict[str, str] = {}
    for variable, name in VARIABLES.items():
        found = find_class('excel', name)
        assert found is not None, name
        lines.update({f'{variable}: {key}': line for key, line in lines_for(variable, found).items()})
    lines.update({f"{line.split('.')[0].split()[-1]}: args_on_function {line}": line for line in NO_PARAMETER_FUNCTIONS})
    for variable, name in HELD_OUT.items():
        found = find_class('excel', name)
        assert found is not None, name
        lines.update({f'{variable}: {key}': line for key, line in list(lines_for(variable, found).items())[::6]})
    return lines


def save(answers: dict[str, object]) -> None:
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'answers': answers}, indent=0) + '\n', encoding='utf-8')


def main() -> None:
    answers: dict[str, object] = json.loads(OUT.read_text(encoding='utf-8'))['answers'] if OUT.exists() else {}
    excel: ExcelSession | None = None
    wanted = asked()
    print(len(wanted), 'lines,', len(answers), 'held', flush=True)
    try:
        for key, line in wanted.items():
            if key in answers:
                continue
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=1200.0))
                excel.__enter__()
                excel.new_document()
            result = excel.run_vba(BODY.replace('@@LINE@@', line), 'Probe', timeout=30.0, line_numbers=False)
            refused = [dialog for dialog in result.dialogs if dialog.classification == 'compile-error']
            answers[key] = refused[0].message.replace('Compile error:\n\n', '')[:70] if refused else ('compiled' if result.ok else 'stopped: ' + str(result.message)[:50])
            if not result.ok:
                excel.__exit__(None, None, None)
                excel = None
            if len(answers) % 10 == 0:
                save(answers)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
        save(answers)
        print(len(answers), 'of', len(wanted), flush=True)


if __name__ == '__main__':
    main()
