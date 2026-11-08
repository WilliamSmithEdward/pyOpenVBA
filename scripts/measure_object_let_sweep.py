"""A sweep of an object given by Let (no Set), member by member, over the classes macros set most: Range, Font, Interior, Worksheet,
Workbook and Application. The first two probes (scripts/measure_late_object_lets.py, measure_late_object_lets_2.py) showed that what a
member makes of an object goes by the member, not its type, so each settable member gets its own two cases here.

Each case makes a new workbook with 7 in A1, gives one object to one member, reads the member back and closes the workbook unsaved. The
objects: an Interior, which has no value a Let can read, and Nothing. Late bound, each member takes both; typed, a Let of Nothing does not
compile, so only the Interior is given.

The members are every property of the class that takes a Let with no arguments, from the type library, less those whose names start
with an underscore, which VBA cannot name, and these:
  on Workbook, those no probe touches (scripts/probe_limits.py) and PythonRuntime, which runs in the cloud;
  on Application, all but the settings scripts/measure_application_settings.py and measure_settable_audit.py already set and put back,
  and of those, the window and the harness's own switches (DisplayFullScreen, Interactive, Caption, MergeInstances, UserControl,
  ScreenUpdating, DisplayAlerts, EnableEvents). Each Application member is read first and put back after its case.

The receiving rule's prediction, from the first two probes, by the member's declared type, is written down before Excel answers, with the
model's own answer for each case run alone:
  late, or typed into a Variant: an Interior is 13 into a String and 1004 into a Boolean or Variant; Nothing is 13 into a String and 5
  into a Boolean or Variant; a Long or Double is left open between them;
  typed into a String, Boolean, Long or Double, the Interior's default value is wanted, and it has none: 438.

    python scripts/measure_object_let_sweep.py predict
    python scripts/measure_object_let_sweep.py measure

writes tests/fixtures/object_let_sweep.json.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/object_let_sweep.json'
SEPARATOR = '%%'
PATHS = {'Range': 'ws.Range("D1")', 'Font': 'ws.Range("D1").Font', 'Interior': 'ws.Range("D1").Interior', 'Worksheet': 'ws',
         'Workbook': 'wb', 'Application': 'Application'}
#: Application members already set and put back by the settings probes, less the window and the harness's own switches.
APPLICATION = frozenset({'displaystatusbar', 'displayformulabar', 'displayscrollbars', 'displaynoteindicator', 'enableanimations',
                         'printcommunication', 'asktoupdatelinks', 'alertbeforeoverwriting', 'celldraganddrop', 'copyobjectswithcells',
                         'editdirectlyincell', 'enableautocomplete', 'fixeddecimal', 'moveafterreturn', 'cursor', 'enablecancelkey',
                         'moveafterreturndirection', 'displaycommentindicator', 'fixeddecimalplaces', 'highqualitymodeforgraphics',
                         'generatetablerefs', 'cursormovement', 'defaultsheetdirection', 'largeoperationcellthousandcount', 'odbctimeout',
                         'calculation', 'statusbar', 'cutcopymode'})
OBJECTS = ['ws.Range("A1").Interior', 'Nothing']
PART = 12


def _limits() -> frozenset[str]:
    spec = importlib.util.spec_from_file_location('probe_limits', ROOT / 'scripts/probe_limits.py')
    assert spec is not None and spec.loader is not None
    limits: Any = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(limits)
    names = limits.MAIL | limits.RIGHTS | limits.USER_LEVEL | limits.NEVER_SET | limits.SERVERS | {'PythonRuntime'}
    return frozenset(name.lower() for name in names)


def targets() -> list[tuple[str, str, str, str]]:
    """Each class's members that take a Let with no arguments: the class, how the probe reaches it, the member, its declared type."""
    sys.path.insert(0, str(ROOT / 'src'))
    from pyopenvba.interpreter._type_library import find_class

    left_out = _limits()
    found: list[tuple[str, str, str, str]] = []
    for name, path in PATHS.items():
        library = find_class('excel', name)
        assert library is not None
        for member, forms in sorted(library.members.items()):
            reading = library.reading(member)
            if reading is None or reading[4] or not any(form[0] == 4 and form[1] <= 1 and not form[4][:-1] for form in forms):
                continue
            # A name with a leading underscore is no VBA identifier: a compile error, which stops the harness.
            if member in left_out or member.startswith('_') or (name == 'Application' and member not in APPLICATION):
                continue
            found.append((name, path, member, reading[3]))
    return found


def rule(kind: str, value: str, typed: bool) -> str:
    """The receiving rule's predicted error for a case, before Excel answers."""
    if typed and value == 'Nothing':
        return 'does not compile'
    if typed and kind != 'Variant':
        return '438'
    if value == 'Nothing':
        return {'String': '13', 'Boolean': '5', 'Variant': '5'}.get(kind, '13 or 5')
    return {'String': '13', 'Boolean': '1004', 'Variant': '1004'}.get(kind, '13 or 1004')


def _case(target: tuple[str, str, str, str], value: str, typed: bool, compiles_only: bool) -> str:
    kind_class, path, prop, _kind = target
    holder = f'{"t" if typed else "o"}{kind_class}'
    if compiles_only and typed and value == 'Nothing':
        return f'    out = out & "{SEPARATOR}" & "does not compile"\n'
    keep = restore = ''
    if kind_class == 'Application':
        keep = f'    Err.Clear: was = Empty: was = Application.{prop}\n'
        restore = f'    Err.Clear: Application.{prop} = was\n'
    return f'''    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = 7
{keep}    Set {holder} = {path}
    Err.Clear
    {holder}.{prop} = {value}
    s = "e" & Err.Number
    Err.Clear: v = Empty
    v = {holder}.{prop}
    s = s & "/" & IIf(Err.Number, "e" & Err.Number, TypeName(v) & ":" & CStr(v))
{restore}    Set {holder} = Nothing
    wb.IsAddin = False
    wb.Close SaveChanges:=False
    out = out & "{SEPARATOR}" & s
'''


def cases() -> list[tuple[str, str, str, str, str]]:
    return [(*target, value) for target in targets() for value in OBJECTS]


def _declare(typed: bool) -> str:
    kind = 'As ' if typed else 'As Object'
    names = ', '.join(f'{"t" if typed else "o"}{name} {kind + name if typed else kind}' for name in PATHS)
    return f'    Dim {names}\n'


def macro(typed: bool, compiles_only: bool = False, chosen: list[tuple[str, str, str, str, str]] | None = None) -> str:
    every = cases() if chosen is None else chosen
    parts = [every[at:at + PART] for at in range(0, len(every), PART)] or [[]]
    procedures = ''.join(f'''Private Sub Part{index}(ByRef out As String)
    Dim wb As Workbook, ws As Worksheet, v As Variant, s As String, was As Variant
{_declare(typed)}    On Error Resume Next
{''.join(_case(case[:4], case[4], typed, compiles_only) for case in part)}End Sub
''' for index, part in enumerate(parts))
    calls = ''.join(f'    Part{index} out\n' for index in range(len(parts)))
    return f'''{procedures}Function Probe() As String
    Dim out As String
    Application.DisplayAlerts = False
{calls}    Application.DisplayAlerts = True
    Probe = Mid(out, {len(SEPARATOR) + 1})
End Function
'''


def alone(case: list[str] | tuple[str, ...], typed: bool) -> str:
    """The model's answer to one case run by itself, so that a refusal answers that case only."""
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBACompileError, VBAUnsupportedError

    target, value = (case[0], case[1], case[2], case[3]), case[4]
    body = _case(target, value, typed, False)
    source = (f'Function Probe() As String\n    Dim out As String, wb As Workbook, ws As Worksheet, v As Variant, s As String, was As Variant\n'
              f'{_declare(typed)}    On Error Resume Next\n{body}    Probe = Mid(out, {len(SEPARATOR) + 1})\nEnd Function\n')
    app = ExcelApplication()
    app.add_workbook()
    try:
        app.add_module(source, name='H')
        return str(app.run('H.Probe'))
    except VBACompileError:
        return 'does not compile'
    except VBAUnsupportedError:
        return 'refused'


def predict() -> None:
    every = cases()
    model = {('typed' if typed else 'late'): [alone(case, typed) for case in every] for typed in (True, False)}
    table = {('typed' if typed else 'late'): [rule(case[3], case[4], typed) for case in every] for typed in (True, False)}
    text = json.dumps({'cases': every, 'rule': table, 'model': model}, indent=1) + '\n'
    OUT.write_text(text, encoding='utf-8', newline='\n')
    print(len(every), 'cases; predictions sha256', hashlib.sha256(text.encode('utf-8')).hexdigest())


def _verdict(predicted: str, answer: str) -> bool:
    if predicted == 'does not compile':
        return answer == predicted
    return answer.split('/')[0].removeprefix('e') in predicted.split(' or ')


def measure() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, Any] = json.loads(OUT.read_text(encoding='utf-8'))
    every = [tuple(case) for case in record['cases']]
    answers: dict[str, list[str]] = {'typed': [], 'late': []}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        excel.run_vba('Function Hello() As String\n    Hello = "x"\nEnd Function\n', 'Hello', (), timeout=60)
        settings = sorted({case[2] for case in every if case[0] == 'Application'})
        reading = ('Function Settings() As String\n    On Error Resume Next\n'
                   + ''.join(f'    Settings = Settings & "|{one}=" & CStr(Application.{one})\n' for one in settings) + 'End Function\n')
        before = excel.run_vba(reading, 'Settings', (), timeout=60).value
        for name in PATHS:
            chosen: list[Any] = [case for case in every if case[0] == name]
            for typed in (True, False):
                result = excel.run_vba(macro(typed, compiles_only=True, chosen=chosen), 'Probe', (), timeout=600, line_numbers=False)
                assert result.ok, (name, typed, result.error, [one.message for one in result.dialogs])
                answers['typed' if typed else 'late'] += str(result.value).split(SEPARATOR)
            print(name, len(chosen), 'cases', flush=True)
        after = excel.run_vba(reading, 'Settings', (), timeout=60).value
    order = {name: index for index, name in enumerate(PATHS)}
    assert [case[0] for case in every] == sorted((case[0] for case in every), key=order.__getitem__)
    record.update({'measured_at': datetime.now(timezone.utc).isoformat(), 'excel': answers})
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8', newline='\n')
    # Excel keeps many of these for the user: any one not put back is to be restored by hand before anything else. StatusBar, which
    # lasts the session only, reads "FALSE" once False is given back to it.
    assert str(before).replace('statusbar=False', 'statusbar=FALSE') == after, (before, after)
    for binding in answers:
        verdicts = [_verdict(p, a) for p, a in zip(record['rule'][binding], answers[binding])]
        misses = sum(a != b and b != 'refused' for a, b in zip(answers[binding], record['model'][binding]))
        refused = record['model'][binding].count('refused')
        print(binding, 'rule held', verdicts.count(True), 'failed', verdicts.count(False), '; model wrong', misses, 'refused', refused)


if __name__ == '__main__':
    {'predict': predict, 'measure': measure}[sys.argv[1]]()
