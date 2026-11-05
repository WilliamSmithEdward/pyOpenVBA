"""What a property that takes nothing answers when it is given an index anyway, bound late, per class and member.

Typed, VBA reads the property and hands the index to what it answered. Late, the object is asked with the argument: some members
refuse with 450, some with 451, some with 438, and some index what they answer, which no declaration says. Each line reads
o.Member(1) under On Error, into a Variant. Only property gets are asked, nothing is set, and no member that shows a window or
touches mail is read.

    python scripts/measure_late_indexes.py

writes tests/fixtures/late_indexes.json, keeps each class as it finishes and asks only the rest. scripts/build_late_lets.py makes
the table from it.
"""
from __future__ import annotations

import ast
import json
from datetime import datetime, timezone
from pathlib import Path

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.interpreter._type_library import find_class

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/late_indexes.json'
TARGETS = {
    'Range': 'ws.Range("A1:B2")', 'Worksheet': 'ws', 'Workbook': 'ActiveWorkbook', 'Application': 'Application', 'Font': 'ws.Range("A1").Font',
    'Interior': 'ws.Range("A1").Interior', 'Worksheets': 'Worksheets', 'Workbooks': 'Workbooks',
}
#: What would show a window or speak, and every member of the mail guard (tests/test_probes_leave_mail_alone.py), read from it so that this file names none.
LEFT_OUT = frozenset({
    'speak', 'speech', 'functionwizard', 'dialogbox', 'checkspelling', 'printpreview', 'showdataform', 'sendkeys', 'inputbox', 'dialogs', 'filedialog', 'getopenfilename',
    'getsaveasfilename', 'export', 'assistant', 'vbe',
})


def mail_members() -> frozenset[str]:
    """What the guard bans, whole words and those it looks for after a dot alike."""
    guard = ast.parse((Path(__file__).resolve().parents[1] / 'tests/test_probes_leave_mail_alone.py').read_text(encoding='utf-8'))
    found: set[str] = set()
    for node in guard.body:
        if isinstance(node, ast.Assign) and isinstance(node.targets[0], ast.Name) and node.targets[0].id.startswith('BANNED'):
            found.update(str(one).lower() for one in ast.literal_eval(node.value))
    return frozenset(found)


HELPERS = '''Function Probe() As String
Dim ws As Object, o As Object, s As String, v As Variant, e As Long
Set ws = Worksheets.Add
ws.Range("A1:B2").Value = 5
Set o = TARGET
On Error Resume Next
'''


def members_of(class_name: str) -> list[str]:
    application = ExcelApplication(sandbox=False)
    library = application.interpreter.host.vba_type_library
    found = find_class(library, class_name)
    assert found is not None
    return sorted(name for name, forms in found.members.items()
                  if name.isidentifier() and not name.startswith('_') and name not in LEFT_OUT and name not in mail_members()
                  and any(form[0] == 2 and not form[4] for form in forms) and not any(form[0] == 1 for form in forms))


def macro(class_name: str, members: list[str]) -> str:
    body = ''.join(f'v = Empty\nv = o.{member}(1)\ne = Err.Number\nErr.Clear\ns = s & e & "|"\n' for member in members)
    return (HELPERS.replace('TARGET', TARGETS[class_name]) + body
            + 'On Error GoTo 0\nApplication.DisplayAlerts = False\nws.Delete\nApplication.DisplayAlerts = True\nProbe = s\nEnd Function\n')


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    record: dict[str, object] = json.loads(OUT.read_text(encoding='utf-8')) if OUT.exists() else {'classes': {}}
    classes: dict[str, dict[str, int]] = record['classes']  # type: ignore[assignment]
    for class_name in TARGETS:
        if class_name in classes:
            continue
        members = members_of(class_name)
        with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
            excel.new_document()
            result = excel.run_vba(macro(class_name, members), 'Probe', timeout=180.0, line_numbers=False)
        if not result.ok or result.dialogs:
            print(class_name, 'stopped', str(result.message)[:100], [dialog.message[:70] for dialog in result.dialogs], flush=True)
            continue
        numbers = str(result.value).split('|')[:-1]
        assert len(numbers) == len(members), (class_name, len(numbers), len(members))
        classes[class_name] = {member: int(number) for member, number in zip(members, numbers)}
        record['measured_at'] = datetime.now(timezone.utc).isoformat()
        OUT.write_text(json.dumps(record, indent=1, sort_keys=True) + '\n', encoding='utf-8', newline='\n')
        counts: dict[int, int] = {}
        for number in classes[class_name].values():
            counts[number] = counts.get(number, 0) + 1
        print(class_name, len(members), counts, flush=True)


if __name__ == '__main__':
    main()
