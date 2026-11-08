"""FormatCondition round 56101's three rules, measured apart: a rule's Font set through the range it was reached by, its
Font.Size, and its NumberFormat after a format Excel refuses. Rules written before this probe:

Reach. B1. Every set takes, whatever range reached the rule (what the model did). B2. Reached through a range that is not the
   rule's own (Cells, a bigger block), Bold and Italic sets change nothing and Color sets take. B3. As B2 for every reach,
   the rule's own range included, on a rule that held no font before.
Size. S1. Refused (what the model did). S2. Taken silently, changing nothing, and read as Null. S3. Error 1004.
Format. N1. A format Excel refuses leaves the one before it (what the model meant). N2. It leaves the rule with none, Empty.

Second look, after round 56102 (a rule reached through D10:D12 once rows went in above it, on D11:D13, set Bold as usual; a
file's rule reached through the sheet's cells kept 0.00 past a refused code). Rules written before it:
Overlap. P1. A range that overlaps the rule's own without holding it or lying in it reaches it as its own cells do. P2. As any
   other range: the object reads what it read before.
Origin. O1. A refused code clears the rule only reached through its own cells; reached otherwise it keeps its code. O2. A
   rule read from a file keeps its code, one made in the session loses it, however reached.

Each case runs on a new workbook whose D1:D5 holds one or two cell-value rules, typed and late bound, reading Err straight after
the statement and the property through the same object and through D1:D5's first rule; by live Excel and the model, written
to tests/fixtures/format_condition_font.json.

    python scripts/measure_format_condition_font.py
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_table_pivot_rounds as rounds  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests' / 'fixtures' / 'format_condition_font.json'
ONE = '    ws.Range("D1:D5").FormatConditions.Add xlCellValue, xlGreater, "=3"\n'
TWO = ONE + '    ws.Range("D1:D5").FormatConditions.Add xlCellValue, xlLess, "=1"\n'
BOLD = '    ws.Range("D1:D5").FormatConditions(1).Font.Bold = True\n'
REACHES = ['ws.Range("D1:D5").FormatConditions(1)', 'ws.Cells.FormatConditions(1)', 'ws.Range("A1:E9").FormatConditions(1)',
           'ws.Range("D2").FormatConditions(1)', 'ws.Range("D1:D5").FormatConditions(2)', 'ws.Cells.FormatConditions(2)']
SETS = [('Font.Bold', 'True'), ('Font.Italic', 'True'), ('Font.Color', '255'), ('Font.Strikethrough', 'True'), ('Font.Underline', 'xlUnderlineStyleSingle'),
        ('Font.Size', '14'), ('Font.Name', '"Arial"'), ('Interior.Color', '65535')]
OWN_FILE = 'ws.Range("D10:D12").FormatConditions(1)'
FORMATS = [('', '"nonsense["'), ('', '"0.00"'), ('"0.00"', '"nonsense["'), ('"0.00"', '""'), ('', '"General"'), ('"0.00"', 'Empty')]


def cases() -> list[dict[str, str]]:
    made = [{'group': 'reach', 'make': TWO, 'reach': reach, 'member': member, 'value': value} for reach in REACHES for member, value in SETS]
    made += [{'group': 'reach', 'make': ONE + BOLD, 'reach': reach, 'member': 'Font.Italic', 'value': 'True'} for reach in REACHES[:4]]
    made += [{'group': 'reach', 'make': ONE + BOLD, 'reach': reach, 'member': 'Font.Bold', 'value': 'False'} for reach in REACHES[:4]]
    made += [{'group': 'size', 'make': ONE, 'reach': REACHES[0], 'member': 'Font.Size', 'value': value} for value in ('14', '"x"', 'Null', '0')]
    made += [{'group': 'format', 'make': ONE + (f'    ws.Range("D1:D5").FormatConditions(1).NumberFormat = {before}\n' if before else ''),
              'reach': REACHES[0], 'member': 'NumberFormat', 'value': value} for before, value in FORMATS]
    # Second look.
    made += [{'group': 'overlap', 'make': TWO, 'reach': reach, 'member': 'Font.Bold', 'value': 'True'}
             for reach in ('ws.Range("D2:D6").FormatConditions(1)', 'ws.Range("D3:D9").FormatConditions(1)',
                           'ws.Range("C1:D5").FormatConditions(1)', 'ws.Range("D1:D4").FormatConditions(1)')]
    made += [{'group': 'overlap', 'make': TWO + '    ws.Rows(1).Insert\n', 'reach': 'ws.Range("D1:D5").FormatConditions(1)',
              'member': member, 'value': 'True', 'base': 'D2:D6'} for member in ('Font.Bold', 'Font.Italic')]
    made += [{'group': 'origin', 'make': '', 'book': 'file', 'reach': reach, 'member': 'NumberFormat', 'value': '"nonsense["', 'base': 'D10:D12'}
             for reach in (OWN_FILE, 'ws.Cells.FormatConditions(1)', 'ws.Range("A1:H30").FormatConditions(1)')]
    made += [{'group': 'origin', 'make': '', 'book': 'file', 'reach': OWN_FILE, 'member': 'NumberFormat', 'value': '"0.0"', 'base': 'D10:D12'}]
    made += [{'group': 'origin', 'make': ONE + '    ws.Range("D1:D5").FormatConditions(1).NumberFormat = "0.00"\n', 'reach': reach,
              'member': 'NumberFormat', 'value': '"nonsense["'}
             for reach in ('ws.Cells.FormatConditions(1)', 'ws.Range("A1:E9").FormatConditions(1)', 'ws.Range("D2").FormatConditions(1)')]
    return made


def source(case: dict[str, str], typed: bool) -> str:
    kind = 'FormatCondition' if typed else 'Object'
    member = case['member']
    base = case.get('base', 'D1:D5')
    book = ('Set wb = Workbooks.Open(path & "\\l.xlsx"): Set ws = wb.Worksheets("Data")' if case.get('book') == 'file'
            else 'Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)')
    return ('Public Function One(ByVal path As String) As String\n'
            f'    Dim wb As Object, ws As Object, o As {kind}, s As String, e As Long, v As Variant, w As Variant\n'
            f'    Application.DisplayAlerts = False\n    {book}\n'
            '    On Error Resume Next\n' + case['make'] + f'    Set o = {case["reach"]}\n    Err.Clear\n'
            f'    o.{member} = {case["value"]}\n    e = Err.Number: Err.Clear\n'
            f'    v = o.{member}\n    w = ws.Range("{base}").FormatConditions(1).{member}\n'
            '    s = e & " | " & TypeName(v) & ":" & v & " | " & TypeName(w) & ":" & w & " | " & ws.Cells.FormatConditions.Count & " E" & Err.Number\n'
            '    On Error GoTo 0\n    wb.Close False\n    One = s\nEnd Function\n')


def model(case: dict[str, str], typed: bool, folder: str = '') -> str:
    from pyopenvba.apps.excel import ExcelApplication
    from pyopenvba.exceptions import VBAError

    app = ExcelApplication(sandbox=False)
    app.add_workbook()
    app.add_module(source(case, typed), name='Fonts')
    try:
        return str(app.run('Fonts.One', folder))
    except VBAError as failed:
        return f'model {type(failed).__name__}: {str(failed)[:120]}'


def _lock() -> Path:
    common = subprocess.run(['git', 'rev-parse', '--git-common-dir'], capture_output=True, text=True, check=True, cwd=ROOT).stdout.strip()
    return (ROOT / common).resolve() / 'excel-coordination' / 'live.lock'


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig

    sys.stdout.reconfigure(encoding='utf-8')  # type: ignore[union-attr]
    made = cases()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        rounds.books(Path(held))
        lock = _lock()
        owner = rounds.take_lock(lock, 'Verification 02')
        try:
            with ExcelSession(HarnessConfig(lock_wait_s=1200.0, exclusive=False)) as excel:
                excel.new_document()
                for case in made:
                    for tag, typed in (('late', False), ('typed', True)):
                        done = excel.run_vba(source(case, typed), 'One', (held,), timeout=60.0, line_numbers=False)
                        case[tag] = str(done.value) if done.ok else f'no answer: {str(done)[:300]}'
        finally:
            rounds.release_lock(lock, owner)
        for case in made:
            for tag, typed in (('late', False), ('typed', True)):
                case[f'model_{tag}'] = model(case, typed, held)
    OUT.write_text(json.dumps({'cases': made}, indent=1, ensure_ascii=False) + '\n', encoding='utf-8')
    for case in made:
        same = 'same' if case['late'] == case['model_late'] else 'DIFF'
        print(same, case['reach'][3:], case['member'], '=', case['value'], '| Excel', case['late'],
              '' if case['typed'] == case['late'] else f'| typed {case["typed"]}', '| model', case['model_late'][:90])


if __name__ == '__main__':
    main()
