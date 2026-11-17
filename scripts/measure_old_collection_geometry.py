"""Left, Top, Width and Height read and set on a whole old collection, Worksheet.Buttons.Left = 45 and the like, typed and late, on the
workbook Excel saved (tests/fixtures/old_collections.xlsx) and on one the macro makes: what the collection reads before, what the set
answers, what the collection and each item read after, and what each item reads once the workbook is saved and opened again.

Rival rules, written before measuring:

- A read over items that agree answers their value as a Double; over items that differ, 1004 as most reads do (rival: the first
  item's value; rival: Null).
- A set reaches every item, each taking the value as a ShapeRange's set does (rival: only the first; rival: the items move or grow
  together, keeping their offsets). DrawingObjects takes it on every kind.
- A value off the display's pixels (31.25) reads back on them, as on one item (31 or 31.5 at 144 DPI) (rival: as set).
- A picture keeps its aspect, so setting its Width moves its Height too (rival: only the Width changes).
- The save keeps what the session reads.

Sizes are in the display's pixels. The values set are on both 96- and 144-DPI pixels (multiples of 1.5pt) but for the one off both.
Every row records StandardHeight before and after, and a run where it changes is refused; the fixture records the scale it ran at
(William's rule of 2026-11-17: DPI-dependent work runs at the display there is). Each row runs in a module of its own.

    python scripts/measure_old_collection_geometry.py [seconds]

writes tests/fixtures/old_collection_geometry.json, which tests/test_excel_old_collection_geometry.py replays. It keeps its place in a
file under TEMP, so a run that stops can go on.
"""
from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'tests/fixtures/old_collection_geometry.json'
BOOK = ROOT / 'tests/fixtures/old_collections.xlsx'
PARTIAL = Path(tempfile.gettempdir()) / 'old_collection_geometry_partial.json'
BUDGET = 520.0
COLLECTIONS = ('Arcs', 'Buttons', 'CheckBoxes', 'DropDowns', 'Lines', 'ListBoxes', 'OptionButtons', 'Ovals', 'Pictures', 'Rectangles', 'TextBoxes', 'DrawingObjects')
#: A value on both the 96- and the 144-DPI pixels (0.75pt and 0.5pt), and one off both.
ON_GRID = {'Left': '45', 'Top': '37.5', 'Width': '75', 'Height': '37.5'}
OFF_GRID = '31.25'
_reads = importlib.util.spec_from_file_location('measure_old_collection_reads', ROOT / 'scripts/measure_old_collection_reads.py')
assert _reads is not None and _reads.loader is not None
_reads_module = importlib.util.module_from_spec(_reads)
_reads.loader.exec_module(_reads_module)
MAKE: str = getattr(_reads_module, 'MAKE')
#: Follows MAKE, whose Show it uses.
DRIVER = r'''Private Function Items(ws As Object) As String
    Dim x As Object, s As String
    On Error Resume Next
    For Each x In ws.{c}
        Err.Clear
        s = s & ";" & x.Left & "," & x.Top & "," & x.Width & "," & x.Height
        If Err.Number Then s = s & "E" & Err.Number
    Next
    Items = Mid(s, 2)
End Function

Private Function Whole(o As {declared}) As String
    Dim v As Variant
    On Error Resume Next
    v = o.{m}
    If Err.Number Then Whole = "E" & Err.Number Else Whole = Show(v)
End Function

Public Function Row(ByVal path As String, ByVal saved As String, ByVal picture As String) As String
    Dim wb As Workbook, ws As Object, o As {declared}, e As Long, s As String, before As String
    Application.DisplayAlerts = False
    {open}
    Set ws = wb.Worksheets(1)
    before = ws.StandardHeight
    Set o = ws.{c}
    s = Whole(o)
    On Error Resume Next
    o.{m} = {v}
    e = Err.Number
    Err.Clear
    On Error GoTo 0
    s = s & "|E" & e & "|" & Whole(o) & "|" & Items(ws)
    wb.SaveAs saved, xlOpenXMLWorkbook
    wb.Close False
    Set wb = Workbooks.Open(saved, 0)
    s = s & "|reopened " & Items(wb.Worksheets(1)) & "|" & before & "," & wb.Worksheets(1).StandardHeight
    wb.Close False
    Row = s
End Function
'''


def cases() -> list[dict[str, str]]:
    found: list[dict[str, str]] = []
    for collection in COLLECTIONS:
        for member, value in ON_GRID.items():
            for book in ('saved', 'made'):
                for binding in ('Typed', 'Late'):
                    found.append({'collection': collection, 'member': member, 'value': value, 'book': book, 'binding': binding})
            found.append({'collection': collection, 'member': member, 'value': OFF_GRID, 'book': 'made', 'binding': 'Late'})
    return found


def key_of(case: dict[str, str]) -> str:
    return ' '.join(case[part] for part in ('collection', 'member', 'value', 'book', 'binding'))


def row_source(case: dict[str, str]) -> str:
    declared = f'Excel.{case["collection"]}' if case['binding'] == 'Typed' else 'Object'
    opened = 'Set wb = Workbooks.Open(path, 0)' if case['book'] == 'saved' else 'Set wb = Made(picture)'
    driver = (DRIVER.replace('{open}', opened).replace('{declared}', declared).replace('{c}', case['collection']).replace('{m}', case['member'])
              .replace('{v}', case['value']))
    return MAKE + driver


def main() -> None:
    from pyvbaharness import ExcelSession, HarnessConfig, HarnessError

    budget = float(sys.argv[1]) if len(sys.argv) > 1 else BUDGET
    all_cases = cases()
    done: dict[str, str] = json.loads(PARTIAL.read_text(encoding='utf-8')) if PARTIAL.exists() else {}
    started = time.monotonic()
    session = ExcelSession(HarnessConfig(exclusive=False))
    session.new_document()
    tiny_png = getattr(_reads_module, 'tiny_png')
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as held:
        picture = Path(held) / 'dot.png'
        picture.write_bytes(tiny_png())
        for number, case in enumerate(all_cases):
            key = key_of(case)
            if key in done:
                continue
            if time.monotonic() - started > budget:
                print('stopping for now', len(done), 'of', len(all_cases), flush=True)
                break
            book = Path(held) / f'book{number}.xlsx'
            book.write_bytes(BOOK.read_bytes())
            try:
                ran = session.run_vba(row_source(case), 'Row', (str(book), str(Path(held) / f'saved{number}.xlsx'), str(picture)), timeout=60.0, line_numbers=False)
            except HarnessError as failed:
                print('restarting', str(failed)[:80], flush=True)
                session.close()
                time.sleep(5)
                session = ExcelSession(HarnessConfig(exclusive=False))
                session.new_document()
                done[key] = f'failed harness {str(failed)[:60]}'
                continue
            answer = str(ran.value) if ran.ok and not ran.dialogs else f'failed {ran.outcome} {[one.message[:90] for one in ran.dialogs]}'
            if ran.ok:
                scale = answer.rsplit('|', 1)[-1]
                first = next((one.rsplit('|', 1)[-1] for one in done.values() if not one.startswith('failed')), scale)
                if scale != f'{scale.split(",")[0]},{scale.split(",")[0]}' or scale != first:
                    session.close()
                    raise SystemExit(f'The display changed during the run at {key}: {scale}, the run began at {first}')
            done[key] = answer
            print(key, answer[:160], flush=True)
            PARTIAL.write_text(json.dumps(done), encoding='utf-8')
    session.close()
    if len(done) == len(all_cases):
        height = next(float(one.rsplit('|', 1)[-1].split(',')[0]) for one in done.values() if not one.startswith('failed'))
        OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'standard_height': height, 'driver': DRIVER,
                                   'cases': [{**case, 'excel': done[key_of(case)]} for case in all_cases]}, indent=0) + '\n', encoding='utf-8')
        PARTIAL.unlink()
        print('wrote', len(all_cases), 'cases', flush=True)


if __name__ == '__main__':
    main()
