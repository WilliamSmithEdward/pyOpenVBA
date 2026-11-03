"""Range.FlashFill: a column filled from the examples above it, read from the columns beside it. Each case gives inputs in column A (and
B where it says), examples at the top of the output column, and reads what Excel filled, or the error.

    python scripts/measure_flash_fill.py

writes tests/fixtures/flash_fill.json, which tests/test_excel_flash_fill.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/flash_fill.json'
NAMES = ['Ann Lee', 'Bob Ray', 'Cy Poe', 'Dee Fox']
THREE = ['Ann May Lee', 'Bob Jo Ray', 'Cy Di Poe', 'Dee Al Fox']
MAIL = ['ann.lee@x.com', 'bob.ray@y.org', 'cy.poe@z.net', 'dee.fox@w.io']
CODES = ['AB-123-X', 'CD-456-Y', 'EF-789-Z', 'GH-012-W']
#: name: inputs, second column or None, examples, the range FlashFill runs on in the output column as (first, last) rows, or None for all.
CASES = {
    'first_word': (NAMES, None, ['Ann'], None),
    'last_word': (NAMES, None, ['Lee'], None),
    'first_word_two_examples': (NAMES, None, ['Ann', 'Bob'], None),
    'middle_word': (THREE, None, ['May'], None),
    'last_of_three': (THREE, None, ['Lee'], None),
    'first_of_three': (THREE, None, ['Ann'], None),
    'first_letter': (NAMES, None, ['A'], None),
    'initials': (NAMES, None, ['AL'], None),
    'initials_dotted': (NAMES, None, ['A.L.'], None),
    'initials_two_examples': (NAMES, None, ['AL', 'BR'], None),
    'swap': (NAMES, None, ['Lee, Ann'], None),
    'swap_space': (NAMES, None, ['Lee Ann'], None),
    'upper_first': (NAMES, None, ['ANN'], None),
    'lower_whole': (NAMES, None, ['ann lee'], None),
    'upper_whole': (NAMES, None, ['ANN LEE'], None),
    'prefix_literal': (NAMES, None, ['Mr Ann'], None),
    'suffix_literal': (NAMES, None, ['Ann!'], None),
    'whole_copy': (NAMES, None, ['Ann Lee'], None),
    'mail_user': (MAIL, None, ['ann.lee'], None),
    'mail_domain': (MAIL, None, ['x.com'], None),
    'mail_first': (MAIL, None, ['ann'], None),
    'mail_tld': (MAIL, None, ['com'], None),
    'code_middle': (CODES, None, ['123'], None),
    'code_first': (CODES, None, ['AB'], None),
    'code_last': (CODES, None, ['X'], None),
    'code_no_dashes': (CODES, None, ['AB123X'], None),
    'first_two_chars': (CODES, None, ['AB-'], None),
    'two_columns_join': (['Ann', 'Bob', 'Cy', 'Dee'], ['Lee', 'Ray', 'Poe', 'Fox'], ['Ann Lee'], None),
    'two_columns_second': (['Ann', 'Bob', 'Cy', 'Dee'], ['Lee', 'Ray', 'Poe', 'Fox'], ['Lee'], None),
    'two_columns_mail': (['Ann', 'Bob', 'Cy', 'Dee'], ['Lee', 'Ray', 'Poe', 'Fox'], ['ann.lee@x.com'], None),
    'numbers_as_text': (['12-5', '34-6', '56-7', '78-8'], None, ['12'], None),
    'numbers_real': ([1234, 5678, 9012, 3456], None, ['12'], None),
    'no_relation': (NAMES, None, ['zzz'], None),
    'contradiction': (NAMES, None, ['Ann', 'Ray'], None),
    'uneven_words': (['Ann Lee', 'Bob', 'Cy Di Poe', 'Dee Fox'], None, ['Ann'], None),
    'uneven_last': (['Ann Lee', 'Bob', 'Cy Di Poe', 'Dee Fox'], None, ['Lee'], None),
    'blank_input': (['Ann Lee', '', 'Cy Poe', 'Dee Fox'], None, ['Ann'], None),
    'only_part_of_column': (NAMES, None, ['Ann'], (2, 3)),
    'single_row_range': (NAMES, None, ['Ann'], (2, 2)),
    'existing_below': (NAMES, None, ['Ann', None, 'keep'], None),
}
PROBE = """Public Function Fill(ByVal inputs As String, ByVal seconds As String, ByVal examples As String, ByVal first As Long, ByVal last As Long) As String
    Dim ws As Object, i As Long, s As String, column As Long, n As Long
    Application.DisplayAlerts = False
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    ws.Cells(1, 1).Value = "In"
    column = 2
    If seconds <> "" Then column = 3
    ws.Cells(1, column).Value = "Out"
    n = 0
    Dim parts As Variant
    parts = Split(inputs, "|")
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then ws.Cells(i + 2, 1).Value = parts(i)
        n = n + 1
    Next
    If seconds <> "" Then
        parts = Split(seconds, "|")
        ws.Cells(1, 2).Value = "In2"
        For i = 0 To UBound(parts)
            ws.Cells(i + 2, 2).Value = parts(i)
        Next
    End If
    parts = Split(examples, "|")
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then ws.Cells(i + 2, column).Value = parts(i)
    Next
    If last = 0 Then last = n + 1
    On Error Resume Next
    ws.Range(ws.Cells(first, column), ws.Cells(last, column)).FlashFill
    s = Err.Number & "|"
    On Error GoTo 0
    For i = 2 To n + 1
        s = s & ws.Cells(i, column).Value & ";"
    Next
    ws.Parent.Close False
    Application.DisplayAlerts = True
    Fill = s
End Function
"""


def joined(values: list[object] | None) -> str:
    return '|'.join('' if value is None else str(value) for value in values or [])


def main() -> None:
    cases: dict[str, dict[str, object]] = {}
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for name, (inputs, seconds, examples, part) in CASES.items():
            first, last = part or (2, 0)
            second = joined(seconds) if seconds else ''
            result = excel.run_vba(PROBE, 'Fill', (joined(inputs), second, joined(examples), first, last), timeout=60.0)
            assert result.ok and not result.dialogs, (name, result.error, result.message)
            cases[name] = {'inputs': inputs, 'seconds': seconds, 'examples': examples, 'part': part, 'answer': result.value}
            print(name.ljust(26), result.value, flush=True)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'source': PROBE, 'cases': cases}, indent=1) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
