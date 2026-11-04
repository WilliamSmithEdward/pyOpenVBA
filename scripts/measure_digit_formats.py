"""What a digit does in a number format. Rival rules for 7 under "1.5" showing as 17.5: a format with no placeholder keeps its digits as
literal text and the number goes in front of the first point, or a digit is a placeholder of its own. Splits by values with a fraction, a
negative, zero, large numbers and text, and by formats with digits beside real placeholders, in other sections, after an exponent, and
with no point at all. Each format is tried through Range.Text and through VBA's Format.

    python scripts/measure_digit_formats.py

writes tests/fixtures/digit_formats.json, which tests/test_excel_digit_formats.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/digit_formats.json'
VALUES = ['7', '7.6', '-7', '0', '1234.5', '0.25', '"abc"', 'True']
FORMATS = ['1.5', '1.0', '2', '12', '3.0.0', '.5', '5.', '1.', '1.5.', '1..5', '0.5', '#.5', '1.#', '1.0#', '1#.5', '1,5', '1,000', '1,000.5', '1.5%', '50%', '1.5E+0', '1E+0', '1E+1', '1/2', '1 1/2', '1.5 x', 'x1.5', 'x.y', '"a"1.5', '1.5"a"',
           '1.5;2.5', '1.5;-2.5', '1.5;2.5;3.5', '0;1.5', '1.5;0', '1.5;;', '1.5;2.5;3.5;4.5', '[Red]1.5', '$1.5', '-1.5', '(1.5)', '1:5', '1-5', '1 5', '1.5 2', '9.9', '0', '00.0', '1_x.5', '1*x.5', '1\\.5', '1?5', '?.5', '1.?']

SOURCE = '''Public Function Tried(ByVal which As String, ByVal value As Variant) As String
    Dim c As Object, s As String
    Set c = ActiveSheet.Range("A1")
    On Error Resume Next
    c.NumberFormat = "General"
    c.ClearContents
    c.Value = value
    Err.Clear
    c.NumberFormat = which
    If Err.Number <> 0 Then
        s = "set E" & Err.Number
    Else
        s = c.Text & "|" & c.NumberFormat
    End If
    Err.Clear
    Dim f As String
    f = Format(value, which)
    If Err.Number <> 0 Then s = s & "|E" & Err.Number Else s = s & "|" & f
    Tried = s
End Function

Public Function RunAll(ByVal formats As Variant, ByVal values As Variant) As String
    Dim one As Variant, v As Variant, s As String
    For Each one In formats
        For Each v In values
            s = s & "@@" & Tried(CStr(one), v)
        Next
    Next
    RunAll = Mid$(s, 3)
End Function
'''


def _driver() -> str:
    formats = ', '.join('"' + one.replace('"', '""') + '"' for one in FORMATS)
    values = ', '.join(VALUES)
    return SOURCE + f'\nPublic Function Go() As String\n    Go = RunAll(Array({formats}), Array({values}))\nEnd Function\n'


def main() -> None:
    source = _driver()
    with ExcelSession(HarnessConfig(exclusive=False)) as excel:
        excel.new_document()
        result = excel.run_vba(source, 'Go', (), timeout=170.0)
    assert result.ok, result
    answers = str(result.value).split('@@')
    assert len(answers) == len(FORMATS) * len(VALUES)
    OUT.write_text(json.dumps({'measured_at': datetime.now(timezone.utc).isoformat(), 'formats': FORMATS, 'values': VALUES, 'source': source, 'answers': answers}, indent=1) + '\n', encoding='utf-8')
    for at, one in enumerate(FORMATS):
        print(repr(one), answers[at * len(VALUES):(at + 1) * len(VALUES)], flush=True)


if __name__ == '__main__':
    main()
