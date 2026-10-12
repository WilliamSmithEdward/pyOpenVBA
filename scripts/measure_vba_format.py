"""Measure VBA's Format function: number, date and text formats, sections, named formats.

Each probe calls Format(value, format) in native Excel's VBA and records the
result as TypeName:text, or e<number> for a run-time error.
"""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_format.json'
T = "#3/4/2024 1:05:09 PM#"
AM = "#3/4/2024 9:07:03 AM#"
PAIRS: list[tuple[str, str]] = [
    # Empty, Null and text
    ("Empty", '"0"'), ("Empty", '""'), ("Empty", '"#"'), ("Empty", '"0;-0;Z"'), ("Empty", '"yyyy"'), ("Empty", '"@"'),
    ("Null", '"0"'), ("Null", '"0;-0;Z;N"'), ("Null", '""'), ('""', '"0"'), ('"abc"', '"0"'), ('"12abc"', '"0"'),
    ('"12"', '"0.00"'), ('"-12.5"', '"0.0"'), ('"abc"', '"yyyy"'), ('"3/4/2024"', '"yyyy-mm-dd"'), ("True", '"0"'),
    ("False", '"0.0"'), ("True", '""'), ("True", '"Yes/No"'),
    # no format
    ("1234.5678", '""'), ("1234.5678", "Empty"), ("CSng(1/3)", '""'), ("CCur(1.5)", '""'), ("#3/4/2024#", '""'),
    (T, '""'), ("-0.5", '""'), ("1E+20", '""'), ("12345678901234567", '""'),
    # named number formats
    ("1234.5", '"General Number"'), ("1234.5", '"Currency"'), ("-1234.5", '"Currency"'), ("1234.5", '"Fixed"'),
    ("-1234.567", '"Fixed"'), ("1234.5", '"Standard"'), ("0.1234", '"Percent"'), ("1234.5", '"Scientific"'),
    ("0", '"Yes/No"'), ("5", '"True/False"'), ("1", '"On/Off"'), ("0", '"On/Off"'), ("-3", '"Yes/No"'),
    # digits, rounding, decimal point
    ("-0.4", '"0"'), ("-0.4", '"#"'), ("0", '"#"'), ("0.5", '"#"'), ("1.5", '"0"'), ("2.5", '"0"'), ("-2.5", '"0"'),
    ("0.125", '"0.00"'), ("0.135", '"0.00"'), ("1.005", '"0.00"'), ("1E+20", '"0"'), ("1E-20", '"0.00"'), ("123", '"00000"'),
    ("123.456", '"#.##"'), ("123", '"#.##"'), ("0.5", '"#.##"'), ("0.5", '"0.##"'), ("10", '"0.##"'), ("0", '"0.##"'),
    ("0", '"#.##"'), ("5", '"."'), ("5", '"0."'), ("0.5", '".0"'), ("5", '"#"'), ("123", '"0#0"'), ("12345", '"##"'),
    ("-12345", '"00"'), ("0.0001", '"0.000"'), ("-0.0001", '"0.000"'), ("-0.0001", '"0.000;(0.000)"'),
    # grouping and scaling
    ("1234567", '"#,##0"'), ("1234567", '"#,##0,"'), ("1234567", '"#,##0,,"'), ("1234567", '"0.0,,"'),
    ("1234567", '"0,.0"'), ("1234567.891", '"#,##0.00"'), ("1234", '"#,#"'), ("1234", '"0,0"'), ("999", '"#,##0"'),
    ("-1234567", '"#,##0"'), ("1234567", '"#,"'), ("1234567", '"0,000"'),
    # percent and scientific
    ("5", '"0%"'), ("0.05", '"0.0%"'), ("0.05", '"%0"'), ("0.0567", '"0.00%"'), ("12", '"0.0E+00"'), ("12", '"0.0E-00"'),
    ("0.00012", '"0.0e-00"'), ("12", '"##0.0E+0"'), ("123456", '"0.00E+00"'), ("-0.00012", '"0.00E+00"'), ("0", '"0.0E+00"'),
    ("12", '"0E+0"'), ("1234", '"#E+#"'),
    # sections
    ("-5", '"0;(0)"'), ("0", '"0;(0);\\z\\e\\r\\o"'), ("5", '"0;(0);\\z"'), ("-5", '"0;;"'), ("5", '";;;"'), ("-5", '";;;"'),
    ("0", '"0;;"'), ("-5", '"0;"'), ("-5", '"#;#"'), ("0", '"0;-0"'), ("-0.4", '"0;-0;\\z"'), ("0.4", '"0;-0;\\z"'),
    ("-5", '"0;-0;\\z;\\n"'),
    # literals
    ("5", '"\\$0"'), ("5", '"""USD ""0"'), ("5", '"0 \\x"'), ("5", '"$0"'), ("-5", '"$#,##0.00"'), ("5", '"0.00 ""pts"""'),
    ("1234.5", '"#,##0.00 ;(#,##0.00)"'), ("-5", '"(0)"'), ("5", '"+0"'), ("-5", '"+0"'), ("5", '"0-"'), ("5", '"[Red]0"'),
    ("-5", '"0;[Red]-0"'), ("5", '"x0y"'), ("5", '"ab0"'), ("12", '"0 0"'), ("5", '"0 hrs"'), ("-5", '"""a""0"'),
    ("0", '"#,##0.00;;\\-"'), ("1", '"#,##0.00;;\\-"'),
    # text formats
    ('"abc"', '">"'), ('"ABC"', '"<"'), ('"12345"', '"@@@-@@"'), ('"ab"', '"@@@@"'), ('"ab"', '"!@@@@"'), ('"abcdef"', '"@@@"'),
    ('"ab"', '"&&&&"'), ('"ab"', '"@-@-@"'), ("123", '"@@@@@"'), ('"abcdef"', '"!@@@"'), ('"ab"', '"(@@@)"'),
    ('"ab"', '"@@@@;x"'), ('"aB"', '">@@@"'), ('"ab"', '"\\@@@@"'),
    # dates and times
    (T, '"h hh n nn s ss AM/PM"'), (T, '"h:nn"'), (T, '"hh:nn am/pm"'), (T, '"h:nn A/P"'), (T, '"h:nn a/p"'), (T, '"h:nn AMPM"'),
    (T, '"h AM/PM"'), (T, '"hh AM/PM"'), (T, '"AM/PM h"'), (T, '"h"'), (T, '"hh"'), (T, '"h n"'), (AM, '"h hh AM/PM"'),
    (AM, '"hh:nn:ss AM/PM"'), ("#12:00:00 AM#", '"hh:nn AM/PM"'), ("#12:30:00 PM#", '"h AM/PM"'), ("#12:30:00 AM#", '"h AM/PM"'),
    (T, '"ttttt"'), (T, '"c"'), ("#3/4/2024#", '"c"'), ("#1:05:09 PM#", '"c"'), (T, '"ddddd"'), (T, '"dddddd"'), (T, '"mmmmm"'),
    (T, '"m/d/yy"'), (T, '"mm/dd/yyyy hh:mm:ss"'), (T, '"hh:mm"'), (T, '"mm"'), (T, '"m"'), (T, '"n"'), (T, '"yyy"'), (T, '"y"'),
    (T, '"w ww"'), ("#1/1/2021#", '"ww"'), (T, '"q"'), (T, '"dddd, mmmm d"'), (T, '"\\d\\a\\y d"'), (T, '"d dd ddd dddd"'),
    (T, '"m mm mmm mmmm"'), (T, '"yy yyyy q w ww y"'), (T, '"s:mm"'), (T, '"h:m:s"'), (T, '"hh:mm:ss"'), (T, '"mm:ss"'),
    (T, '"[h]:mm"'), (T, '"yyyy-mm-ddThh:nn:ss"'), (T, '"dd.mm.yyyy"'), (T, '"""Day ""d"'), (T, '"/"'), (T, '":"'),
    (T, '"Short Date"'), (T, '"Long Date"'), (T, '"Medium Date"'), (T, '"Short Time"'), (T, '"Medium Time"'),
    (T, '"Long Time"'), (T, '"General Date"'), ("#3/4/2024#", '"General Date"'), ("#1:05:09 PM#", '"General Date"'),
    ("45355.5", '"dd/mm/yyyy"'), ("1.5", '"hh:nn"'), ("-1", '"yyyy-mm-dd"'), ("-1.5", '"yyyy-mm-dd hh:nn"'),
    ("45355.5", '"0.0"'), (T, '"0.00"'), ("2958466", '"yyyy"'), ("#12/31/9999#", '"dddddd"'),
    # token case, literal-only formats, early years, mixed kinds
    (T, '"YYYY-MM-DD"'), (T, '"DDDD MMMM"'), (T, '"HH:NN:SS"'), (T, '"Hh:Nn"'), (T, '"Q W WW Y"'), (T, '"C"'),
    (T, '"short date"'), ("1234.5", '"currency"'), ("1234.5", '"STANDARD"'), ("5", '"abc"'), ("5", '"\\a\\b"'),
    ("-5", '"abc"'), ("0", '"abc"'), (T, '"abc"'), ("#1/1/100#", '"yyyy"'), ("#1/1/100#", '"yy"'), ("#3/4/999#", '"m/d/yyyy"'),
    (T, '"@"'), ("45355", '"@@@@@@"'), ("5", '"yyyy"'), ('"12"', '"@@@@"'), ("-5", '"@@@"'), ("True", '"@"'),
    ("CCur(-1234.5)", '"Currency"'), ("0", '"Currency"'), ("-0.004", '"Currency"'), ("1234.5", '"General Number;x"'),
    ("1E+16", '"General Number"'), ("0.1", '"Scientific"'), ('CDec("79228162514264337593543950335")', '"0"'),
    ("-5", '"0.0;"'), ("0", '";-0;"'), ("5", '"0;-0;0;0"'), ("0", '"#;-#;"'),
]
HEAD = r'''
Private Function F(v As Variant, fmt As Variant) As String
On Error Resume Next
Dim r As Variant
r = Format(v, fmt)
If Err.Number Then F = "e" & Err.Number Else F = TypeName(r) & ":" & CStr(r)
End Function
'''


def macro(value: str, pattern: str) -> str:
    return HEAD + f'\nFunction Probe() As String\nProbe = F({value}, {pattern})\nEnd Function\n'


def main() -> None:
    probes = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        for value, pattern in PAIRS:
            result = excel.run_vba(macro(value, pattern), 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (value, pattern, result)
            probes.append({'value': value, 'format': pattern, 'macro': macro(value, pattern), 'result': result.value})
            print(value, pattern, '=>', repr(result.value), flush=True)
    OUT.write_text(json.dumps({'excel_version_build': version.value, 'measured_at': datetime.now(timezone.utc).isoformat(),
                               'probes': probes}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
