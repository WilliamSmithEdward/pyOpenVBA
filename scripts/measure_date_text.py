"""Measure how native VBA reads text as a date: IsDate, and CDate's result.

Each probe records "T" or "F" for IsDate, then CDate's result as yyyy-mm-dd
hh:nn:ss, or the error it raises. Text with no year takes the year it is read
in; the fixture records that year, and no text spells it out.
"""
from __future__ import annotations
import json
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/date_text.json'
TEXTS = [
    "3/4/2024", "3/4/24", "3/4", "13/4/2024", "3/13/2024", "31/12/2024", "2024/3/4", "3-4-2024", "3.4.2024",
    "03/04/2024", "3/4/1", "3/4/29", "3/4/30", "3/4/99", "3/4/100", "3/4/9999", "3/4/10000", "0/4/2024",
    "2/30/2024", "2/29/2023", "2/29/2024", "13/13/2024", "4/31/2024",
    "2024-03-04", "2024-3-4", "2024-03-04 13:00", "2024-03-04 13:00:05", "2024-03-04T13:00", "2024-03-04 1:05 PM",
    "2024-13-01", "24-03-04", "2024-04-31",
    "Mar 4 2024", "March 4, 2024", "4 Mar 2024", "4-Mar-2024", "4-Mar-24", "Mar-2024", "March 2024", "Mar 4",
    "4 March", "Mar. 4, 2024", "Sept 4 2024", "Sep 4 2024", "Mon, Mar 4 2024", "Monday, March 4, 2024",
    "mar 4 2024", "MARCH 4 2024", "Marc 4 2024", "2024 Mar 4", "4 2024 Mar", "Mar 2024 4", "Ma 4 2024",
    "Febr 4 2024", "June 31 2024", "Mar 4 24", "Mar 32 2024",
    "13:05", "1:05 PM", "1:05PM", "1 PM", "1PM", "13:05:07", "13:05:07.5", "25:00", "12:60", "0:00",
    "12:00 AM", "12:00 PM", "13:00 PM", "1:05 a", "1:05 p", "1:05 am", "1:5", "1:05:60", "24:00", "23:59:59",
    "3/4/2024 13:05", "3/4/2024 1:05 PM", "1:05 PM 3/4/2024", "3/4/2024 1 PM", "March 4, 2024 1:05:07 PM",
    "3/4/2024 25:00", "13:05 Mar 4 2024",
    "2024", "3", "3/4/2024 x", "20240304", "2024.03.04", "4.3.2024", "1/1/100", "12/31/9999", "1/1/99",
    "1/1/1899", "today", "  3/4/2024  ", "3 / 4 / 2024", "3//4/2024", "3/4/2024/5", "-3/4/2024", "3,4,2024",
    "3 4 2024", "3/4/2024 PM", "PM 1:05", "1:05 PM PM", "12/30/1899", "1/0/2024",
    "1/1/49", "1/1/50", "1/1/48", "1/1/51", "3/2024", "13/4", "4/13", "Mar", "March 4 2024 5", "Mar 24 4",
    "Mar 4 1", "Mar 2 4", "4 Mar", "2024 4 Mar", "4 24 Mar", "Mar 32", "Mar 99", "Mar 100", "2024-03",
    "03-2024", "1 1 1", "13 13 13", "Mar 4, 2024, 1:05", "3/4/2024, 13:05", "3/4/2024 -", "Mar 4 2024 Apr",
    "1:05:07 PM 3/4/2024", "3/4/2024 1:05:07 AM", "1:05 3/4/2024 PM", "3/4/2024 13", "3/4/2024 1",
    "Sat 3/4/2024", "3/4/2024 Mon", "Feb 29 2023", "Feb 29 2024", "may 4 2024", "Ju 4 2024", "Jul 4 2024",
    "Augu 4 2024", "Decem 4 2024", "Dece 4 2024", "Decembe 4 2024", "Decemberx 4 2024", "Marc. 4 2024",
    "Mar 4. 2024", "4.Mar.2024", "4/Mar/2024", "Mar/4/2024", "4-mar-2024", "2024-Mar-04",
    "1/1/0", "1/1/00", "1/1/000", "1/1/0100", "1/1/99999", "100/1/1", "99/1/1", "32/1/1", "31/1/1", "12/31/1",
    "0:5", "00:00:00", "1:2:3", "1:02:03:04", "12 AM", "0 AM", "0 PM", "12 PM", "11 PM", "a 1:05",
    "1:05 A.M.", "1:05 AM.", "13 AM", "1 2", "1/2 PM",
]
HEAD = r'''
Private Function Try(ByVal s As String) As String
Dim r As String
r = IIf(IsDate(s), "T", "F") & "|"
On Error Resume Next
Dim d As Date
d = CDate(s)
If Err.Number Then r = r & "e" & Err.Number Else r = r & Format(d, "yyyy-mm-dd hh:nn:ss")
Try = r
End Function
'''


def macro(text: str) -> str:
    return HEAD + '\nFunction Probe() As String\nProbe = Try("' + text.replace('"', '""') + '")\nEnd Function\n'


def main() -> None:
    probes = []
    with ExcelSession(HarnessConfig(lock_wait_s=30.0)) as excel:
        excel.new_document()
        version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build & "|" & Year(Date)\nEnd Function', 'VersionBuild')
        assert version.ok and not version.dialogs, version
        build, measured_year = str(version.value).rsplit('|', 1)
        for text in TEXTS:
            result = excel.run_vba(macro(text), 'Probe', timeout=30.0)
            assert result.ok and not result.dialogs, (text, result)
            probes.append({'text': text, 'macro': macro(text), 'result': result.value})
            print(repr(text), result.value, flush=True)
    OUT.write_text(json.dumps({'excel_version_build': build, 'measured_year': int(measured_year),
                               'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n',
                   encoding='utf-8')


if __name__ == '__main__':
    main()
