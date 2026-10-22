"""How bond and date functions take dates at the start of the 1900 calendar.

Serial 60 is the 1900 calendar's false 29 February. Excel's bond
functions refuse a settlement anywhere near it, its coupon dates there
stop at serial 0, and YEARFRAC on basis 1 averages the real years'
lengths. Each formula is written to a fresh sheet's cell and read back
as its value or error.

    python scripts/measure_early_1900_dates.py

writes tests/fixtures/formula/early_1900_dates.json, which
tests/test_formula_early_1900_dates.py replays.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/formula/early_1900_dates.json'
FORMULAS = [
    '=ODDLPRICE(60,181,1,0.05,0.06,100,2,1)', '=ODDLPRICE(61,181,1,0.05,0.06,100,2,1)',
    '=ODDLPRICE(59,181,1,0.05,0.06,100,2,1)', '=ODDLPRICE(30,181,1,0.05,0.06,100,2,1)',
    '=ODDLPRICE(70,181,60,0.05,0.06,100,2,1)', '=ODDLPRICE(30,60,1,0.05,0.06,100,2,1)',
    '=ODDLYIELD(60,181,1,0.05,98,100,2,1)', '=ODDFPRICE(60,1000,1,181,0.05,0.06,100,2,1)',
    '=ODDFPRICE(61,1000,1,181,0.05,0.06,100,2,1)', '=ODDFPRICE(59,1000,1,181,0.05,0.06,100,2,1)',
    '=ODDFPRICE(30,1000,1,181,0.05,0.06,100,2,1)',
    '=PRICE(60,1000,0.05,0.06,100,2,1)', '=PRICE(61,1000,0.05,0.06,100,2,1)', '=PRICE(59,1000,0.05,0.06,100,2,1)',
    '=YIELD(60,1000,0.05,98,100,2,1)', '=COUPDAYS(60,1000,2,1)', '=COUPPCD(60,1000,2,1)', '=COUPNCD(60,1000,2,1)',
    '=ACCRINT(1,181,60,0.05,100,2,1)', '=YEARFRAC(60,1000,1)', '=DISC(60,1000,98,100,1)', '=EDATE(60,1)',
    '=EOMONTH(60,0)', '=WORKDAY(60,1)', '=NETWORKDAYS(1,60)', '=WEEKNUM(60)', '=DATEDIF(1,60,"d")',
]


def macro() -> str:
    lines = ['Function Probe() As String', 'Dim ws As Worksheet, s As String, v As Variant', 'Set ws = Worksheets.Add',
             'On Error Resume Next']
    for index, formula in enumerate(FORMULAS):
        quoted = formula.replace('"', '""')
        lines += ['Err.Clear', f'ws.Cells({index + 1}, 1).Formula = "{quoted}"', f'v = ws.Cells({index + 1}, 1).Value2',
                  'If Err.Number <> 0 Then s = s & "|e" & Err.Number Else s = s & "|" & CStr(v)']
    lines += ['Application.DisplayAlerts = False', 'ws.Delete', 'Application.DisplayAlerts = True',
              'Probe = Mid(s, 2)', 'End Function']
    return '\n'.join(lines) + '\n'


if __name__ == '__main__':
    record({'early_1900': macro()}, OUT)
