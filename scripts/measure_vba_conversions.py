"""Measure the conversion functions over a matrix of inputs, and how a Double or
a Single becomes a Decimal.

Each conversion probe passes 32 inputs (Empty, Null, Booleans, numeric and
non-numeric text, dates, numbers at rounding edges, a Decimal, a Currency and
a Single) through one of CBool, CByte, CInt, CLng, CSng, CDbl, CCur, CDec,
CDate, CStr and CVar, and records TypeName and value, or e and the error
number. Each probe is a module whose ``Function Probe() As String`` the
fixture records (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_conversions.json'

INPUTS = ['Empty', 'Null', 'True', 'False', '""', '" "', '"0"', '"-1"', '"1.5"', '"2.5"', '"  7  "', '"1e3"', '"$5"',
          '"5%"', '"(5)"', '"1,234.5"', '"abc"', '"True"', '"false"', '"&H1F"', '"12/31/2024"', '"13:45"',
          '#1/2/2024 6:00:00 AM#', '0.5', '1.5', '-0.5', '32767.5', '255.5', '-1', 'CDec("1.23456789012345678901234")',
          'CCur(1.23455)', 'CSng(0.1)']
FUNCTIONS = ["CBool", "CByte", "CInt", "CLng", "CSng", "CDbl", "CCur", "CDec", "CDate", "CStr", "CVar"]
DECIMAL_INPUTS = ["1 / 3", "CSng(1 / 3)", "2 ^ 0.5", "1E-20", "123456789.123456789", "CSng(123456.7)", "CSng(1E-10)",
                  "0.1 + 0.2", "1E+20", "CSng(3.4E+38)", "-CSng(2.5)", "CDbl(CSng(0.1))", "1.0000000000000002",
                  "CSng(16777217)", "CCur(1.5)", "#1/1/2024 12:00:00 PM#", "CSng(0.000001234567)", "9.99999999999999E+27",
                  "CDec(0) + (0.1 + 0.2)", "CDec(1) * (1 / 3)", "CDec(2) - CSng(0.1)"]


def converted(function: str) -> str:
    lines = ['Dim s As String, r As Variant', 'On Error Resume Next']
    for value in INPUTS:
        lines += ['Err.Clear', f'r = {function}({value})',
                  'If Err.Number Then s = s & "e" & Err.Number & ";" Else s = s & TypeName(r) & ":" & r & ";"']
    return "\n".join([*lines, 'Probe = s'])


def decimals() -> str:
    lines = ['Dim s As String, r As Variant', 'On Error Resume Next']
    for value in DECIMAL_INPUTS:
        lines += ['Err.Clear', f'r = CDec({value})', 'If Err.Number Then s = s & "e" & Err.Number & ";" Else s = s & CStr(r) & ";"']
    return "\n".join([*lines, 'Probe = s'])


CASES = {f"convert_{function}": converted(function) for function in FUNCTIONS}
CASES.update({
    "decimal_from_numbers": decimals(),
    "byte_from_boolean": 'Dim b As Byte\nb = True\nProbe = b & "|" & CByte(False) & "|" & CByte(True) + 0\nDim v As Variant\nv = CByte(True)\nProbe = Probe & "|" & TypeName(v)',
    "integer_from_boolean": 'Dim i As Integer, l As Long, c As Currency, d As Double, s As Single\ni = True: l = True: c = True: d = True: s = True\nProbe = i & l & c & d & s',
    "byte_overflow_negative": 'On Error Resume Next\nDim b As Byte\nb = -1\nProbe = Err.Number & "|" & b',
    "variant_decimal_sum": 'Dim v As Variant\nv = CDec(CSng(0.1)) + 0\nProbe = v & "|" & TypeName(v)',
})


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
