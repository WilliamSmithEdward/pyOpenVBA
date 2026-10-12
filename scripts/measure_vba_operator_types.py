"""Measure the type each arithmetic operator gives for every pair of operand
types, and the language details the sweep found next to it: hex literal
widths, the Mid$ statement, As New arrays and Like's bracket classes.

The matrix probes hold twelve Variants, Byte to Empty, and record the first
three letters of TypeName for each pair, or e and the error number. Each
probe is a module whose ``Function Probe() As String`` the fixture records
(see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_operator_types.json'

MATRIX_HEAD = '''Dim v(11) As Variant, i As Long, j As Long, s As String, r As Variant
v(0) = CByte(2): v(1) = CInt(2): v(2) = CLng(2): v(3) = CLngLng(2): v(4) = CSng(2): v(5) = CDbl(2)
v(6) = CCur(2): v(7) = CDec(2): v(8) = CDate(2): v(9) = "2": v(10) = True: v(11) = Empty
On Error Resume Next
For i = 0 To 11
For j = 0 To 11
Err.Clear
r = Empty
'''
MATRIX_TAIL = '''
If Err.Number Then s = s & "e" & Err.Number & "," Else s = s & Left(TypeName(r), 3) & ","
Next
s = s & ";"
Next
Probe = s'''
OPERATIONS = {"add": "r = v(i) + v(j)", "sub": "r = v(i) - v(j)", "mul": "r = v(i) * v(j)", "div": "r = v(i) / v(j)",
              "idiv": "r = v(i) \\ v(j)", "mod": "r = v(i) Mod v(j)", "pow": "r = v(i) ^ v(j)", "neg": "r = -v(i)"}

CASES = {f"types_{name}": MATRIX_HEAD + body + MATRIX_TAIL for name, body in OPERATIONS.items()}
CASES.update({
    "byte_bitwise": 'Dim a As Byte, b As Byte\na = 12: b = 10\nProbe = TypeName(a And b) & TypeName(a Or b) & TypeName(a Xor b) & TypeName(Not a) & "|" & (Not a) & "|" & TypeName(a And 1)',
    "hex_literals": 'Probe = &H8000 & "|" & &H8000& & "|" & &HFFFF & "|" & &H10000 & "|" & &O777 & "|" & TypeName(&H8000) & TypeName(&H10000) & TypeName(&H8000&) & "|" & &HFFFFFFFF & "|" & &HFFFF&',
    "zero_to_negative_power": 'On Error Resume Next\nDim r As Double\nr = 0 ^ -1\nProbe = Err.Number',
    "mid_statement": 'Dim s As String\ns = "abcdef"\nMid(s, 2, 3) = "XYZW"\nProbe = s\nMid(s, 5) = "12345"\nProbe = Probe & "|" & s\nMid$(s, 1, 1) = "q"\nProbe = Probe & "|" & s',
    "as_new_array": 'Dim c(1) As New Collection\nWith c(1)\n.Add 1\n.Add 2\nEnd With\nProbe = c(1).Count & "|" & c(0).Count & "|" & TypeName(c) & "|" & (c(0) Is Nothing)',
    "like_bracket_edges": 'On Error Resume Next\nDim r As Boolean\nr = ("a" Like "[")\nProbe = Err.Number & "|" & r\nErr.Clear\nr = ("-" Like "[-a]")\nProbe = Probe & "|" & r & ("b" Like "[a-]") & ("[" Like "[[]") & ("^" Like "[^]") & ("&" Like "[&]")',
    "date_arithmetic": 'Dim d As Date\nd = #1/1/2024#\nProbe = TypeName(d + 1) & TypeName(d - 1) & TypeName(d - d) & TypeName(d * 1) & TypeName(1 - d) & TypeName(-d) & "|" & (d + 1) & "|" & (d - #12/31/2023#) & "|" & (5 - d)',
    "currency_arithmetic": 'Dim c As Currency\nc = 1.5\nProbe = TypeName(c + 1.5) & TypeName(c * 1.5) & TypeName(c / 2) & TypeName(c * 2) & "|" & (c + 0.123456) & "|" & (c * 1.5)',
})


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
