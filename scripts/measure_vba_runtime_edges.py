"""Measure VBA run-time edges the differential sweep found: typed-array return
types, Collection keys and positions, Round on a Double, negative zero, 0 / 0,
Val of hex text, AscW, IsDate of a number, dotted time text and the optional
arguments of FormatNumber, FormatCurrency and FormatPercent.

Each probe is a module whose ``Function Probe() As String`` the fixture records
as a value or a run-time error number (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_runtime_edges.json'

HELPERS = '''
Private Function MakeArr() As Variant
    MakeArr = Array(1, "two", 3.5)
End Function

Private Function LongArr() As Long()
    Dim a(1 To 3) As Long
    a(1) = 10: a(2) = 20: a(3) = 30
    LongArr = a
End Function

Private Function Words() As String()
    Words = Split("a b c")
End Function
'''

ROUND_VALUES = [("2.345", 2), ("2.675", 2), ("1.005", 2), ("0.125", 2), ("0.375", 2), ("1.115", 2), ("2.5", 0), ("0.285", 2),
                ("1.255", 2), ("8.345", 2), ("4.355", 2), ("0.045", 2), ("1234.5675", 3), ("0.0005", 3), ("-2.345", 2),
                ("-0.125", 2), ("5.015", 2), ("3.14159", 4), ("10.005", 2), ("-0.4", 0), ("-0.5", 0), ("-0.6", 0),
                ("-0.05", 1), ("-0.04", 1), ("-0.005", 2), ("-0#", 0), ("-1E-300", 0), ("-0.25", 1), ("CSng(-0.5)", 0),
                ("CSng(2.345)", 2), ("CCur(2.345)", 2), ('CDec("2.345")', 2)]
# Year-free texts only: "1,5" and the like take the current year.
DATE_TEXTS = ["1.5", "1.5.2024", "3.14", "12.31", "13.1", "1.5.24", "0.5", "1.0", "1. 5", "1.5.", "10.30", "2.30", " 2.5 ",
              "1.5AM", "1.5 PM", "23.59", "24.00", "1.60", "1.5.60", "1.255", "001.05", "+1.5", "12.30.45", "1.5 1/2/2024",
              "1/2/2024 1.5", "1.5:30", "12:30.45", "25.1", "23.59.60", "1.5.2024 10.30", "45000", "1.5%"]


def date_text(text: str) -> str:
    return (f'On Error Resume Next\nDim d As Date\nd = CDate("{text}")\n'
            f'If Err.Number <> 0 Then Probe = IsDate("{text}") & "|err" & Err.Number '
            f'Else Probe = IsDate("{text}") & "|" & Format(d, "yyyy-mm-dd hh:nn:ss")')


CASES = {
    "typed_array_return": 'Dim a() As Long\na = LongArr()\nProbe = LBound(a) & UBound(a) & "|" & a(2) & "|" & TypeName(a) & "|" & TypeName(LongArr()) & "|" & Join(Words(), "")',
    "function_returns_array": 'Dim v\nv = MakeArr()\nProbe = UBound(v) & "|" & v(1) & "|" & MakeArr()(2) & "|" & TypeName(v)',
    "collection_missing": 'Dim c As New Collection\nOn Error Resume Next\nDim x\nx = c("nope")\nProbe = Err.Number\nErr.Clear\nx = c(1)\nProbe = Probe & "|" & Err.Number\nErr.Clear\nc.Remove 1\nProbe = Probe & "|" & Err.Number',
    "collection_index_errors": 'Dim c As New Collection\nc.Add 1: c.Add 2\nOn Error Resume Next\nDim x\nx = c(3)\nProbe = Err.Number\nErr.Clear\nx = c(0)\nProbe = Probe & "|" & Err.Number\nErr.Clear\nc.Remove 3\nProbe = Probe & "|" & Err.Number\nErr.Clear\nc.Remove "zz"\nProbe = Probe & "|" & Err.Number\nErr.Clear\nx = c(1.6)\nProbe = Probe & "|" & Err.Number & "|" & x',
    "collection_key_types": 'Dim c As New Collection\nOn Error Resume Next\nc.Add "a", 7&\nProbe = Err.Number\nErr.Clear\nc.Add "b", True\nProbe = Probe & "|" & Err.Number\nErr.Clear\nc.Add "c", 1.5\nProbe = Probe & "|" & Err.Number\nErr.Clear\nDim k As Variant\nk = "v"\nc.Add "d", k\nProbe = Probe & "|" & Err.Number & "|" & c.Count\nErr.Clear\nc.Add "e", ""\nProbe = Probe & "|" & Err.Number & "|" & c.Count',
    "collection_key_empty_null_date": 'Dim c As New Collection\nOn Error Resume Next\nc.Add "a", Empty\nProbe = Err.Number & "|" & c.Count\nErr.Clear\nc.Add "b", Null\nProbe = Probe & "|" & Err.Number & "|" & c.Count\nErr.Clear\nDim d As Date\nc.Add "c", d\nProbe = Probe & "|" & Err.Number',
    "collection_before_after": 'Dim c As New Collection\nc.Add 1\nOn Error Resume Next\nc.Add 2, , 5\nProbe = Err.Number\nErr.Clear\nc.Add 2, , "nokey"\nProbe = Probe & "|" & Err.Number\nErr.Clear\nc.Add 3, , 1, 1\nProbe = Probe & "|" & Err.Number & "|" & c.Count',
    "round_values": 'Dim s As String\n' + "\n".join(f's = s & CStr(Round({v}, {d})) & ";"' for v, d in ROUND_VALUES) + "\nProbe = s",
    "negative_zero": 'Dim x As Double\nx = -0.5\nProbe = Fix(x) & "|" & CStr(-0#) & "|" & (0 * -1) & "|" & Fix(-0.2) & "|" & Int(-0#) & "|" & Format(Fix(-0.5)) & "|" & Str(Fix(-0.5)) & "|" & CSng(-0.25 * 0)',
    "negative_zero_ops": 'Dim z As Double\nz = Fix(-0.5)\nProbe = (z = 0) & "|" & Sgn(z) & "|" & (z + 0) & "|" & (z * 1) & "|" & CLng(z) & "|" & TypeName(z) & "|" & Abs(z)',
    "zero_division_kinds": 'On Error Resume Next\nDim v, n As Long\nv = 0 / 0\nProbe = Err.Number\nErr.Clear\nv = 0# / 0#\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = 0 \\ 0\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = 0 Mod 0\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = CDec(0) / CDec(0)\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = CCur(0) / 0\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = n / n\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = 1 / CDec(0)\nProbe = Probe & "|" & Err.Number\nErr.Clear\nv = -1 / 0\nProbe = Probe & "|" & Err.Number',
    "val_hex_widths": 'Probe = Val("&HFFFF") & "|" & Val("&H7FFF") & "|" & Val("&H8000") & "|" & Val("&H10000") & "|" & Val("&HFFFFFFFF") & "|" & Val("&H80000000") & "|" & Val("&o177777") & "|" & Val("&HFFFF&") & "|" & TypeName(Val("&HFFFF"))',
    "ascw_signed": 'Probe = AscW(ChrW(&HFFFF)) & "|" & AscW(ChrW(40000)) & "|" & AscW(ChrW(32767)) & "|" & TypeName(AscW("a")) & "|" & Asc(ChrW(8364)) & "|" & Asc(ChrW(40000))',
    "isdate_numbers": 'Dim d As Date\nd = 45000\nProbe = IsDate(45000) & IsDate(45000.5) & IsDate(d) & IsDate(CVar(d)) & IsDate("45000") & IsDate(CCur(1)) & IsDate(True) & IsDate(Empty) & IsDate(Null) & IsDate("1.5") & IsDate("2024")',
    "formatnumber_args": 'Probe = FormatNumber(0.5, 2, vbFalse) & "|" & FormatNumber(-0.5, 0, , vbTrue) & "|" & FormatNumber(1234567.891, 1, , , vbFalse) & "|" & FormatNumber(-1234.5, 2, vbTrue, vbFalse, vbTrue) & "|" & FormatNumber(0.25) & "|" & FormatNumber(-0.25, -1, vbFalse, vbTrue) & "|" & FormatNumber(1234.5, 0)',
    "formatpercent_args": 'Probe = FormatPercent(0.005, 2, vbFalse) & "|" & FormatPercent(-0.5, 0, , vbTrue) & "|" & FormatPercent(12.345, 1, , , vbFalse) & "|" & FormatPercent(0.5)',
    "formatcurrency_args": 'Probe = FormatCurrency(-12.5) & "|" & FormatCurrency(-12.5, 1, , vbFalse) & "|" & FormatCurrency(0.5, 2, vbFalse) & "|" & FormatCurrency(1234567, 0, , , vbFalse) & "|" & FormatCurrency(1234.5, -1)',
    "formatdatetime_long_date": 'Probe = FormatDateTime(#3/5/2024 1:02:03 PM#, vbLongDate) & "|" & FormatDateTime(#1/2/2024#, vbLongDate) & "|" & FormatDateTime(#12/25/2024 11:05:00 PM#)',
    "datevalue_timevalue_dotted": 'On Error Resume Next\nProbe = Format(DateValue("1.5"), "yyyy-mm-dd") & "|" & Err.Number\nErr.Clear\nProbe = Probe & "|" & Format(TimeValue("1.5"), "hh:nn:ss") & "|" & Err.Number',
}
CASES.update({f"date_text_{index:02d}": date_text(text) for index, text in enumerate(DATE_TEXTS)})


def macro(name: str) -> str:
    return HELPERS + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
