"""Measure the string functions at their edges: zero, negative and fractional
lengths and starts, Null and Empty, numbers and Booleans as text, StrConv's
proper case, StrComp, String and Space, Asc and Chr, and Format's named formats.

Each probe runs a group of calls and records TypeName and value of each, or e
and the error number. A call that answers NULs is left out: the harness reads a
NUL back as a space. Each probe is a module whose ``Function Probe() As String``
the fixture records (see vba_probe_fixture.py).
"""
from __future__ import annotations
from pathlib import Path
from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_string_functions.json'

CALLS = {
    'left': [
        'Left("abc", 0)',
        'Left("abc", 5)',
        'Left("abc", -1)',
        'Left(123, 2)',
        'Left(Null, 1)',
        'Left(Empty, 1)',
        'Left("abc", 1.5)',
        'Left("abc", 2.5)',
        'Left(True, 2)',
        'Left(#1/2/2024#, 3)',
    ],
    'mid': [
        'Mid("abc", 0)',
        'Mid("abc", 4)',
        'Mid("abc", 2, 0)',
        'Mid("abc", 2, -1)',
        'Mid("abc", 2, 10)',
        'Mid("abc", 1.5)',
        'Mid(Null, 1)',
        'Mid("abc", 3, 1)',
        'Mid(12345, 2, 2)',
        'Mid("", 1)',
    ],
    'instr': [
        'InStr("abc", "")',
        'InStr("", "")',
        'InStr("", "a")',
        'InStr(0, "abc", "a")',
        'InStr(4, "abc", "")',
        'InStr(5, "abc", "")',
        'InStr(Null, "a")',
        'InStr("a", Null)',
        'InStr(1, Null, "a")',
        'InStr(2, "abab", "ab")',
        'InStr(1, "ABC", "b", 1)',
        'InStr(123, 2)',
        'InStr(1.5, "abc", "b")',
    ],
    'replace': [
        'Replace("abc", "b", "x", 0)',
        'Replace("abc", "b", "x", 4)',
        'Replace("abc", "b", "x", 1, 0)',
        'Replace("abc", "b", "x", 1, -1)',
        'Replace("", "a", "b")',
        'Replace("aaa", "aa", "b")',
        'Replace(Null, "a", "b")',
        'Replace("abc", Null, "b")',
        'Replace(123, 2, 9)',
        'Replace("abc", "b", "x", 2)',
    ],
    'strconv': [
        'StrConv("hello world", vbProperCase)',
        'StrConv("HELLO-WORLD", vbProperCase)',
        'StrConv("o\'neil mc", vbProperCase)',
        'StrConv("abc", vbUpperCase)',
        'StrConv("ABC", vbLowerCase)',
        'StrConv("a1b2", 3)',
        'StrConv("x y", 3)',
        'StrConv(Null, 1)',
    ],
    'trim_case': [
        'Trim("")',
        'Trim(Null)',
        'LCase(123)',
        'UCase(True)',
        'Trim(12)',
        'Len(Null)',
        'Len(Empty)',
        'LenB("ab")',
        'Len(CStr(12))',
    ],
    'strcomp': [
        'StrComp("a", "A")',
        'StrComp("a", "A", 1)',
        'StrComp("", "")',
        'StrComp(Null, "a")',
        'StrComp("a", "b", 0)',
        'StrComp("B", "a", 1)',
        'StrComp(1, "1")',
        'StrComp("10", "9")',
    ],
    'string_space': [
        'String(0, "a")',
        'String(3, "")',
        'String(2, -1)',
        'Space(0)',
        'String(Null, "a")',
        'String(2, Null)',
    ],
    'instrrev': [
        'InStrRev("abcabc", "c")',
        'InStrRev("abcabc", "c", 5)',
        'InStrRev("abc", "", 2)',
        'InStrRev("abc", "x")',
        'InStrRev("abc", "c", -1)',
        'InStrRev("abc", "c", 0)',
        'InStrRev("abc", "a", 10)',
        'InStrRev(Null, "a")',
    ],
    'asc_chr': [
        'Asc("")',
        'Asc("A")',
        'Asc(65)',
        'Asc(Null)',
        'Chr(65.5)',
        'Chr(66.5)',
        'ChrW(65.5)',
        'Chr(-1)',
        'AscW("")',
        'Chr$(97)',
    ],
    'format_misc': [
        'Format(Null)',
        'Format(Empty)',
        'Format("abc")',
        'Format(True)',
        'Format(0.5)',
        'Format("12", "000")',
        'Format("x", "000")',
        'Format(1234.5, "Currency")',
        'Format(0.123, "Percent")',
        'Format(1, "Yes/No")',
        'Format(0, "On/Off")',
        'Format(-1, "True/False")',
        'Format(1234567.891, "Standard")',
        'Format(1234567.891, "Fixed")',
        'Format(1234567.891, "General Number")',
        'Format(0.000012345, "Scientific")',
    ],
}


def body(group: list[str]) -> str:
    lines = ['Dim s As String, r As Variant', 'On Error Resume Next']
    for call in group:
        lines += ['Err.Clear', 'r = Empty', f'r = {call}',
                  'If Err.Number Then s = s & "e" & Err.Number & ";" Else s = s & TypeName(r) & ":" & r & ";"']
    return '\n'.join([*lines, 'Probe = s'])


CASES = {f'strings_{name}': body(group) for name, group in CALLS.items()}


def macro(name: str) -> str:
    return '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    record({name: macro(name) for name in CASES}, OUT)


if __name__ == '__main__':
    main()
