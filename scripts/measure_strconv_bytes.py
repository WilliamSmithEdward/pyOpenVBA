"""What StrConv(text, vbFromUnicode) gives: a String whose bytes are the text's ANSI bytes, not a Byte array, which is
what it looks like until it lands in one. An odd count of bytes leaves half a character.

Each statement's value is its TypeName and text, with every character's code in hex, or its error.

    python scripts/measure_strconv_bytes.py

writes tests/fixtures/strconv_bytes.json, which tests/test_strconv_bytes.py replays.
"""
from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/strconv_bytes.json'
#: What each case sets up, then what it reads. v is a Variant holding the conversion of abc, w of abcd, and e of "".
READS = [
    'TypeName(v)', 'VarType(v)', 'IsArray(v)', 'Len(v)', 'LenB(v)', 'v', 'UBound(v)', 'LBound(v)', 'v(0)', 'AscB(v)', 'AscW(v)', 'Asc(v)',
    'AscB(MidB(v, 1, 1))', 'AscB(MidB(v, 2, 1))', 'AscB(MidB(v, 3, 1))', 'LenB(MidB(v, 1, 1))', 'LenB(MidB(v, 2))', 'MidB(v, 1, 2)',
    'LeftB(v, 2)', 'RightB(v, 2)', 'LenB(LeftB(v, 1))', 'LenB(RightB(v, 1))', 'Mid(v, 1, 1)', 'Left(v, 1)', 'Right(v, 1)', 'Len(Mid(v, 2))',
    'StrConv(v, vbUnicode)', 'LenB(StrConv(v, vbUnicode))', 'LenB(v & "")', 'LenB(v & v)', 'Len(v & v)', 'LenB("x" & v)', 'v = v', 'v = "abc"',
    'v & "x"', 'InStr(v, "c")', 'InStrB(v, "c")', 'InStrB(v, ChrB(99))', 'CStr(v)', 'LenB(CStr(v))', 'LCase(v)', 'LenB(LCase(v))', 'UCase(v)',
    'TypeName(StrConv("abc", vbFromUnicode))', 'LenB(StrConv("abc", vbFromUnicode))', 'UBound(StrConv("abc", vbFromUnicode))',
    'StrConv("abc", vbFromUnicode)', 'StrConv(StrConv("abc", vbFromUnicode), vbUnicode)', 'LenB(StrConv("", vbFromUnicode))',
    'TypeName(StrConv("", vbFromUnicode))', 'TypeName(w)', 'Len(w)', 'LenB(w)', 'w', 'TypeName(e)', 'LenB(e)', 'e = ""', 'IsEmpty(e)',
    'StrConv(ChrW(233) & ChrW(8364) & ChrW(1234), vbFromUnicode)', 'LenB(StrConv(ChrW(233) & ChrW(1234), vbFromUnicode))',
    'StrConv(ChrB(99), vbUnicode)', 'LenB(ChrB(99))', 'TypeName(ChrB(99))', 'ChrB(99)', 'LenB(ChrB(99) & ChrB(100))', 'ChrB(99) & ChrB(100)',
    'LenB(MidB("abc", 1, 1))', 'MidB("abc", 1, 1)', 'LenB(MidB("abc", 1, 3))', 'LeftB("abc", 3)', 'TypeName(MidB("abc", 1, 1))',
]
STEPS = {
    'into a byte array': ['Dim b() As Byte', 'b = v', 'Say TypeName(b)', 'Say UBound(b)', 'Say b(0)', 'Say b(2)', 'Say StrConv(b, vbUnicode)',
                          'Say TypeName(StrConv(b, vbUnicode))'],
    'even into a byte array': ['Dim b() As Byte', 'b = w', 'Say UBound(b)', 'Say b(3)'],
    'empty into a byte array': ['Dim b() As Byte', 'b = e', 'Say LBound(b)', 'Say UBound(b)'],
    'direct into a byte array': ['Dim b() As Byte', 'b = StrConv("abc", vbFromUnicode)', 'Say UBound(b)', 'Say b(1)'],
    'bytes back': ['Dim b() As Byte', 'b = v', 'Dim x', 'x = b', 'Say TypeName(x)', 'Say StrConv(x, vbUnicode)', 'Say LenB(x)', 'Say Len(x)'],
    'variant of bytes': ['Dim b() As Byte', 'b = "abc"', 'Dim x', 'x = b', 'Say TypeName(x)', 'Say UBound(x)', 'x = StrConv(x, vbFromUnicode)', 'Say TypeName(x)'],
    'into a string': ['Dim t As String', 't = v', 'Say TypeName(t)', 'Say LenB(t)', 'Say Len(t)', 'Say t'],
    'through a string': ['Dim t As String', 't = v', 'Say StrConv(t, vbUnicode)', 't = t & "x"', 'Say LenB(t)'],
    'into a fixed string': ['Dim t As String * 3', 't = v', 'Say LenB(t)', 'Say Len(t)', 'Say t'],
    'in a function': ['Say LenB(Passed(v))', 'Say TypeName(Passed(v))'],
    'into a cell': ['Range("A1").Value = v', 'Say Range("A1").Value', 'Say LenB(Range("A1").Value)'],
    'for each': ['Dim one', 'For Each one In v', 'Say one', 'Next'],
    'redim': ['ReDim Preserve v(1)'],
    'join': ['Say Join(v)'],
    'array of': ['Dim x', 'x = Array(v)', 'Say LenB(x(0))'],
    'chrb concat': ['Dim t As String', 't = ChrB(97)', 'Say LenB(t)', 't = t & ChrB(98)', 'Say LenB(t)', 'Say t', 'Say StrConv(t, vbUnicode)'],
    'byte string compare': ['Say StrComp(v, "abc", vbBinaryCompare)', 'Say StrComp(v, v, vbBinaryCompare)', 'Say v < "b"'],
    'put': ['Dim t As String', 't = v', 'Say LenB(t & t)', 'Say Mid(t & t, 2)'],
}


def _lines() -> list[str]:
    lines = [r'''
Private Function Hexed(ByVal value As Variant) As String
    Dim i As Long, out As String
    If VarType(value) <> vbString Then
        Hexed = TypeName(value) & " " & CStr(value)
        Exit Function
    End If
    For i = 1 To Len(value)
        out = out & " " & Hex$(AscW(Mid$(value, i, 1)) And &HFFFF&)
    Next
    Hexed = "String" & out & " /" & LenB(value)
End Function

Private Function Passed(ByVal value As String) As String
    Passed = value
End Function
''']
    cases = [(read, [f'Say {read}']) for read in READS] + list(STEPS.items())
    for index, (_, body) in enumerate(cases):
        lines += [f'Private Function Case{index}() As String', '    Dim s As String, r As Variant, v As Variant, w As Variant, e As Variant',
                  '    On Error Resume Next', '    v = StrConv("abc", vbFromUnicode)', '    w = StrConv("abcd", vbFromUnicode)',
                  '    e = StrConv("", vbFromUnicode)']
        for text in body:
            if text.startswith(('Dim ', 'For Each', 'Next')):
                lines.append(f'    {text}')
            elif text.startswith('Say '):
                lines.append(f'    Err.Clear: r = {text[4:]}: If Err.Number Then s = s & "|E" & Err.Number Else s = s & "|" & Hexed(r)')
            else:
                lines.append(f'    Err.Clear: {text}: If Err.Number Then s = s & "|E" & Err.Number')
        lines += [f'    Case{index} = s', 'End Function', '']
    lines += ['Public Function Probe(ByVal first As Long, ByVal last As Long) As String', '    Dim i As Long, s As String', '    For i = first To last',
              '        Select Case i']
    lines += [f'            Case {index}: s = s & "@@@" & Case{index}()' for index in range(len(cases))]
    lines += ['        End Select', '    Next', '    Probe = Mid$(s, 4)', 'End Function']
    return lines


NAMES = [*READS, *STEPS]


def main() -> None:
    source = '\n'.join(_lines()) + '\n'
    answers: list[str] = []
    with ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for first in range(0, len(NAMES), 40):
            result = excel.run_vba(source, 'Probe', (first, min(first + 39, len(NAMES) - 1)), timeout=120.0, line_numbers=False)
            assert result.ok and not result.dialogs, result
            answers += str(result.value).split('@@@')
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'cases': [list(pair) for pair in zip(NAMES, answers, strict=True)]}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer in record['cases']:
        print(f'{name:58} {answer}')


if __name__ == '__main__':
    main()
