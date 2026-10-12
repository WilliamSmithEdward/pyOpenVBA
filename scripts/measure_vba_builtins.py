"""Measure native VBA built-ins: StrConv, Error, LSet/RSet, numeric text, unallocated
and Byte arrays, and VBA's Collection.Count.

Each probe is the body of ``Function Probe() As String``; the fixture records its
value, or the number of the run-time error it stops with.
"""
from __future__ import annotations
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/vba_builtins.json'

# Whether each text is a number: IsNumeric, CDbl, arithmetic and CLng, which agree.
NUMBER_TEXTS = ["1,000", ",5", "5,", "1,2", "$", "(5", "5)", "(-5)", "-(5)", "$-5", "-$5", "5-", "5 -", "+5", "5+",
                "- 5", "$ 5", "5$", "5 $", "$$5", "1e3", "(1e3)", "1,000.5", "1.000,5", "  5  ", "5%", "€5", "1d2",
                "&H10", "(&H10)", "$&H10", "1,0e2", "--5", "5--", "(5)-", "1 000", "1.", ".5", "-.5", "5.-", "($5)",
                "$(5)", "(5$)", "1,000,000", "1e3,0", "5-$", "$5-", " (5) ", "(+5)", "+(5)", "\t5", "5\t"]
HEAD = r'''
Private Function Try(ByVal s As String) As String
Dim r As String
r = IIf(IsNumeric(s), "T", "F") & "|"
On Error Resume Next
Dim d As Double
d = CDbl(s)
If Err.Number Then r = r & "e" & Err.Number Else r = r & CStr(d)
Err.Clear
r = r & "|"
Dim v As Variant
v = s + 0
If Err.Number Then r = r & "e" & Err.Number Else r = r & CStr(v) & ":" & TypeName(v)
Err.Clear
r = r & "|"
Dim l As Long
l = CLng(s)
If Err.Number Then r = r & "e" & Err.Number Else r = r & CStr(l)
Try = r
End Function
'''


def literal(text: str) -> str:
    return "".join(f'" & ChrW({ord(c)}) & "' if ord(c) > 126 or c == "\t" else c.replace('"', '""') for c in text)


WORD_BREAKS = {"space": "Chr(32)", "tab": "Chr(9)", "nul": "Chr(0)", "cr": "Chr(13)", "lf": "Chr(10)", "dash": '"-"',
               "dot": '"."', "comma": '","', "apos": '"\'"', "digit": '"1"', "under": '"_"', "nbsp": "ChrW(160)",
               "ff": "Chr(12)", "vt": "Chr(11)", "paren": '"("', "slash": '"/"'}

CASES: dict[str, str] = {
    f"proper_after_{name}": f'Dim s As String\ns = StrConv("ab" & {char} & "cd", vbProperCase)\n'
                            'Probe = Asc(Mid$(s, 1, 1)) & "," & Asc(Mid$(s, Len(s) - 1, 1))'
    for name, char in WORD_BREAKS.items()
}
CASES.update({
    "strconv_cases": 'Probe = StrConv("hELLO wORLD-x y", vbProperCase) & "|" & StrConv("abc", vbUpperCase) & "|" & StrConv("AbC", vbLowerCase) & "|" & StrConv("a b", 3)',
    "strconv_proper_accents": 'Dim s As String\ns = StrConv(ChrW(233) & "lan " & ChrW(201) & "COLE " & ChrW(223) & "a", vbProperCase)\nProbe = AscW(Mid$(s, 1, 1)) & "," & AscW(Mid$(s, 6, 1)) & "," & Mid$(s, 7, 4) & "," & AscW(Mid$(s, 12, 1))',
    "strconv_proper_leading": 'Probe = "[" & StrConv("  ab", vbProperCase) & "][" & StrConv("", vbProperCase) & "]"',
    "strconv_from_unicode": 'Dim b() As Byte\nb = StrConv("AB", vbFromUnicode)\nProbe = UBound(b) & "|" & b(0) & "|" & StrConv(b, vbUnicode)',
    "strconv_from_unicode_unmappable": 'Dim b() As Byte\nb = StrConv(ChrW(8364) & ChrW(&H4E2D) & ChrW(233), vbFromUnicode)\nProbe = UBound(b) & "|" & b(0) & "," & b(1) & "," & b(2)',
    "strconv_unicode_bytes": 'Dim b(2) As Byte\nb(0) = 65: b(1) = 128: b(2) = 233\nProbe = AscW(Mid$(StrConv(b, vbUnicode), 2, 1)) & "," & AscW(Mid$(StrConv(b, vbUnicode), 3, 1))',
    "strconv_null": 'Probe = IsNull(StrConv(Null, vbUpperCase))',
    "strconv_number": 'Probe = StrConv(12.5, vbUpperCase) & "|" & TypeName(StrConv(12.5, vbUpperCase))',
    "strconv_flag_4": 'Dim s As String\ns = StrConv("a", 4)\nProbe = s',
    "strconv_case_with_bytes": 'Dim b() As Byte\nb = StrConv("ab", vbUpperCase + vbFromUnicode)\nProbe = b(0)',
    "strconv_locale": 'Probe = StrConv("ab", vbUpperCase, 1033)',
    "error_of_number": 'Probe = Error(11) & "|" & Error$(9) & "|" & Error(13) & "|" & Error$(1000) & "|" & Error(0) & "|" & TypeName(Error(5))',
    "error_numbers": 'Probe = Error(3) & "|" & Error(20) & "|" & Error(52) & "|" & Error(91) & "|" & Error(424) & "|" & Error(1004)',
    "error_current": 'On Error Resume Next\nErr.Raise 6\nProbe = Error() & "|" & Error',
    "error_after_clear": 'On Error Resume Next\nErr.Raise 11\nErr.Clear\nProbe = "[" & Error & "]"',
    "error_negative": 'Probe = Error(-1)',
    "error_overflow": 'Dim s As String\ns = Error(70000)\nProbe = s',
    "error_text_argument": 'Probe = Error("11")',
    "lset_rset_strings": 'Dim s As String, t As String\ns = "abcdef"\nLSet s = "xy"\nt = "123456"\nRSet t = "xy"\nProbe = "[" & s & "][" & t & "]"',
    "lset_rset_truncate": 'Dim s As String, t As String\ns = "abc"\nLSet s = "wxyz12"\nt = "abc"\nRSet t = "wxyz12"\nProbe = "[" & s & "][" & t & "]"',
    "lset_fixed_string": 'Dim f As String * 5\nf = "12345"\nLSet f = "ab"\nProbe = "[" & f & "]"',
    "rset_number": 'Dim s As String\ns = "abcde"\nRSet s = 12\nProbe = "[" & s & "]"',
    "lset_variant": 'Dim v As Variant\nv = "abcd"\nLSet v = "x"\nProbe = "[" & v & "]"',
    "rset_empty_target": 'Dim s As String\nRSet s = "abc"\nProbe = "[" & s & "]" & Len(s)',
    "lset_null": 'On Error Resume Next\nDim s As String\ns = "abc"\nLSet s = Null\nProbe = Err.Number & "[" & s & "]"',
    "unallocated_ubound": 'Dim a() As Long\nProbe = UBound(a)',
    "unallocated_lbound": 'Dim a() As String\nProbe = LBound(a)',
    "unallocated_index": 'Dim a() As Long\nProbe = a(0)',
    "unallocated_after_erase": 'Dim a() As Long\nReDim a(3)\nErase a\nProbe = UBound(a)',
    "unallocated_for_each": 'Dim a() As Long, v As Variant\nFor Each v In a\nNext\nProbe = "done"',
    "unallocated_to_variant": 'Dim a() As Long, v As Variant\nv = a\nOn Error Resume Next\nDim n As Long\nn = UBound(v)\nProbe = TypeName(v) & "|" & Err.Number',
    "unallocated_join": 'Dim a() As String\nProbe = "[" & Join(a, ",") & "]"',
    "unallocated_isarray": 'Dim a() As Long\nProbe = IsArray(a) & "|" & TypeName(a)',
    "unallocated_redim_preserve": 'Dim a() As Long\nReDim Preserve a(2)\nProbe = UBound(a)',
    "erase_then_redim": 'Dim a() As Long\nReDim a(2)\nErase a\nReDim a(1)\nProbe = UBound(a)',
    "erase_fixed": 'Dim a(3) As Long\na(1) = 5\nErase a\nProbe = UBound(a) & "|" & a(1)',
    "empty_split_and_array": 'Dim a As Variant, b As Variant\na = Split("")\nb = Array()\nProbe = LBound(a) & "|" & UBound(a) & "|" & LBound(b) & "|" & UBound(b)',
    "byte_array_from_string": 'Dim b() As Byte\nb = "AB"\nProbe = UBound(b) & "|" & b(0) & "," & b(1) & "," & b(2)',
    "string_from_byte_array": 'Dim b(3) As Byte, s As String\nb(0) = 72: b(2) = 105\ns = b\nProbe = s & "|" & Len(s)',
    "collection_count_method": 'Dim c As New Collection\nc.Add 5\nProbe = CallByName(c, "Count", VbMethod) & "|" & c.Count() & "|" & c.Count',
    "collection_count_get": 'Dim c As New Collection\nc.Add 5\nProbe = CallByName(c, "Count", VbGet)',
})
CASES.update({f"number_text_{index:02d}": f'Probe = Try("{literal(text)}")' for index, text in enumerate(NUMBER_TEXTS)})


def macro(name: str) -> str:
    return HEAD + '\nFunction Probe() As String\n' + CASES[name] + '\nEnd Function\n'


def main() -> None:
    probes = []
    version_value = None
    excel: ExcelSession | None = None
    try:
        for name in CASES:
            if excel is None:
                excel = ExcelSession(HarnessConfig(lock_wait_s=30.0))
                excel.__enter__()
                excel.new_document()
                version = excel.run_vba('Function VersionBuild() As String\nVersionBuild = Application.Version & "|" & Application.Build\nEnd Function', 'VersionBuild')
                assert version.ok and not version.dialogs, version
                version_value = version.value
            result = excel.run_vba(macro(name), 'Probe', timeout=30.0)
            assert not result.dialogs, (name, result)
            if result.ok:
                probe = {'name': name, 'macro': macro(name), 'result': result.value}
            else:
                found = re.search(r'VBA error (\d+)', str(result.error))
                assert found, (name, result)
                probe = {'name': name, 'macro': macro(name), 'error': int(found.group(1))}
                excel.__exit__(None, None, None)
                excel = None
            probes.append(probe)
            print(name, probe.get('result', probe.get('error')), flush=True)
    finally:
        if excel is not None:
            excel.__exit__(None, None, None)
    OUT.write_text(json.dumps({'excel_version_build': version_value, 'measured_at': datetime.now(timezone.utc).isoformat(), 'probes': probes}, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
