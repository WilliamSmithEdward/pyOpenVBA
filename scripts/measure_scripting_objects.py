"""VBScript.RegExp and FileSystemObject's path functions, as CreateObject makes them.

RegExp is computation alone: Test, Replace and Execute over patterns,
flags, replacement codes, sub-matches, empty matches, Unicode and bad
patterns. Of FileSystemObject only the functions that take a path apart or
put one together are measured, which read nothing from the disk; what
reads or writes files stays outside the model.

    python scripts/measure_scripting_objects.py

writes tests/fixtures/scripting_objects.json, which
tests/test_scripting_objects.py replays. Each probe answers one line per
case, %% apart: the harness hands back Chr(30) as a space.
"""
from __future__ import annotations

from pathlib import Path

from vba_probe_fixture import record

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/scripting_objects.json'
#: The cases: (pattern, flags g/i/m, method, text as a VBA expression, replacement as a VBA expression).
REGEXP = [
    (r'\d+', '', 'test', '"abc123"', '""'),
    (r'^abc$', '', 'test', '"abc" & vbLf & "abc"', '""'),
    (r'^abc$', 'm', 'test', '"abc" & vbLf & "abc"', '""'),
    ('ABC', 'i', 'test', '"xabcx"', '""'),
    ('ABC', '', 'test', '"xabcx"', '""'),
    ('', '', 'test', '"anything"', '""'),
    ('(', '', 'test', '"x"', '""'),
    ('a{2,}', '', 'test', '"caaat"', '""'),
    (r'(\w+)@(\w+)', '', 'replace', '"joe@site and ann@web"', '"$2:$1"'),
    (r'(\w+)@(\w+)', 'g', 'replace', '"joe@site and ann@web"', '"$2:$1"'),
    (r'\s+', 'g', 'replace', '"a  b   c"', '" "'),
    ('o', 'g', 'replace', '"foo"', '"$&$&"'),
    ('b', '', 'replace', '"abc"', '"[$`|$\']"'),
    ('b', '', 'replace', '"abc"', '"$$"'),
    ('(a)(b)?', 'g', 'replace', '"ab a"', '"<$2>"'),
    ('x', 'g', 'replace', '"xxx"', '""'),
    ('a*', 'g', 'replace', '"baaac"', '"-"'),
    ('$', 'g', 'replace', '"abc"', '"!"'),
    ('^', 'gm', 'replace', '"a" & vbLf & "b"', '"> "'),
    (r'(\d)(\d)', 'g', 'replace', '"1234"', '"$2$1"'),
    (r'\$(\d+)', 'g', 'replace', '"cost $15 and $7"', '"USD$1"'),
    ('(?:ab)+', '', 'replace', '"ababab c"', '"X"'),
    ('a(?=b)', 'g', 'replace', '"ab ac ab"', '"X"'),
    ('a(?!b)', 'g', 'replace', '"ab ac ab"', '"X"'),
    ('[^a-z]', 'g', 'replace', '"aB3c!"', '"_"'),
    (r'\bcat\b', 'gi', 'replace', '"Cat concat cat."', '"dog"'),
    (r'(\w+)\s(\w+)', '', 'replace', '"John Smith"', '"$2, $1"'),
    ('a', 'g', 'replace', '"banana"', '"$0"'),
    ('(a)', 'g', 'replace', '"banana"', '"$10"'),
    (r'\.', 'g', 'replace', '"a.b.c"', '"\\"'),
    ('.', 'g', 'replace', '"a" & vbCrLf & "b"', '"x"'),
    ('(?=b)', 'g', 'replace', '"abab"', '"|"'),
    ('a|b|', 'g', 'replace', '"ac"', '"-"'),
    ('(x)(y)(z)', '', 'replace', '"xyz"', '"$3$2$1$4"'),
    (r'\d+', 'g', 'execute', '"a1 b22 c333"', '""'),
    (r'\d+', '', 'execute', '"a1 b22"', '""'),
    (r'(\w)(\d)?', 'g', 'execute', '"a1 b c3"', '""'),
    ('x*', 'g', 'execute', '"axxb"', '""'),
    (r'^\w+', 'gm', 'execute', '"one" & vbLf & "two" & vbCrLf & "three"', '""'),
    (r'(?:\d{3})-(\d{4})', 'g', 'execute', '"555-1234 and 666-9876"', '""'),
    ('[A-Z]+', 'gi', 'execute', '"abc DEF"', '""'),
    ('\u00e9', 'gi', 'execute', 'ChrW(201) & ChrW(233)', '""'),
    (r'\w+', 'g', 'execute', '"na" & ChrW(239) & "ve caf" & ChrW(233)', '""'),
    ('', 'g', 'execute', '"abc"', '""'),
    ('q', 'g', 'execute', '"abc"', '""'),
    ('(a)|(b)', 'g', 'execute', '"ab"', '""'),
    (r'\t\d', 'g', 'execute', '"a" & vbTab & "1"', '""'),
    (r'[\d.]+', 'g', 'execute', '"v1.2.3 and 4.5"', '""'),
    (r'(\w)\1', 'g', 'execute', '"aabbcd"', '""'),
    (r'\x41', 'g', 'execute', '"zAz"', '""'),
    (r'\u0041', 'g', 'execute', '"zAz"', '""'),
    (r'\s', 'g', 'execute', 'ChrW(160) & " " & vbTab', '""'),
    (r'\d', 'g', 'execute', 'ChrW(1633) & "7"', '""'),
    ('[', '', 'execute', '"x"', '""'),
    ('a{,2}', 'g', 'execute', '"a{,2}"', '""'),
    ('(?<=a)b', 'g', 'execute', '"ab"', '""'),
]
#: The paths FileSystemObject takes apart: base name, extension, file name, parent folder and drive of each.
PATHS = ['C:\\a\\b\\report.final.xlsx', 'report.xlsx', 'C:\\', 'C:', '\\\\server\\share\\dir\\file.txt', 'C:\\a\\b\\',
         'C:\\a\\b', 'a/b/c.txt', 'noext', '.gitignore', 'file.', '', 'C:\\a\\..\\b.txt', 'x.tar.gz', 'C:\\a\\b\\\\',
         '\\\\server\\share', 'D:file.txt', ' spaced name .txt ', 'C:\\b.txt', '\\b.txt', '/a.txt',
         '\\\\server\\share\\x.txt', 'C:/a/b.txt', 'C:\\a.b\\c', 'a\\', '\\']
#: The folders and names BuildPath joins.
JOINS = [('C:\\a', 'b.txt'), ('C:\\a\\', 'b.txt'), ('C:\\a', '\\b.txt'), ('C:\\a\\', '\\b.txt'), ('', 'b'), ('a', ''),
         ('C:', 'b'), ('\\\\server\\share', 'x'), ('a/b', 'c'), ('a/', 'c'), ('C:\\a\\\\', '\\\\b')]

HELPERS = '''Private Function Describe(m As Object) As String
    Dim k As Long, out As String
    out = m.FirstIndex & "," & m.Length & "," & m.Value & "["
    For k = 0 To m.SubMatches.Count - 1
        If IsEmpty(m.SubMatches(k)) Then out = out & "<empty>" Else out = out & m.SubMatches(k)
        If k < m.SubMatches.Count - 1 Then out = out & ","
    Next
    Describe = out & "]"
End Function

Private Function Run(pattern As String, flags As String, method As String, text As String, repl As String) As String
    Dim re As Object, ms As Object, m As Object, out As String
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = InStr(flags, "g") > 0
    re.IgnoreCase = InStr(flags, "i") > 0
    re.MultiLine = InStr(flags, "m") > 0
    On Error Resume Next
    Select Case method
        Case "test": out = CStr(re.Test(text))
        Case "replace": out = re.Replace(text, repl)
        Case "execute"
            Set ms = re.Execute(text)
            If Err.Number = 0 Then
                out = ms.Count & ":"
                For Each m In ms
                    out = out & Describe(m) & ";"
                Next
            End If
    End Select
    If Err.Number <> 0 Then out = "E" & Err.Number
    Run = out
End Function
'''


def _vba(text: str) -> str:
    """A VBA expression for ``text``, ChrW for what is past ASCII."""
    pieces: list[str] = []
    plain = ''
    for char in text:
        if 32 <= ord(char) < 127:
            plain += char.replace('"', '""')
            continue
        if plain:
            pieces.append(f'"{plain}"')
            plain = ''
        pieces.append(f'ChrW({ord(char)})')
    if plain or not pieces:
        pieces.append(f'"{plain}"')
    return ' & '.join(pieces)


def regexp_probe(cases: list[tuple[str, str, str, str, str]]) -> str:
    lines = '\n'.join(f's = s & "%%" & Run({_vba(pattern)}, "{flags}", "{method}", {text}, {repl})'
                      for pattern, flags, method, text, repl in cases)
    return f'''{HELPERS}
Function Probe() As String
Dim s As String
{lines}
Probe = Mid$(s, 3)
End Function
'''


#: What a new RegExp holds, its collections' edges and the errors they raise.
OBJECT_PROBE = '''Function Probe() As String
Dim re As Object, ms As Object, s As String
Set re = CreateObject("VBScript.RegExp")
s = "[" & re.Pattern & "]" & re.Global & re.IgnoreCase & re.MultiLine & TypeName(re)
re.Pattern = "(\\d)(x)?": re.Global = True
Set ms = re.Execute("1x 2")
s = s & "%%" & TypeName(ms) & "," & TypeName(ms(0)) & "," & TypeName(ms(0).SubMatches) & "," & ms(0) & "," & ms.Item(1).Value
s = s & "%%" & ms(1).SubMatches.Count & "," & IsEmpty(ms(1).SubMatches(1)) & "," & TypeName(ms(1).SubMatches(1))
On Error Resume Next
Dim v As Variant
v = ms(2)
s = s & "%%" & Err.Number
Err.Clear
v = ms(0).SubMatches(5)
s = s & "%%" & Err.Number
Err.Clear
v = re.Test(Null)
s = s & "%%" & Err.Number & "," & v
Err.Clear
re.Pattern = "a"
v = re.Replace("aaa", Null)
s = s & "%%" & Err.Number & "," & v
Err.Clear
v = re.Replace(123, 4)
s = s & "%%" & Err.Number & "," & v & "," & TypeName(v)
Probe = s
End Function
'''


def fso_probe() -> str:
    lines = '\n'.join(f's = s & "%%" & o.GetBaseName({_vba(path)}) & "|" & o.GetExtensionName({_vba(path)}) & "|" & '
                      f'o.GetFileName({_vba(path)}) & "|" & o.GetParentFolderName({_vba(path)}) & "|" & '
                      f'o.GetDriveName({_vba(path)})' for path in PATHS)
    joins = '\n'.join(f's = s & "%%" & o.BuildPath({_vba(folder)}, {_vba(name)})' for folder, name in JOINS)
    return f'''Function Probe() As String
Dim o As Object, s As String
Set o = CreateObject("Scripting.FileSystemObject")
{lines}
{joins}
s = s & "%%" & TypeName(o)
Probe = Mid$(s, 3)
End Function
'''


def main() -> None:
    half = len(REGEXP) // 2
    record({'regexp_first': regexp_probe(REGEXP[:half]), 'regexp_second': regexp_probe(REGEXP[half:]),
            'regexp_object': OBJECT_PROBE, 'fso_paths': fso_probe()}, OUT)


if __name__ == '__main__':
    main()
