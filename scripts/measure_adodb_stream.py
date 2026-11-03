"""ADODB.Stream in live Excel: binary and text, its charsets, saving and loading, and its errors.

Built like scripts/measure_file_statements.py, and run the same way: each case has its own empty folder, the answer is
what it says and every file it leaves, the folder's path is masked, and the harness's line numbers stay off. st is made
for each case, and o is there for a second one.

    python scripts/measure_adodb_stream.py

writes tests/fixtures/adodb_stream.json, which tests/test_vba_adodb_stream.py replays.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

sys.path.insert(0, str(Path(__file__).resolve().parent))
from measure_file_statements import _HELPERS, _statement, files_of, seed, tidy  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/adodb_stream.json'
STATE = ['Say st.State', 'Say st.Type', 'Say st.Mode', 'Say st.Charset', 'Say st.LineSeparator', 'Say st.Position', 'Say st.Size', 'Say st.EOS']
BIN = ['st.Type = 1', 'st.Open']
BYTES = ['Dim b(2) As Byte', 'b(0) = 1: b(1) = 200: b(2) = 255']
ABC = ['Dim b() As Byte', 'b = StrConv("abcde", vbFromUnicode)']
WHERE = ['Say st.Position', 'Say st.Size', 'Say st.EOS']


def text(charset: str, value: str, *more: str) -> list[str]:
    return [f'st.Charset = "{charset}"', 'st.Open', f'st.WriteText {value}', *WHERE, 'st.Position = 0', 'Say Hx(RawOf(st))', 'st.Position = 0', 'st.Type = 2', f'st.Charset = "{charset}"',
            'Say st.ReadText', 'st.SaveToFile P("a.txt"), 2', *more]


def loaded(data: str, charset: str, *more: str) -> list[str]:
    return [f'@seed a.txt {data}', f'st.Charset = "{charset}"', 'st.Open', 'st.LoadFromFile P("a.txt")', *WHERE, 'Say st.ReadText', *WHERE, *more]


CASES: dict[str, list[str]] = {
    'made': STATE,
    'opened': ['st.Open', *STATE],
    'closed again': ['st.Open', 'st.Close', *STATE],
    'open twice': ['st.Open', 'st.Open'],
    'close twice': ['st.Close', 'st.Open', 'st.Close', 'st.Close'],
    'closed members': ['Say st.Read', 'Say st.ReadText', 'st.Write "a"', 'st.WriteText "a"', 'st.SaveToFile P("a.txt")', 'st.LoadFromFile P("a.txt")', 'st.SetEOS', 'st.SkipLine', 'st.Flush', 'st.Cancel', 'Say st.EOS', 'Say st.Size', 'Say st.Position', 'st.Position = 0'],
    'open with a source': ['st.Open P("a.txt")'],
    'kinds': ['Say TypeName(st)', 'Say TypeName(st.Type)', 'Say TypeName(st.State)', 'Say TypeName(st.Mode)', 'Say TypeName(st.Charset)', 'Say TypeName(st.LineSeparator)', 'st.Open', 'Say TypeName(st.Position)', 'Say TypeName(st.Size)', 'Say TypeName(st.EOS)'],
    'properties': ['st.Type = 1', 'Say st.Type', 'st.Type = 2', 'Say st.Type', 'st.Type = 0', 'st.Type = 3', 'st.Mode = 1', 'Say st.Mode', 'st.Mode = 3', 'Say st.Mode', 'st.Mode = 99', 'st.LineSeparator = 13', 'Say st.LineSeparator', 'st.LineSeparator = 10', 'Say st.LineSeparator', 'st.LineSeparator = 5', 'st.LineSeparator = -1', 'Say st.LineSeparator'],
    'properties open': ['st.Open', 'st.Type = 1', 'Say st.Type', 'st.Mode = 1', 'Say st.Mode', 'st.Charset = "utf-8"', 'Say st.Charset', 'st.Charset = "nonsense"', 'Say st.Charset', 'st.LineSeparator = 10', 'Say st.LineSeparator', 'st.Charset = "UTF-8"', 'Say st.Charset'],
    'state constants': ['Say st.State', 'st.Open', 'Say st.State', 'Say st.Mode'],
    # binary
    'binary': [*BYTES, *BIN, 'st.Write b', *WHERE, 'st.Position = 0', 'Say st.EOS', 'Say Hx(st.Read)', *WHERE, 'Say Hx(st.Read)', 'Say IsNull(st.Read)', 'st.SaveToFile P("a.bin")'],
    'binary pieces': [*ABC, *BIN, 'st.Write b', 'st.Position = 1', 'Say Hx(st.Read(2))', 'Say st.Position', 'Say Hx(st.Read(9))', 'Say st.Position', 'Say st.EOS', 'Say Hx(st.Read(1))', 'st.Position = 0', 'Say Hx(st.Read(0))', 'Say Hx(st.Read(-1))', 'st.Position = 0', 'Say Hx(st.Read(-2))'],
    'binary write in the middle': [*ABC, *BIN, 'st.Write b', 'st.Position = 1', 'Dim c(0) As Byte', 'c(0) = 90', 'st.Write c', 'Say st.Position', 'Say st.Size', 'st.Position = 0', 'Say Hx(st.Read)'],
    'binary position': [*ABC, *BIN, 'st.Write b', 'st.Position = 5', 'Say st.EOS', 'st.Position = 6', 'Say st.Position', 'st.Position = -1', 'Say st.Position'],
    'binary set eos': [*ABC, *BIN, 'st.Write b', 'st.Position = 2', 'st.SetEOS', 'Say st.Size', 'Say st.Position', 'Say st.EOS', 'st.Position = 0', 'Say Hx(st.Read)'],
    'binary text calls': [*BIN, 'st.WriteText "a"', 'Say st.ReadText', 'st.SkipLine', 'Say st.Size'],
    'binary write kinds': [*BIN, 'st.Write "abc"', 'st.Write 5', 'st.Write Null', 'st.Write Empty', 'Dim n(1) As Long', 'st.Write n', 'Say st.Size'],
    'binary write variant': ['Dim v2 As Variant', 'Dim b() As Byte', 'b = StrConv("ab", vbFromUnicode)', 'v2 = b', *BIN, 'st.Write v2', 'Say st.Size'],
    'binary write empty': ['Dim b() As Byte', *BIN, 'st.Write b', 'Say st.Size'],
    'binary flush': [*ABC, *BIN, 'st.Write b', 'st.Flush', 'Say st.Size', 'st.Cancel', 'Say st.Size'],
    'binary kind change': [*ABC, *BIN, 'st.Write b', 'st.Type = 2', 'Say st.Type', 'st.Position = 0', 'st.Type = 2', 'Say st.Type', 'Say st.Size', 'Say st.ReadText', 'st.Position = 0', 'st.Charset = "us-ascii"', 'Say st.ReadText', 'st.Position = 0', 'st.Type = 1', 'Say st.Type'],
    # text
    'text default': text('Unicode', '"abc"'),
    'text utf8': text('utf-8', '"a" & ChrW(233) & ChrW(8364)'),
    'text utf8 upper': text('UTF-8', '"a"'),
    'text latin1': text('iso-8859-1', '"a" & ChrW(233) & ChrW(1234)'),
    'text 1252': text('windows-1252', '"a" & ChrW(233) & ChrW(8364) & ChrW(1234)'),
    'text ascii': text('us-ascii', '"a" & ChrW(233)'),
    'text utf16': text('utf-16', '"a"'),
    'text unicode name': text('unicode', '"a"'),
    'text big endian': text('unicodeFFFE', '"a"'),
    'text ansi': text('_autodetect', '"a"'),
    'text shift jis': text('shift_jis', '"a" & ChrW(12354)'),
    'text x user': text('x-user-defined', '"a"'),
    'text utf8 twice': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a"', 'st.WriteText "b"', *WHERE, 'st.SaveToFile P("a.txt")'],
    'text default twice': ['st.Open', 'st.WriteText "a"', 'st.WriteText "b"', *WHERE, 'st.SaveToFile P("a.txt")'],
    'text line': ['st.Open', 'st.WriteText "a", 1', 'st.WriteText "b", 0', 'st.WriteText "c", 2', *WHERE, 'st.SaveToFile P("a.txt")'],
    'text line lf': ['st.LineSeparator = 10', 'st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a", 1', 'st.SaveToFile P("a.txt")'],
    'text line cr': ['st.LineSeparator = 13', 'st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a", 1', 'st.SaveToFile P("a.txt")'],
    'text empty': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText ""', *WHERE, 'st.SaveToFile P("a.txt")'],
    'text no write': ['st.Charset = "utf-8"', 'st.Open', 'st.SaveToFile P("a.txt")', 'Say st.Size'],
    'text kinds': ['st.Open', 'st.WriteText 5', 'st.WriteText True', 'st.WriteText Null', 'st.WriteText Empty', 'st.Write "a"', 'st.Position = 0', 'Say st.ReadText'],
    'text charset after': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "ab"', 'st.Charset = "Unicode"', 'Say st.Charset', 'st.Position = 0', 'st.Charset = "Unicode"', 'Say st.Charset', 'Say st.ReadText', 'Say st.Size'],
    'text read': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "abc" & vbCrLf & "de" & vbLf & "f" & vbCr & "g"', 'st.Position = 0', 'Say st.ReadText(2)', 'Say st.Position', 'Say st.ReadText(-2)', 'Say st.Position', 'Say st.ReadText(-2)', 'Say st.EOS', 'Say st.ReadText(-1)', 'Say st.EOS', 'Say st.ReadText', 'Say st.ReadText(1)', 'Say st.ReadText(-2)', 'Say st.ReadText(0)', 'st.Position = 0', 'Say st.ReadText(0)', 'Say st.ReadText(-3)', 'Say st.ReadText(100)'],
    'text read lf': ['st.LineSeparator = 10', 'st.Charset = "utf-8"', 'st.Open', 'st.WriteText "abc" & vbCrLf & "de" & vbLf & "f"', 'st.Position = 0', 'Say st.ReadText(-2)', 'Say st.ReadText(-2)', 'Say st.ReadText(-2)', 'Say st.EOS'],
    'text read cr': ['st.LineSeparator = 13', 'st.Charset = "utf-8"', 'st.Open', 'st.WriteText "abc" & vbCrLf & "de" & vbCr & "f"', 'st.Position = 0', 'Say st.ReadText(-2)', 'Say st.ReadText(-2)', 'Say st.ReadText(-2)', 'Say st.EOS'],
    'text read accented': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText ChrW(233) & ChrW(8364) & "a"', 'st.Position = 0', 'Say st.ReadText(1)', 'Say st.Position', 'Say st.ReadText(1)', 'Say st.Position', 'st.Position = 1', 'Say st.ReadText(1)'],
    'text skip line': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "ab" & vbCrLf & "cd" & vbCrLf', 'st.Position = 0', 'st.SkipLine', 'Say st.Position', 'Say st.ReadText(-2)', 'st.SkipLine', 'Say st.EOS', 'st.SkipLine', 'Say st.EOS'],
    'text read position in unicode': ['st.Open', 'st.WriteText "abc"', 'st.Position = 0', 'Say st.ReadText(1)', 'Say st.Position', 'st.Position = 2', 'Say st.ReadText(1)', 'Say st.Position', 'st.Position = 3', 'Say Hx(RawText(st.ReadText(1)))', 'st.Position = 1', 'Say st.ReadText'],
    'text set eos': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "abcd"', 'st.Position = 2', 'st.SetEOS', 'Say st.Size', 'st.Position = 0', 'Say st.ReadText'],
    'text write in the middle': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "abcd"', 'st.Position = 4', 'st.WriteText "Z"', 'Say st.Position', 'Say st.Size', 'st.Position = 0', 'Say st.ReadText'],
    'text read write': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "ab"', 'Say st.ReadText', 'st.WriteText "cd"', 'Say st.Position', 'Say st.Size'],
    'text nul': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a" & Chr(0) & "b"', 'st.Position = 0', 'Say st.ReadText', 'Say st.Size'],
    # saving
    'save': [*ABC, *BIN, 'st.Write b', 'st.SaveToFile P("a.bin")', 'Say st.Position', 'Say st.EOS', 'st.SaveToFile P("a.bin")', 'st.SaveToFile P("a.bin"), 1', 'st.SaveToFile P("a.bin"), 2', 'st.SaveToFile P("b.bin"), 3', 'st.SaveToFile P("c.bin"), 0', 'st.SaveToFile P("d.bin"), 5'],
    'save position': [*ABC, *BIN, 'st.Write b', 'st.Position = 2', 'st.SaveToFile P("a.bin")', 'Say st.Position'],
    'save overwrite shorter': ['@seed a.bin 010203040506', *ABC, 'b = StrConv("a", vbFromUnicode)', *BIN, 'st.Write b', 'st.SaveToFile P("a.bin"), 2'],
    'save no folder': [*BIN, 'st.SaveToFile P("none\\a.bin")', 'st.SaveToFile P("none\\a.bin"), 2'],
    'save over folder': ['MkDir P("d")', *BIN, 'st.SaveToFile P("d"), 2', 'st.SaveToFile P("d")'],
    'save empty name': [*BIN, 'st.SaveToFile ""', 'st.SaveToFile "", 2', 'st.SaveToFile Null'],
    'save read only': ['@seed a.bin 61', 'SetAttr P("a.bin"), vbReadOnly', *BIN, 'st.SaveToFile P("a.bin"), 2', 'SetAttr P("a.bin"), vbNormal'],
    'save open file': [*BIN, 'Open P("a.bin") For Output As #1', 'st.SaveToFile P("a.bin"), 2', 'Close #1'],
    'save url': [*BIN, 'st.SaveToFile "http://x/a", 2', 'st.SaveToFile "ftp://x/a", 2'],
    # loading
    'load binary': ['@seed a.bin 0102FF', *BIN, 'st.LoadFromFile P("a.bin")', *WHERE, 'Say Hx(st.Read)', *WHERE],
    'load replaces': ['@seed a.bin 61', *ABC, *BIN, 'st.Write b', 'st.LoadFromFile P("a.bin")', 'Say st.Size', 'Say st.Position'],
    'load missing': ['st.Open', 'st.LoadFromFile P("none")', 'st.LoadFromFile P("none\\x")', 'st.LoadFromFile ""', 'st.LoadFromFile Null'],
    'load folder': ['MkDir P("d")', 'st.Open', 'st.LoadFromFile P("d")'],
    'load empty': ['@seed a.bin ', *BIN, 'st.LoadFromFile P("a.bin")', 'Say st.Size', 'Say st.EOS'],
    'load open file': ['@seed a.bin 61', *BIN, 'Open P("a.bin") For Output As #1', 'st.LoadFromFile P("a.bin")', 'Say st.Size', 'Close #1'],
    'load text utf8 bom': loaded('EFBBBF61C3A9', 'utf-8'),
    'load text utf8': loaded('61C3A9', 'utf-8'),
    'load text utf8 bad': loaded('61E962', 'utf-8'),
    'load text unicode bom': loaded('FFFE6100', 'Unicode'),
    'load text unicode none': loaded('6100', 'Unicode'),
    'load text unicode odd': loaded('610062', 'Unicode'),
    'load text ansi': loaded('61E9', 'iso-8859-1'),
    'load text 1252': loaded('6180E9', 'windows-1252'),
    'load text ascii': loaded('61E9', 'us-ascii'),
    'load text big endian': loaded('FEFF0061', 'unicodeFFFE'),
    'load text bom other': loaded('FFFE6100', 'utf-8'),
    'load text utf8 bom other': loaded('EFBBBF61', 'Unicode'),
    'load text then write': loaded('61', 'utf-8', 'st.WriteText "b"', 'st.Position = 0', 'Say st.ReadText', 'st.SaveToFile P("b.txt")'),
    'load auto': loaded('EFBBBF61', '_autodetect'),
    'load auto all': loaded('EFBBBF61', '_autodetect_all'),
    # CopyTo
    'copy': [*ABC, *BIN, 'st.Write b', 'st.Position = 1', 'Set o = CreateObject("ADODB.Stream")', 'o.Type = 1', 'o.Open', 'st.CopyTo o', 'Say st.Position', 'Say o.Position', 'Say o.Size', 'o.Position = 0', 'Say Hx(o.Read)'],
    'copy some': [*ABC, *BIN, 'st.Write b', 'st.Position = 0', 'Set o = CreateObject("ADODB.Stream")', 'o.Type = 1', 'o.Open', 'st.CopyTo o, 2', 'st.CopyTo o, 0', 'st.CopyTo o, -1', 'Say o.Size', 'Say st.Position', 'st.CopyTo o, 99', 'Say o.Size'],
    'copy text': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "ab"', 'st.Position = 0', 'Set o = CreateObject("ADODB.Stream")', 'o.Charset = "Unicode"', 'o.Open', 'st.CopyTo o', 'Say o.Size', 'o.Position = 0', 'Say o.ReadText'],
    'copy closed': [*BIN, 'Set o = CreateObject("ADODB.Stream")', 'st.CopyTo o', 'o.CopyTo st'],
    'copy other kinds': [*BIN, 'Set o = CreateObject("ADODB.Stream")', 'o.Open', 'st.CopyTo o', 'o.CopyTo st', 'st.CopyTo Nothing', 'st.CopyTo 5'],
    # the common round trip
    'utf8 no bom': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a"', 'st.Position = 3', 'Set o = CreateObject("ADODB.Stream")', 'o.Type = 1', 'o.Open', 'st.CopyTo o', 'o.SaveToFile P("a.txt")', 'o.Close', 'st.Close', 'Say st.State'],
    'text to bytes': ['st.Charset = "utf-8"', 'st.Open', 'st.WriteText "a" & ChrW(233)', 'st.Position = 0', 'st.Type = 1', 'st.Position = 3', 'Say Hx(st.Read)'],
    'bytes to text': [*BIN, 'Dim b() As Byte', 'b = StrConv("abc", vbFromUnicode)', 'st.Write b', 'st.Position = 0', 'st.Type = 2', 'st.Charset = "utf-8"', 'Say st.ReadText'],
}

_MORE = r'''
Private Function Hx(ByVal value As Variant) As String
    Dim i As Long, out As String
    If IsNull(value) Then
        Hx = "Null"
    ElseIf Not IsArray(value) Then
        Hx = TypeName(value) & " " & Esc(CStr(value))
    Else
        On Error Resume Next
        out = TypeName(value) & " " & LBound(value) & " " & UBound(value) & " "
        If Err.Number Then
            Hx = TypeName(value) & " unbound"
            Exit Function
        End If
        On Error GoTo 0
        For i = LBound(value) To UBound(value)
            out = out & Right$("0" & Hex$(value(i)), 2)
        Next
        Hx = out
    End If
End Function

Private Function RawOf(ByVal st As Object) As Variant
    st.Type = 1
    RawOf = st.Read
End Function

Private Function RawText(ByVal value As String) As Variant
    RawText = AscW(value)
End Function
'''


def _probe() -> str:
    lines = [_HELPERS, _MORE]
    names = list(CASES)
    for index, name in enumerate(names):
        lines += [f'Private Function Case{index}() As String', '    Dim s As String, v As Variant, st As Object, o As Object', '    On Error Resume Next',
                  '    Set st = CreateObject("ADODB.Stream")']
        for one in CASES[name]:
            lines += _statement(one)
        lines += ['    If st.State = 1 Then st.Close', '    If Not o Is Nothing Then', '        If o.State = 1 Then o.Close', '    End If', '    Close', '    Err.Clear',
                  f'    Case{index} = s', 'End Function', '']
    lines += ['Public Function Probe(ByVal folder As String, ByVal sep As String, ByVal which As Long) As String',
              '    gFolder = folder: gSep = sep', '    Select Case which']
    lines += [f'        Case {index}: Probe = Case{index}()' for index in range(len(names))]
    lines += ['    End Select', 'End Function']
    return '\n'.join(lines) + '\n'


def masked(answer: str, folder: Path) -> str:
    for path, word in ((str(folder), 'FOLDER'), (str(folder.parent), 'ROOT')):
        answer = answer.replace(path, word).replace(path.replace('\\', '\\u005C'), word)
    return answer.replace('\\u005C', '/')


def main() -> None:
    source = _probe()
    names = list(CASES)
    answers: list[list[str]] = []
    with tempfile.TemporaryDirectory() as root, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, name in enumerate(names):
            folder = Path(root) / f'c{index}'
            folder.mkdir()
            seed(folder, CASES[name])
            result = excel.run_vba(source, 'Probe', (str(folder), '\\', index), timeout=60.0, line_numbers=False)
            if not result.ok or result.dialogs:
                message = result.dialogs[0].message if result.dialogs else result.outcome
                print(f'FAILED {name}: {message}', flush=True)
                answers.append([name, f'failed: {" ".join(str(message).split())}', files_of(folder)])
                excel.new_document()
                continue
            answers.append([name, masked(str(result.value), folder), files_of(folder)])
        tidy(Path(root))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'seeds': {name: [one for one in body if one.startswith('@seed')] for name, body in CASES.items()},
              'cases': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer, files in answers:
        print(f'{name}: {answer}  ## {files}')


if __name__ == '__main__':
    main()
