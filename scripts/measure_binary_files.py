"""VBA's Binary and Random files in live Excel: how Put lays out each type, what Get reads back, Lock and Unlock, and what
the other file statements do in these modes.

Built like scripts/measure_file_statements.py, and run the same way: each case has its own empty folder, the answer is
what it says and every file it leaves, and the harness's line numbers stay off.

    python scripts/measure_binary_files.py

writes tests/fixtures/binary_files.json, which tests/test_vba_binary_files.py replays.
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

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/binary_files.json'
OPEN_B = 'Open P("a.bin") For Binary As #1'
OPEN_R = 'Open P("a.bin") For Random As #1'
AFTER = ['Say Seek(1)', 'Say Loc(1)', 'Say LOF(1)', 'Say EOF(1)', 'Close #1']


def put(declared: str, *setup: str, opening: str = OPEN_B) -> list[str]:
    return [f'Dim x As {declared}', *setup, opening, 'Put #1, , x', *AFTER]


def get(data: str, declared: str, *setup: str, opening: str = OPEN_B) -> list[str]:
    return [f'@seed a.bin {data}', f'Dim x As {declared}', *setup, opening, 'Get #1, , x', 'Say Show(x)', *AFTER]


SCALARS = {
    'byte': ('Byte', 'x = 200', 'C8'), 'integer': ('Integer', 'x = 258', '0201'), 'long': ('Long', 'x = 16909060', '04030201'),
    'single': ('Single', 'x = 1.5', '0000C03F'), 'double': ('Double', 'x = 1.5', '000000000000F83F'),
    'currency': ('Currency', 'x = 1.5', 'F03C000000000000'),
    'date': ('Date', 'x = #1/2/2020#', '0000000040E8E540'), 'boolean': ('Boolean', 'x = True', 'FFFF'),
    'longlong': ('LongLong', 'x = 5', '0500000000000000'),
}
VARIANTS = {'integer': 'CInt(258)', 'long': 'CLng(258)', 'single': 'CSng(1.5)', 'double': 'CDbl(1.5)', 'currency': 'CCur(1.5)',
            'date': '#1/2/2020#', 'boolean': 'True', 'string': '"abc"', 'byte': 'CByte(7)', 'empty': 'Empty', 'null': 'Null',
            'decimal': 'CDec(1.5)', 'error': 'CVErr(5)', 'empty string': '""', 'nothing': 'Nothing'}

CASES: dict[str, list[str]] = {
    **{f'put {name}': put(declared, setup) for name, (declared, setup, _) in SCALARS.items()},
    **{f'get {name}': get(data, declared) for name, (declared, _, data) in SCALARS.items()},
    **{f'put variant {name}': ['Dim x As Variant', f'x = {value}', OPEN_B, 'Put #1, , x', *AFTER] for name, value in VARIANTS.items()},
    **{f'put random variant {name}': ['Dim x As Variant', f'x = {value}', OPEN_R, 'Put #1, , x', *AFTER]
       for name, value in VARIANTS.items()},
    **{f'put literal {name}': [OPEN_B, f'Put #1, , {value}', *AFTER] for name, value in {
        'small': '5', 'large': '70000', 'float': '1.5', 'text': '"abc"', 'true': 'True', 'expression': '1 + 1', 'concat': '"a" & "b"'}.items()},
    'put string': put('String', 'x = "abc"'),
    'put string accented': put('String', 'x = ChrW(233) & ChrW(1234)'),
    'put string empty': put('String'),
    'put fixed': put('String * 5', 'x = "abc"'),
    'put fixed long': put('String * 2', 'x = "abc"'),
    'get string': get('6162630000', 'String', 'x = Space(3)'),
    'get string empty': get('616263', 'String'),
    'get string short': get('6162', 'String', 'x = Space(5)'),
    'get fixed': get('6162636465', 'String * 3'),
    'get fixed short': get('61', 'String * 3'),
    'get long short': get('0102', 'Long'),
    'get none left': get('', 'Long', 'x = 9'),
    'get double short': get('01', 'Double', 'x = 9'),
    'get string none': get('', 'String', 'x = "abc"'),
    'get variant integer': get('02000201', 'Variant'),
    'get variant string': get('0800030061626364', 'Variant'),
    'get variant empty': get('0000', 'Variant'),
    'get variant null': get('0100', 'Variant'),
    'get variant bad': get('6300010203', 'Variant'),
    'get variant none': get('', 'Variant'),
    # arrays
    'put fixed array': ['Dim x(2) As Long', 'x(0) = 1: x(1) = 2: x(2) = 3', OPEN_B, 'Put #1, , x', *AFTER],
    'put dynamic array': ['Dim x() As Long', 'ReDim x(2)', 'x(0) = 1: x(1) = 2: x(2) = 3', OPEN_B, 'Put #1, , x', *AFTER],
    'put byte array': ['Dim x() As Byte', 'ReDim x(2)', 'x(0) = 1: x(1) = 2: x(2) = 3', OPEN_B, 'Put #1, , x', *AFTER],
    'put two dimensions': ['Dim x(1, 1) As Integer', 'x(0, 0) = 1: x(1, 0) = 2: x(0, 1) = 3: x(1, 1) = 4', OPEN_B, 'Put #1, , x', *AFTER],
    'put string array': ['Dim x(1) As String', 'x(0) = "a": x(1) = "bc"', OPEN_B, 'Put #1, , x', *AFTER],
    'put variant array': ['Dim x As Variant', 'x = Array(1, "b")', OPEN_B, 'Put #1, , x', *AFTER],
    'put variant byte array': ['Dim b() As Byte', 'Dim x As Variant', 'ReDim b(1)', 'b(0) = 7', 'x = b', OPEN_B, 'Put #1, , x', *AFTER],
    'put unsized array': ['Dim x() As Long', OPEN_B, 'Put #1, , x', *AFTER],
    'put array element': ['Dim x(2) As Long', 'x(1) = 7', OPEN_B, 'Put #1, , x(1)', *AFTER],
    'get fixed array': ['@seed a.bin 010000000200000003000000', 'Dim x(1) As Long', OPEN_B, 'Get #1, , x', 'Say Show(x(0))', 'Say Show(x(1))', *AFTER],
    'get dynamic array': ['@seed a.bin 0102030405', 'Dim x() As Byte', 'ReDim x(2)', OPEN_B, 'Get #1, , x', 'Say Show(x(0))', 'Say Show(x(2))', *AFTER],
    'get unsized array': ['@seed a.bin 0102030405', 'Dim x() As Byte', OPEN_B, 'Get #1, , x', 'Say Show(UBound(x))', *AFTER],
    'get string array': ['@seed a.bin 010061620200636465', 'Dim x(1) As String', OPEN_B, 'Get #1, , x', 'Say Show(x(0))', 'Say Show(x(1))', *AFTER],
    'get array element': ['@seed a.bin 0102', 'Dim x(1) As Byte', OPEN_B, 'Get #1, , x(1)', 'Say Show(x(1))', *AFTER],
    # user-defined types
    'put record': ['Dim x As Rec', 'x.a = 1: x.b = "xy": x.c = 2: x.d = 3', OPEN_B, 'Put #1, , x', 'Say Len(x)', *AFTER],
    'put record text': ['Dim x As TextRec', 'x.a = 1: x.s = "abc": x.c = 2', OPEN_B, 'Put #1, , x', *AFTER],
    'put record variant': ['Dim x As VarRec', 'x.a = 1: x.v = "abc": x.c = 2', OPEN_B, 'Put #1, , x', *AFTER],
    'put record array': ['Dim x As ArrRec', 'x.n(0) = 1: x.n(1) = 2: x.c = 3', OPEN_B, 'Put #1, , x', *AFTER],
    'put record dynamic': ['Dim x As DynRec', 'ReDim x.n(1)', 'x.n(0) = 1: x.n(1) = 2: x.c = 3', OPEN_B, 'Put #1, , x', *AFTER],
    'put nested': ['Dim x As Outer', 'x.r.a = 1: x.r.b = "xy": x.z = 5', OPEN_B, 'Put #1, , x', *AFTER],
    'put records': ['Dim x(1) As Rec', 'x(0).a = 1: x(1).a = 2', OPEN_B, 'Put #1, , x', *AFTER],
    'get record': ['@seed a.bin 010000007879020003', 'Dim x As Rec', OPEN_B, 'Get #1, , x', 'Say Show(x.a)', 'Say Show(x.b)', 'Say Show(x.c)', 'Say Show(x.d)', *AFTER],
    'get record text': ['@seed a.bin 01000000030061626302', 'Dim x As TextRec', OPEN_B, 'Get #1, , x', 'Say Show(x.a)', 'Say Show(x.s)', 'Say Show(x.c)', *AFTER],
    # where it goes
    'put at a place': ['Dim x As Integer', 'x = 1', OPEN_B, 'Put #1, 5, x', *AFTER],
    'put then back': ['Dim x As Integer', 'x = 1', OPEN_B, 'Put #1, , x', 'x = 2', 'Put #1, 1, x', 'Say Seek(1)', 'x = 3', 'Put #1, , x', *AFTER],
    'put at zero': ['Dim x As Integer', OPEN_B, 'Put #1, 0, x', 'Put #1, -1, x', *AFTER],
    'get at a place': ['@seed a.bin 0102030405', 'Dim x As Byte', OPEN_B, 'Get #1, 3, x', 'Say Show(x)', 'Say Seek(1)', 'Get #1, , x', 'Say Show(x)', *AFTER],
    'get at zero': ['@seed a.bin 01', 'Dim x As Byte', OPEN_B, 'Get #1, 0, x', 'Get #1, -1, x', *AFTER],
    'get past': ['@seed a.bin 01', 'Dim x As Byte', 'x = 9', OPEN_B, 'Get #1, 50, x', 'Say Show(x)', *AFTER],
    'seek past': ['@seed a.bin 01', OPEN_B, 'Seek #1, 50', *AFTER],
    'measures': ['@seed a.bin 0102030405', 'Dim x As Integer', OPEN_B, 'Say Seek(1)', 'Say Loc(1)', 'Say EOF(1)', 'Get #1, , x', 'Say Seek(1)',
                 'Say Loc(1)', 'Say EOF(1)', 'Get #1, , x', 'Say EOF(1)', 'Get #1, , x', 'Say EOF(1)', 'Say Seek(1)', 'Say Loc(1)', 'Get #1, , x',
                 'Say EOF(1)', 'Say Seek(1)', *AFTER],
    'ctrl z': ['@seed a.bin 011A02', 'Dim x As Byte', OPEN_B, 'Get #1, , x', 'Say EOF(1)', 'Get #1, , x', 'Say EOF(1)', 'Say Show(x)', *AFTER],
    # the mode
    'creates': [OPEN_B, 'Close #1'],
    'open twice': [OPEN_B, 'Open P("a.bin") For Binary As #2', 'Close'],
    'open with input': ['@seed a.bin 61', OPEN_B, 'Open P("a.bin") For Input As #2', 'Close'],
    'read only': ['@seed a.bin 61', 'SetAttr P("a.bin"), vbReadOnly', OPEN_B, 'Say LOF(1)', 'Dim x As Byte', 'Get #1, , x', 'Say Show(x)', 'x = 1',
                  'Put #1, , x', 'Close #1', 'SetAttr P("a.bin"), vbNormal'],
    'access read': ['@seed a.bin 61', 'Dim x As Byte', 'Open P("a.bin") For Binary Access Read As #1', 'Get #1, , x', 'Say Show(x)', 'Put #1, , x',
                    'Say FileAttr(1, 1)', 'Close #1'],
    'access write': ['@seed a.bin 61', 'Dim x As Byte', 'Open P("a.bin") For Binary Access Write As #1', 'Put #1, , x', 'Get #1, , x', 'Close #1'],
    'access read missing': ['Open P("a.bin") For Binary Access Read As #1', 'Close #1'],
    'access read write': ['Dim x As Byte', 'Open P("a.bin") For Binary Access Read Write As #1', 'Put #1, , x', 'Close #1'],
    'file attr': [OPEN_B, 'Say FileAttr(1, 1)', 'Say FileAttr(1)', 'Close #1', OPEN_R, 'Say FileAttr(1, 1)', 'Close #1'],
    'text in binary': ['@seed a.bin 6162630D0A642C65', OPEN_B, 'Say Input(2, #1)', 'Line Input #1, t', 'Say t', 'Dim a, b', 'Input #1, a, b', 'Say Show(a)', 'Say Show(b)',
                       'Say EOF(1)', 'Close #1'],
    'print in binary': ['@seed a.bin 616263', OPEN_B, 'Print #1, "x"', 'Write #1, "y"', 'Say Seek(1)', 'Width #1, 3', 'Close #1'],
    'get in input': ['@seed a.bin 61', 'Dim x As Byte', 'Open P("a.bin") For Input As #1', 'Get #1, , x', 'Put #1, , x', 'Close #1'],
    'get in output': ['Dim x As Byte', 'Open P("a.bin") For Output As #1', 'Get #1, , x', 'Put #1, , x', 'Close #1'],
    'get closed': ['Dim x As Byte', 'Get #1, , x', 'Put #1, , x'],
    # Random
    'random default': ['Dim x As Long', 'x = 1', OPEN_R, 'Put #1, , x', 'x = 2', 'Put #1, , x', 'Say Seek(1)', 'Say Loc(1)', 'Say LOF(1)', 'Close #1'],
    'random no for': ['Dim x As Long', 'x = 1', 'Open P("a.bin") As #1', 'Put #1, , x', 'Say FileAttr(1, 1)', 'Close #1'],
    'random len': ['Dim x As Long', 'x = 1', 'Open P("a.bin") For Random As #1 Len = 6', 'Put #1, , x', 'x = 2', 'Put #1, 3, x', 'Say Seek(1)', 'Say Loc(1)',
                   'Say LOF(1)', 'Say EOF(1)', 'Close #1'],
    'random len tight': ['Dim x As Long', 'Open P("a.bin") For Random As #1 Len = 4', 'Put #1, , x', 'Close #1'],
    'random too long': ['Dim x As Long', 'Open P("a.bin") For Random As #1 Len = 3', 'Put #1, , x', 'Say LOF(1)', 'Close #1'],
    'random string': ['Dim x As String', 'x = "abc"', 'Open P("a.bin") For Random As #1 Len = 8', 'Put #1, , x', 'Say LOF(1)', 'Close #1'],
    'random string full': ['Dim x As String', 'x = "abc"', 'Open P("a.bin") For Random As #1 Len = 5', 'Put #1, , x', 'Close #1'],
    'random string over': ['Dim x As String', 'x = "abc"', 'Open P("a.bin") For Random As #1 Len = 4', 'Put #1, , x', 'Say LOF(1)', 'Close #1'],
    'random fixed': ['Dim x As String * 3', 'x = "abc"', 'Open P("a.bin") For Random As #1 Len = 8', 'Put #1, , x', 'Close #1'],
    'random record': ['Dim x As Rec', 'x.a = 1: x.b = "xy": x.c = 2: x.d = 3', 'Open P("a.bin") For Random As #1 Len = Len(x)', 'Put #1, , x', 'Say Len(x)', 'Say LOF(1)',
                      'Put #1, 2, x', 'Say LOF(1)', 'Close #1'],
    'random record long': ['Dim x As Rec', 'x.a = 1', 'Open P("a.bin") For Random As #1 Len = 20', 'Put #1, , x', 'Close #1'],
    'random record short': ['Dim x As Rec', 'Open P("a.bin") For Random As #1 Len = 4', 'Put #1, , x', 'Close #1'],
    'random record text': ['Dim x As TextRec', 'x.a = 1: x.s = "abc": x.c = 2', 'Open P("a.bin") For Random As #1 Len = 20', 'Put #1, , x', 'Close #1'],
    'random array': ['Dim x(1) As Integer', 'x(0) = 1: x(1) = 2', 'Open P("a.bin") For Random As #1 Len = 8', 'Put #1, , x', 'Close #1'],
    'random dynamic': ['Dim x() As Integer', 'ReDim x(1)', 'x(0) = 1: x(1) = 2', 'Open P("a.bin") For Random As #1 Len = 12', 'Put #1, , x', 'Close #1'],
    'random dynamic two': ['Dim x() As Integer', 'ReDim x(1, 1)', 'Open P("a.bin") For Random As #1 Len = 20', 'Put #1, , x', 'Close #1'],
    'random variant array': ['Dim x As Variant', 'x = Array(1)', 'Open P("a.bin") For Random As #1 Len = 30', 'Put #1, , x', 'Close #1'],
    'random get': ['@seed a.bin 0100000000000200000000000300', 'Dim x As Long', 'Open P("a.bin") For Random As #1 Len = 6', 'Get #1, 2, x', 'Say Show(x)', 'Say Seek(1)',
                   'Say Loc(1)', 'Get #1, , x', 'Say Show(x)', 'Say Seek(1)', 'Say EOF(1)', 'Get #1, , x', 'Say Show(x)', 'Say EOF(1)', 'Close #1'],
    'random get string': ['@seed a.bin 030061626300000000', 'Dim x As String', 'Open P("a.bin") For Random As #1 Len = 8', 'Get #1, , x', 'Say Show(x)', 'Say Seek(1)', 'Close #1'],
    'random get string long': ['@seed a.bin 6300616263', 'Dim x As String', 'Open P("a.bin") For Random As #1 Len = 8', 'Get #1, , x', 'Say Show(x)', 'Close #1'],
    'random get record': ['@seed a.bin 010000007879020003', 'Dim x As Rec', 'Open P("a.bin") For Random As #1 Len = 9', 'Get #1, , x', 'Say Show(x.a)', 'Say Show(x.b)',
                          'Say Show(x.c)', 'Say Show(x.d)', 'Close #1'],
    'random get too long': ['@seed a.bin 0100000000', 'Dim x As Long', 'Open P("a.bin") For Random As #1 Len = 2', 'Get #1, , x', 'Say Show(x)', 'Close #1'],
    'random get variant': ['@seed a.bin 0200010000', 'Dim x As Variant', 'Open P("a.bin") For Random As #1 Len = 8', 'Get #1, , x', 'Say Show(x)', 'Close #1'],
    'random zero': ['Dim x As Long', OPEN_R, 'Put #1, 0, x', 'Get #1, 0, x', 'Seek #1, 0', 'Close #1'],
    'random seek': ['@seed a.bin 010002000300', 'Dim x As Integer', 'Open P("a.bin") For Random As #1 Len = 2', 'Seek #1, 3', 'Say Seek(1)', 'Get #1, , x', 'Say Show(x)', 'Say Seek(1)',
                    'Say Loc(1)', 'Close #1'],
    'random len bad': ['Open P("a.bin") For Random As #1 Len = 0', 'Close', 'Open P("a.bin") For Random As #1 Len = -1', 'Close', 'Open P("b.bin") For Random As #1 Len = 32767', 'Close',
                       'Open P("c.bin") For Random As #1 Len = 32768', 'Close'],
    'random len zero': ['Dim x As Long', 'Open P("a.bin") For Random As #1 Len = 0', 'Say FileAttr(1, 1)', 'Put #1, , x', 'Say LOF(1)', 'Say Seek(1)', 'Close'],
    'random len minus': ['Dim x As Long', 'Open P("a.bin") For Random As #1 Len = -1', 'Say FileAttr(1, 1)', 'Put #1, , x', 'Say LOF(1)', 'Close'],
    'random len one': ['Dim x As Byte', 'Open P("a.bin") For Random As #1 Len = 1', 'Put #1, , x', 'Put #1, , x', 'Say Seek(1)', 'Close'],
    'binary len': ['Dim x As Long', 'Open P("a.bin") For Binary As #1 Len = 2', 'Put #1, , x', 'Say Seek(1)', 'Close'],
    'random text': ['@seed a.bin 6162630D0A', OPEN_R, 'Print #1, "x"', 'Write #1, "y"', 'Line Input #1, t', 'Input #1, t', 'Say Input(1, #1)', 'Close #1'],
    'random twice': [OPEN_R, 'Open P("a.bin") For Random As #2', 'Close'],
    'random measures': ['Dim x As Integer', 'Open P("a.bin") For Random As #1 Len = 2', 'Say Seek(1)', 'Say Loc(1)', 'Say EOF(1)', 'Say LOF(1)', 'Put #1, , x', 'Say Seek(1)', 'Say Loc(1)', 'Say EOF(1)',
                        'Say LOF(1)', 'Close #1'],
    # Lock and Unlock
    'lock all': ['@seed a.bin 010203', OPEN_B, 'Lock #1', 'Unlock #1', 'Close #1'],
    'lock range': ['@seed a.bin 010203', OPEN_B, 'Lock #1, 1 To 2', 'Unlock #1, 1 To 2', 'Lock #1, 3', 'Unlock #1, 3', 'Lock #1, To 2', 'Unlock #1, To 2', 'Close #1'],
    'lock twice': ['@seed a.bin 010203', OPEN_B, 'Lock #1, 1 To 2', 'Lock #1, 1 To 2', 'Close #1'],
    'unlock unlocked': ['@seed a.bin 010203', OPEN_B, 'Unlock #1, 1 To 2', 'Unlock #1', 'Close #1'],
    'unlock different': ['@seed a.bin 010203', OPEN_B, 'Lock #1, 1 To 2', 'Unlock #1, 1 To 3', 'Close #1'],
    'lock backwards': ['@seed a.bin 010203', OPEN_B, 'Lock #1, 3 To 1', 'Close #1'],
    'lock zero': ['@seed a.bin 010203', OPEN_B, 'Lock #1, 0', 'Lock #1, 0 To 1', 'Close #1'],
    'lock closed': ['Lock #1', 'Unlock #1'],
    'lock output': ['Open P("a.bin") For Output As #1', 'Lock #1', 'Unlock #1', 'Close #1'],
    'lock input': ['@seed a.bin 61', 'Open P("a.bin") For Input As #1', 'Lock #1', 'Unlock #1', 'Close #1'],
    'lock random': [OPEN_R, 'Lock #1, 1', 'Unlock #1, 1', 'Lock #1, 1 To 3', 'Unlock #1, 1 To 3', 'Close #1'],
    'locked read': ['@seed a.bin 010203', 'Dim x As Byte', OPEN_B, 'Lock #1, 1 To 2', 'Get #1, 1, x', 'Say Show(x)', 'Open P("a.bin") For Input As #2', 'Close'],
    # Lock and the rest
    'kill open binary': ['@seed a.bin 61', OPEN_B, 'Kill P("a.bin")', 'Close #1'],
    'copy open binary': ['@seed a.bin 61', OPEN_B, 'FileCopy P("a.bin"), P("b.bin")', 'Close #1'],
    'file len open binary': ['Dim x As Long', OPEN_B, 'Put #1, , x', 'Say FileLen(P("a.bin"))', 'Close #1', 'Say FileLen(P("a.bin"))'],
}

_TYPES = '''
Private Type Rec
    a As Long
    b As String * 2
    c As Integer
    d As Byte
End Type

Private Type TextRec
    a As Long
    s As String
    c As Byte
End Type

Private Type VarRec
    a As Long
    v As Variant
    c As Byte
End Type

Private Type ArrRec
    n(1) As Integer
    c As Byte
End Type

Private Type DynRec
    n() As Integer
    c As Byte
End Type

Private Type Outer
    r As Rec
    z As Byte
End Type
'''


def _probe() -> str:
    lines = [_TYPES, _HELPERS]
    names = list(CASES)
    for index, name in enumerate(names):
        lines += [f'Private Function Case{index}() As String', '    Dim s As String, v As Variant, t As Variant',
                  '    On Error Resume Next']
        for text in CASES[name]:
            lines += _statement(text)
        lines += ['    Close', '    Err.Clear', f'    Case{index} = s', 'End Function', '']
    lines += ['Public Function Probe(ByVal folder As String, ByVal sep As String, ByVal which As Long) As String',
              '    gFolder = folder: gSep = sep', '    Select Case which']
    lines += [f'        Case {index}: Probe = Case{index}()' for index in range(len(names))]
    lines += ['    End Select', 'End Function']
    return '\n'.join(lines) + '\n'


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
            answers.append([name, str(result.value).replace(str(folder), 'FOLDER'), files_of(folder)])
        tidy(Path(root))
    record = {'measured_at': datetime.now(timezone.utc).isoformat(), 'source': source,
              'seeds': {name: [text for text in body if text.startswith('@seed')] for name, body in CASES.items()},
              'cases': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer, files in answers:
        print(f'{name}: {answer}  ## {files}')


if __name__ == '__main__':
    main()
