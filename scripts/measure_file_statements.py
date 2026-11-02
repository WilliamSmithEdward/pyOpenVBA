"""VBA's text files and its file system statements in live Excel: Open, Print #, Write #, Input #, Line Input #, Close, the
functions that go with them, and Kill, MkDir, RmDir, FileCopy, Name, Dir, FileLen, GetAttr, SetAttr, ChDir and CurDir.

Each case runs in its own empty folder under a temporary one. It runs its statements, each followed by any error's
number and description, and the answer is that with what the statements recorded, then every file the folder ends up
with, in hex. Nothing outside the temporary folder is touched, and ChDir is put back. Run without the harness's line
numbers, whose injected On Error would clear Err on entry to the helpers.

    python scripts/measure_file_statements.py

writes tests/fixtures/file_statements.json, which tests/test_vba_file_statements.py replays.
"""
from __future__ import annotations

import json
import os
import stat
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from pyvbaharness import ExcelSession, HarnessConfig

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/file_statements.json'


def out(*items: str) -> list[str]:
    """Open a file for output, print each item to it, and close it."""
    return ['Open P("a.txt") For Output As #1', *[f'Print #1, {item}' for item in items], 'Close #1']


def writes(*items: str) -> list[str]:
    return ['Open P("a.txt") For Output As #1', *[f'Write #1, {item}' for item in items], 'Close #1']


def read_lines() -> list[str]:
    return ['Open P("a.txt") For Input As #1', 'Do While Not EOF(1)', 'Line Input #1, t', 'Say t', 'Loop', 'Close #1']


def seeded(data: str, *body: str) -> list[str]:
    """A file of these bytes (hex), written by the case before it starts, and the body."""
    return [f'@seed a.txt {data}', *body]


CASES: dict[str, list[str]] = {
    # Print #
    'print text': out('"abc"'),
    'print nothing': ['Open P("a.txt") For Output As #1', 'Print #1,', 'Close #1'],
    'print none at all': ['Open P("a.txt") For Output As #1', 'Close #1'],
    'print numbers': out('1', '-1', '1.5', '1E+20', 'CInt(3)', 'CLng(3)', 'CCur(1.5)', 'CDbl(0.1)', 'CSng(0.1)', 'CDec(2)'),
    'print others': out('True', 'False', '#1/2/2020#', '#13:14:15#', 'Null', 'Empty', 'CVErr(5)'),
    'print semicolon': ['Open P("a.txt") For Output As #1', 'Print #1, "a"; "b"; 1; 2;', 'Print #1, "c"', 'Close #1'],
    'print comma': ['Open P("a.txt") For Output As #1', 'Print #1, "a", "b", 1,', 'Print #1, "c"', 'Close #1'],
    'print long comma': out('String(13, "a"), "b"', 'String(14, "a"), "b"', 'String(15, "a"), "b"'),
    'print tab': out('"a"; Tab(5); "b"', 'Tab(3); "c"', '"abcdef"; Tab(3); "d"', '"a"; Tab; "b"', 'Tab(0); "e"'),
    'print spc': out('"a"; Spc(3); "b"', 'Spc(2); "c"', '"a"; Spc(0); "b"'),
    'print accented': out('ChrW(233)', 'ChrW(8364)', 'ChrW(1234)', 'ChrW(55357)'),
    'print breaks': out('"a" & vbCrLf & "b"', '"c" & vbLf & "d"', '"e" & Chr(0) & "f"'),
    'print array': ['Dim arr(1) As Long', 'v = arr', 'Open P("a.txt") For Output As #1', 'Print #1, v', 'Close #1'],
    'print object': ['Open P("a.txt") For Output As #1', 'Print #1, Nothing', 'Close #1'],
    'width': ['Open P("a.txt") For Output As #1', 'Width #1, 3', 'Print #1, "abcdefgh"', 'Print #1, "a"; "b"; "c"; "d"',
              'Close #1'],
    'width zero': ['Open P("a.txt") For Output As #1', 'Width #1, 0', 'Print #1, "abcdefgh"', 'Close #1'],
    'width too big': ['Open P("a.txt") For Output As #1', 'Width #1, 256', 'Print #1, "abcdefgh"', 'Close #1'],
    # Write #
    'write text': writes('"abc"', '"a""b"', '""', '"a,b"'),
    'write numbers': writes('1', '-1', '1.5', '1E+20', 'CCur(1.5)', 'CDbl(0.1)', 'CSng(0.1)'),
    'write others': writes('True', 'False', '#1/2/2020#', '#13:14:15#', '#1/2/2020 13:14:15#', 'Null', 'Empty', 'CVErr(5)'),
    'write many': writes('"a", 1, True', '1; 2; 3'),
    'write trailing': ['Open P("a.txt") For Output As #1', 'Write #1, "a";', 'Write #1, "b",', 'Write #1,', 'Close #1'],
    'write accented': writes('ChrW(233)', 'ChrW(1234)'),
    # Open and Close
    'append': ['Open P("a.txt") For Output As #1', 'Print #1, "a"', 'Close #1', 'Open P("a.txt") For Append As #1',
               'Print #1, "b"', 'Close #1'],
    'output empties': seeded('616263', 'Open P("a.txt") For Output As #1', 'Close #1'),
    'as without sharp': ['Open P("a.txt") For Output As 1', 'Print #1, "a"', 'Close 1'],
    'as a variable': ['n = FreeFile', 'Open P("a.txt") For Output As #n', 'Print #n, "a"', 'Close #n', 'Say n'],
    'close several': ['Open P("a.txt") For Output As #1', 'Open P("b.txt") For Output As #2', 'Print #1, "a"',
                      'Print #2, "b"', 'Close #1, #2', 'Print #1, "c"'],
    'close all': ['Open P("a.txt") For Output As #1', 'Open P("b.txt") For Output As #2', 'Close', 'Print #1, "c"',
                  'Say FreeFile'],
    'reset': ['Open P("a.txt") For Output As #1', 'Reset', 'Print #1, "c"', 'Say FreeFile'],
    'close not open': ['Close #9', 'Close #9, #8', 'Say "fine"'],
    'open twice same number': ['Open P("a.txt") For Output As #1', 'Open P("b.txt") For Output As #1', 'Close #1'],
    'output twice': ['Open P("a.txt") For Output As #1', 'Open P("a.txt") For Output As #2', 'Close'],
    'input twice': seeded('61', 'Open P("a.txt") For Input As #1', 'Open P("a.txt") For Input As #2', 'Close'),
    'input then output': seeded('61', 'Open P("a.txt") For Input As #1', 'Open P("a.txt") For Output As #2', 'Close'),
    'output then input': ['Open P("a.txt") For Output As #1', 'Open P("a.txt") For Input As #2', 'Close'],
    'append twice': ['Open P("a.txt") For Append As #1', 'Open P("a.txt") For Append As #2', 'Close'],
    'number zero': ['Open P("a.txt") For Output As #0'],
    'number too big': ['Open P("a.txt") For Output As #512'],
    'number 511': ['Open P("a.txt") For Output As #511', 'Print #511, "a"', 'Close #511'],
    'number negative': ['Open P("a.txt") For Output As #-1'],
    'number text': ['Open P("a.txt") For Output As #"3"', 'Print #3, "a"', 'Close #3'],
    'number float': ['Open P("a.txt") For Output As #2.6', 'Print #3, "a"', 'Close #3'],
    'input missing': ['Open P("none.txt") For Input As #1', 'Say "opened"'],
    'append missing': ['Open P("none.txt") For Append As #1', 'Close #1'],
    'no folder': ['Open P("none\\a.txt") For Output As #1'],
    'no folder input': ['Open P("none\\a.txt") For Input As #1'],
    'empty name': ['Open "" For Output As #1'],
    'open a folder': ['MkDir P("d")', 'Open P("d") For Output As #1'],
    'open a folder input': ['MkDir P("d")', 'Open P("d") For Input As #1'],
    'read only output': seeded('61', 'SetAttr P("a.txt"), vbReadOnly', 'Open P("a.txt") For Output As #1',
                               'SetAttr P("a.txt"), vbNormal'),
    'read only input': seeded('61', 'SetAttr P("a.txt"), vbReadOnly', 'Open P("a.txt") For Input As #1', 'Close #1',
                              'SetAttr P("a.txt"), vbNormal'),
    'access clauses': ['Open P("a.txt") For Output Access Write As #1', 'Print #1, "a"', 'Close #1',
                       'Open P("a.txt") For Input Access Read As #1', 'Close #1'],
    'lock clauses': ['Open P("a.txt") For Output Lock Write As #1', 'Open P("a.txt") For Input As #2', 'Close'],
    'shared': ['Open P("a.txt") For Output Shared As #1', 'Close #1', 'Open P("a.txt") For Input Lock Read As #1', 'Close'],
    'len clause': ['Open P("a.txt") For Output As #1 Len = 2', 'Print #1, "abcdef"', 'Close #1'],
    # the wrong number or mode
    'print to closed': ['Print #1, "a"'],
    'write to closed': ['Write #1, "a"'],
    'line input closed': ['Line Input #1, t'],
    'input closed': ['Input #1, t'],
    'eof closed': ['Say EOF(1)'],
    'lof closed': ['Say LOF(1)'],
    'loc closed': ['Say Loc(1)'],
    'seek closed': ['Say Seek(1)'],
    'print to input': seeded('61', 'Open P("a.txt") For Input As #1', 'Print #1, "a"', 'Close #1'),
    'write to input': seeded('61', 'Open P("a.txt") For Input As #1', 'Write #1, "a"', 'Close #1'),
    'input from output': ['Open P("a.txt") For Output As #1', 'Line Input #1, t', 'Input #1, t', 'Say Input(1, #1)',
                          'Say EOF(1)', 'Close #1'],
    # Line Input # and friends
    'lines crlf': seeded('610D0A620D0A', *read_lines()),
    'lines lf': seeded('610A620A', *read_lines()),
    'lines cr': seeded('610D620D', *read_lines()),
    'lines lfcr': seeded('610A0D62', *read_lines()),
    'lines no end': seeded('610D0A62', *read_lines()),
    'lines blank': seeded('0D0A0D0A61', *read_lines()),
    'lines empty': seeded('', *read_lines()),
    'lines nul': seeded('6100620D0A', *read_lines()),
    'lines ctrl z': seeded('610D0A1A620D0A', *read_lines()),
    'lines ctrl z end': seeded('610D0A1A', *read_lines()),
    'lines latin': seeded('E98081FF0D0A', *read_lines()),
    'lines utf8 bom': seeded('EFBBBF61C3A90D0A', *read_lines()),
    'past the end': seeded('610D0A', 'Open P("a.txt") For Input As #1', 'Line Input #1, t', 'Say t', 'Line Input #1, t',
                           'Say t', 'Close #1'),
    'into a string': seeded('310D0A', 'Dim k As String', 'Open P("a.txt") For Input As #1', 'Line Input #1, k', 'Say k',
                            'Close #1'),
    'into a cell': seeded('610D0A', 'Open P("a.txt") For Input As #1', 'Line Input #1, gCell(1)', 'Say gCell(1)', 'Close #1'),
    'measures': seeded('6162630D0A64', 'Open P("a.txt") For Input As #1', 'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)',
                       'Say EOF(1)', 'Line Input #1, t', 'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)', 'Say EOF(1)',
                       'Line Input #1, t', 'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)', 'Say EOF(1)', 'Close #1'),
    'measures output': ['Open P("a.txt") For Output As #1', 'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)', 'Print #1, "abc"',
                        'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)', 'Close #1', 'Say FileLen(P("a.txt"))'],
    'seek back': seeded('6162630D0A64', 'Open P("a.txt") For Input As #1', 'Line Input #1, t', 'Seek #1, 2',
                        'Line Input #1, t', 'Say t', 'Seek #1, 100', 'Say EOF(1)', 'Say Seek(1)', 'Seek #1, 0', 'Say Seek(1)',
                        'Close #1'),
    'seek output': ['Open P("a.txt") For Output As #1', 'Print #1, "abcdef"', 'Seek #1, 2', 'Print #1, "x"', 'Close #1'],
    'input chars': seeded('6162630D0A64', 'Open P("a.txt") For Input As #1', 'Say Input(2, #1)', 'Say Input(3, #1)',
                          'Say Loc(1)', 'Say Input$(1, 1)', 'Say EOF(1)', 'Say Input(1, #1)', 'Close #1'),
    'input zero': seeded('61', 'Open P("a.txt") For Input As #1', 'Say Input(0, #1)', 'Say Input(-1, #1)', 'Close #1'),
    'file attr': seeded('61', 'Open P("a.txt") For Input As #1', 'Say FileAttr(1, 1)', 'Close #1',
                        'Open P("b.txt") For Output As #1', 'Say FileAttr(1)', 'Close #1',
                        'Open P("b.txt") For Append As #1', 'Say FileAttr(1, 1)', 'Say FileAttr(2, 1)', 'Close #1'),
    'free file': ['Say FreeFile', 'Say FreeFile()', 'Say FreeFile(0)', 'Say FreeFile(1)', 'Say FreeFile(2)',
                  'Open P("a.txt") For Output As #1', 'Say FreeFile', 'Close #1', 'Open P("a.txt") For Output As #256',
                  'Say FreeFile(1)', 'Close #256', 'Say TypeName(FreeFile)'],
    'loc large': seeded('61' * 1000, 'Open P("a.txt") For Input As #1', 'Say Loc(1)', 'Say Input(127, #1)', 'Say Loc(1)',
                        'Say Input(1, #1)', 'Say Loc(1)', 'Say Input(1, #1)', 'Say Loc(1)', 'Say Input(127, #1)', 'Say Loc(1)',
                        'Say Input(1, #1)', 'Say Loc(1)', 'Seek #1, 300', 'Say Loc(1)', 'Close #1'),
    'loc output large': ['Open P("a.txt") For Output As #1', 'Print #1, String(125, "a")', 'Say Loc(1)', 'Say Seek(1)',
                         'Print #1, "b"', 'Say Loc(1)', 'Print #1, String(125, "c")', 'Say Loc(1)', 'Say Seek(1)', 'Close #1'],
    'measures append': seeded('616263', 'Open P("a.txt") For Append As #1', 'Say LOF(1)', 'Say Loc(1)', 'Say Seek(1)',
                              'Print #1, "d"', 'Say LOF(1)', 'Say Seek(1)', 'Seek #1, 1', 'Print #1, "e"', 'Close #1'),
    'input past ctrl z': seeded('611A62', 'Open P("a.txt") For Input As #1', 'Say Input(1, #1)', 'Say EOF(1)', 'Say Input(1, #1)',
                                'Say EOF(1)', 'Say Input(1, #1)', 'Say EOF(1)', 'Close #1'),
    'number huge': ['Open P("a.txt") For Output As #600', 'Print #600, "a"', 'Close #600', 'Open P("b.txt") For Output As #32767',
                    'Close #32767', 'Open P("c.txt") For Output As #40000', 'Say "x"'],
    # Input #
    'fields': seeded('612C622C0D0A632C640D0A', 'Dim a, b, c, d', 'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d',
                     'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)', 'Close #1'),
    'fields quoted': seeded('2261206222202C2022632C64222C22652222660D0A', 'Dim a, b, c, d', 'Open P("a.txt") For Input As #1',
                            'Input #1, a, b, c', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Close #1'),
    'fields numbers': seeded('312C20312E352C2D322C31452B322C26483130', 'Dim a, b, c, d, e', 'Open P("a.txt") For Input As #1',
                             'Input #1, a, b, c, d, e', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)',
                             'Say Show(e)', 'Close #1'),
    'fields written': ['Open P("a.txt") For Output As #1', 'Write #1, "a", 1, 1.5, True, #1/2/2020#, Null, Empty, CVErr(5)',
                       'Close #1', 'Dim a, b, c, d, e, f, g, h', 'Open P("a.txt") For Input As #1',
                       'Input #1, a, b, c, d, e, f, g, h', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)',
                       'Say Show(e)', 'Say Show(f)', 'Say Show(g)', 'Say Show(h)', 'Close #1'],
    'fields typed': seeded('312C322E352C612C2C0D0A', 'Dim a As Long, b As Double, c As String, d As String, e As Long',
                           'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d, e', 'Say Show(a)', 'Say Show(b)',
                           'Say Show(c)', 'Say Show(d)', 'Say Show(e)', 'Close #1'),
    'fields typed bad': seeded('612C31', 'Dim a As Long, b As Long', 'Open P("a.txt") For Input As #1', 'Input #1, a, b',
                               'Say Show(a)', 'Say Show(b)', 'Close #1'),
    'fields overflow': seeded('393939393939', 'Dim a As Integer', 'Open P("a.txt") For Input As #1', 'Input #1, a',
                              'Say Show(a)', 'Close #1'),
    'fields past end': seeded('61', 'Dim a, b', 'Open P("a.txt") For Input As #1', 'Input #1, a, b', 'Say Show(a)',
                              'Say Show(b)', 'Input #1, a', 'Close #1'),
    'fields blank lines': seeded('0D0A0D0A612C0D0A0D0A62', 'Dim a, b, c', 'Open P("a.txt") For Input As #1',
                                 'Input #1, a, b, c', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say EOF(1)', 'Close #1'),
    'fields spaces': seeded('2020612062202020202C2020632020', 'Dim a, b', 'Open P("a.txt") For Input As #1', 'Input #1, a, b',
                            'Say Show(a)', 'Say Show(b)', 'Close #1'),
    'fields after quote': seeded('2261227863' + '2C64', 'Dim a, b', 'Open P("a.txt") For Input As #1', 'Input #1, a, b',
                                 'Say Show(a)', 'Say Show(b)', 'Close #1'),
    'fields tab': seeded('6109622C63', 'Dim a, b', 'Open P("a.txt") For Input As #1', 'Input #1, a, b', 'Say Show(a)',
                         'Say Show(b)', 'Close #1'),
    'fields semicolon': seeded('613B622C63', 'Dim a, b', 'Open P("a.txt") For Input As #1', 'Input #1, a, b', 'Say Show(a)',
                               'Say Show(b)', 'Close #1'),
    'fields date text': seeded('312F322F323032302C23323032302D30312D3032230D0A', 'Dim a, b',
                               'Open P("a.txt") For Input As #1', 'Input #1, a, b', 'Say Show(a)', 'Say Show(b)', 'Close #1'),
    'fields date typed': seeded('312F322F32303230', 'Dim a As Date', 'Open P("a.txt") For Input As #1', 'Input #1, a',
                                'Say Show(a)', 'Close #1'),
    'fields boolean typed': seeded('747275652C2354525545232C30', 'Dim a As Boolean, b As Boolean, c As Boolean',
                                   'Open P("a.txt") For Input As #1', 'Input #1, a, b, c', 'Say Show(a)', 'Say Show(b)',
                                   'Say Show(c)', 'Close #1'),
    'fields widths': seeded('37303030302C333030303030303030302C312E32333435362C312E32332C312E32333435352C2E352C31302E30',
                            'Dim a, b, c, d, e, f, g', 'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d, e, f, g', 'Say Show(a)',
                            'Say Show(b)', 'Say Show(c)', 'Say Show(d)', 'Say Show(e)', 'Say Show(f)', 'Say Show(g)', 'Close #1'),
    'fields prefixes': seeded('316162632C2D2C2B312C31204132202C2E2C312E352E36', 'Dim a, b, c, d, e, f', 'Open P("a.txt") For Input As #1',
                              'Input #1, a, b, c, d, e, f', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)', 'Say Show(e)',
                              'Say Show(f)', 'Close #1'),
    'fields words': seeded('547275652C46414C53452C4E756C6C2C454D5054592C2354525545', 'Dim a, b, c, d, e', 'Open P("a.txt") For Input As #1',
                           'Input #1, a, b, c, d, e', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)', 'Say Show(e)', 'Close #1'),
    'fields boolean number': seeded('312C352C302C2D31', 'Dim a As Boolean, b As Boolean, c As Boolean, d As Boolean',
                                    'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)',
                                    'Say Show(d)', 'Close #1'),
    'fields date number': seeded('352C23313A3030232C232D33232C31', 'Dim a As Date, b As Date, c As Date, d As String', 'Open P("a.txt") For Input As #1',
                                 'Input #1, a, b, c, d', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)', 'Close #1'),
    'fields typed marks': seeded('234E554C4C232C2354525545232C23323032302D30312D3032232C234552524F522035232C2C2C', 'Dim a As String, b As String, c As String, d As String, e As Double, f',
                                 'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d, e, f', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)',
                                 'Say Show(e)', 'Say Show(f)', 'Close #1'),
    'fields typed numbers': seeded('312E352C312E352C312E352C313030303030302C3635353335', 'Dim a As Integer, b As Single, c As Currency, d As Long, e As Byte',
                                   'Open P("a.txt") For Input As #1', 'Input #1, a, b, c, d, e', 'Say Show(a)', 'Say Show(b)', 'Say Show(c)', 'Say Show(d)',
                                   'Say Show(e)', 'Close #1'),
    'fields into a cell': seeded('312C32', 'Open P("a.txt") For Input As #1', 'Input #1, gCell(0), gCell(1)', 'Say Show(gCell(0))', 'Say Show(gCell(1))', 'Close #1'),
    # the file system
    'kill': seeded('61', 'Kill P("a.txt")', 'Say Dir(P("a.txt")) = ""'),
    'kill missing': ['Kill P("none.txt")'],
    'kill pattern': ['@seed a.txt 61', '@seed b.txt 61', '@seed c.log 61', 'Kill P("*.txt")', 'Say Dir(P("*"))'],
    'kill question': ['@seed a1.txt 61', '@seed a2.txt 61', '@seed a10.txt 61', 'Kill P("a?.txt")'],
    'kill no match': ['@seed a.txt 61', 'Kill P("*.log")'],
    'kill open': seeded('61', 'Open P("a.txt") For Input As #1', 'Kill P("a.txt")', 'Close #1'),
    'kill read only': seeded('61', 'SetAttr P("a.txt"), vbReadOnly', 'Kill P("a.txt")', 'SetAttr P("a.txt"), vbNormal'),
    'kill folder': ['MkDir P("d")', 'Kill P("d")', 'Say Dir(P("d"), vbDirectory)'],
    'kill empty': ['Kill ""'],
    'mkdir': ['MkDir P("d")', 'Say GetAttr(P("d"))', 'Say Dir(P("d"), vbDirectory)'],
    'mkdir exists': ['MkDir P("d")', 'MkDir P("d")'],
    'mkdir nested': ['MkDir P("d\\e")'],
    'mkdir over file': seeded('61', 'MkDir P("a.txt")'),
    'mkdir empty': ['MkDir ""'],
    'rmdir': ['MkDir P("d")', 'RmDir P("d")', 'Say Dir(P("d"), vbDirectory) = ""'],
    'rmdir missing': ['RmDir P("d")'],
    'rmdir full': ['MkDir P("d")', '@seed d/a.txt 61', 'RmDir P("d")', 'Kill P("d\\a.txt")', 'RmDir P("d")'],
    'rmdir file': seeded('61', 'RmDir P("a.txt")'),
    'rmdir empty': ['RmDir ""'],
    'file copy': seeded('616263', 'FileCopy P("a.txt"), P("b.txt")', 'FileCopy P("a.txt"), P("b.txt")'),
    'file copy missing': ['FileCopy P("none.txt"), P("b.txt")'],
    'file copy open': seeded('61', 'Open P("a.txt") For Input As #1', 'FileCopy P("a.txt"), P("b.txt")', 'Close #1'),
    'file copy open output': ['Open P("a.txt") For Output As #1', 'FileCopy P("a.txt"), P("b.txt")', 'Close #1'],
    'file copy onto open': ['@seed a.txt 61', 'Open P("b.txt") For Output As #1', 'FileCopy P("a.txt"), P("b.txt")', 'Close #1'],
    'file copy pattern': ['@seed a.txt 61', 'FileCopy P("*.txt"), P("b.txt")'],
    'file copy no folder': seeded('61', 'FileCopy P("a.txt"), P("none\\b.txt")'),
    'name': seeded('61', 'Name P("a.txt") As P("b.txt")', 'Say Dir(P("*"))'),
    'name missing': ['Name P("none.txt") As P("b.txt")'],
    'name exists': ['@seed a.txt 61', '@seed b.txt 62', 'Name P("a.txt") As P("b.txt")'],
    'name folder': ['MkDir P("d")', 'Name P("d") As P("e")', 'Say Dir(P("e"), vbDirectory)'],
    'name move': ['MkDir P("d")', '@seed a.txt 61', 'Name P("a.txt") As P("d\\a.txt")', 'Say Dir(P("d\\*"))', 'Say Dir(P("*"))'],
    'name open': seeded('61', 'Open P("a.txt") For Input As #1', 'Name P("a.txt") As P("b.txt")', 'Close #1'),
    'name no folder': seeded('61', 'Name P("a.txt") As P("none\\b.txt")'),
    'file len': seeded('616263', 'Say FileLen(P("a.txt"))', 'Say TypeName(FileLen(P("a.txt")))'),
    'file len missing': ['Say FileLen(P("none.txt"))'],
    'file len folder': ['MkDir P("d")', 'Say FileLen(P("d"))'],
    'file len open': ['Open P("a.txt") For Output As #1', 'Print #1, "abc"', 'Say FileLen(P("a.txt"))', 'Close #1',
                      'Say FileLen(P("a.txt"))'],
    'file date': seeded('61', 'Say TypeName(FileDateTime(P("a.txt")))', 'Say Abs(FileDateTime(P("a.txt")) - Now) < 1'),
    'file date missing': ['Say FileDateTime(P("none.txt"))'],
    'get attr': ['@seed a.txt 61', 'Say GetAttr(P("a.txt"))', 'MkDir P("d")', 'Say GetAttr(P("d"))',
                 'Say TypeName(GetAttr(P("d")))', 'Say GetAttr(P("none.txt"))'],
    'set attr': seeded('61', 'SetAttr P("a.txt"), vbReadOnly', 'Say GetAttr(P("a.txt"))', 'SetAttr P("a.txt"), vbNormal',
                       'Say GetAttr(P("a.txt"))', 'SetAttr P("a.txt"), vbArchive', 'Say GetAttr(P("a.txt"))'),
    'set attr hidden': seeded('61', 'SetAttr P("a.txt"), vbHidden', 'Say GetAttr(P("a.txt"))', 'Say Dir(P("*"))',
                              'Say Dir(P("*"), vbHidden)', 'SetAttr P("a.txt"), vbNormal'),
    'set attr missing': ['SetAttr P("none.txt"), vbNormal'],
    'set attr bad': seeded('61', 'SetAttr P("a.txt"), 1024'),
    'set attr open': seeded('61', 'Open P("a.txt") For Input As #1', 'SetAttr P("a.txt"), vbReadOnly', 'Close #1',
                            'SetAttr P("a.txt"), vbNormal'),
    # Dir
    'dir walk': ['@seed b.txt 61', '@seed a.txt 61', '@seed C.txt 61', 't = Dir(P("*"))', 'Do While t <> ""', 'Say t',
                 't = Dir()', 'Loop', 'Say Dir()'],
    'dir none left': ['@seed a.txt 61', 'Say Dir(P("*"))', 'Say Dir', 'Say Dir', 'Say Dir'],
    'dir fresh': ['Say Dir()'],
    'dir one': ['@seed a.txt 61', 'Say Dir(P("a.txt"))', 'Say Dir(P("A.TXT"))', 'Say Dir(P("none.txt"))'],
    'dir pattern': ['@seed a.txt 61', '@seed b.log 61', '@seed ab.txt 61', 'Say Dir(P("?.txt"))', 'Say Dir', 'Say Dir(P("*.log"))',
                    'Say Dir', 'Say Dir(P("a*"))', 'Say Dir', 'Say Dir', 'Say Dir(P("*.*"))'],
    'dir folders': ['@seed a.txt 61', 'MkDir P("d")', 't = Dir(P("*"), vbDirectory)', 'Do While t <> ""', 'Say t', 't = Dir()',
                    'Loop', 'Say Dir(P("*"))', 'Say Dir'],
    'dir slash': ['MkDir P("d")', '@seed d/a.txt 61', 'Say Dir(P("d\\"))', 'Say Dir(P("d"))', 'Say Dir(P("d"), vbDirectory)',
                  'Say Dir(P("d\\"), vbDirectory)'],
    'dir no folder': ['Say Dir(P("none\\a.txt"))', 'Say Dir(P("none\\"))'],
    'dir empty': ['Say TypeName(Dir(""))', 'Say Dir("", vbDirectory) <> ""'],
    'dir attributes': ['@seed a.txt 61', 'Say Dir(P("a.txt"), vbNormal)', 'Say Dir(P("a.txt"), vbReadOnly)',
                       'Say Dir(P("a.txt"), vbArchive)', 'Say Dir(P("a.txt"), vbHidden)', 'Say Dir(P("a.txt"), vbSystem)',
                       'Say Dir(P("a.txt"), vbVolume)', 'Say Dir(P("a.txt"), 1024)'],
    'dir null': ['Say Dir(Null)'],
    'dir while kill': ['@seed a.txt 61', '@seed b.txt 61', '@seed c.txt 61', 't = Dir(P("*"))', 'Kill P("b.txt")', 'Say t',
                       'Say Dir', 'Say Dir'],
    'dir type': ['Say TypeName(Dir(P("*")))'],
    'dir dollar': ['@seed a.txt 61', 'Say Dir$(P("*"))', 'Say TypeName(Dir$(P("*")))'],
    # the current folder
    'chdir': ['old = CurDir', 'ChDir P("")', 'Say Right(CurDir, 0) = ""', 'Say StrComp(CurDir, Left(P(""), Len(P("")) - 1), vbTextCompare)',
              'Open "b.txt" For Output As #1', 'Close #1', 'Say Dir("b.txt")', 'Say Dir("*")', 'Kill "b.txt"', 'ChDir old',
              'Say CurDir = old'],
    'chdir missing': ['ChDir P("none")'],
    'chdir file': seeded('61', 'ChDir P("a.txt")'),
    'chdir empty': ['ChDir ""'],
    'curdir': ['Say TypeName(CurDir)', 'Say TypeName(CurDir())', 'Say TypeName(CurDir$())', 'Say CurDir("") = CurDir',
               'Say CurDir(Left(CurDir, 1)) = CurDir'],
    'curdir bad drive': ['Say CurDir("?")'],
    'environ': ['Say TypeName(Environ("PATH"))', 'Say Environ("PATH") <> ""', 'Say Environ("NO_SUCH_VARIABLE_HERE")',
                'Say TypeName(Environ$("PATH"))', 'Say Environ("")', 'Say Len(Environ(1)) > 0', 'Say Environ(100000)',
                'Say Environ(0)', 'Say Environ(-1)'],
    'environ by number': ['Say InStr(Environ(1), "=") > 0'],
}

_HELPERS = r'''
Private gFolder As String, gSep As String, gCell(1) As Variant

Private Function P(ByVal name As String) As String
    If gSep <> "\" Then name = Replace(name, "\", gSep)
    P = gFolder & gSep & name
End Function

Private Function Esc(ByVal text As Variant) As String
    Dim i As Long, code As Long, out As String
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code < 32 Or code > 126 Or code = 92 Or code = 124 Or code = 64 Then
            out = out & "\u" & Right$("000" & Hex$(code), 4)
        Else
            out = out & Mid$(text, i, 1)
        End If
    Next
    Esc = out
End Function

Private Function Show(ByVal value As Variant) As String
    If IsNull(value) Then
        Show = "Null"
    ElseIf IsEmpty(value) Then
        Show = "Empty"
    ElseIf IsError(value) Then
        Show = "Error " & CStr(CLng(value))
    ElseIf IsDate(value) And VarType(value) = vbDate Then
        Show = "Date " & Format(value, "yyyy-mm-dd hh:nn:ss")
    Else
        Show = TypeName(value) & " " & Esc(CStr(value))
    End If
End Function

Private Function CStrOf(ByVal value As Variant) As String
    If VarType(value) = vbString Then
        CStrOf = value
    Else
        CStrOf = Show(value)
    End If
End Function

Private Sub Mark(ByRef s As String, ByVal number As Long, ByVal description As String)
    s = s & "[E" & number & ":" & Esc(description) & "]"
End Sub
'''


def _statement(text: str) -> list[str]:
    if text.startswith('@seed'):
        return []
    if text.startswith('Dim '):
        return [f'    {text}']
    if text.startswith('Do While'):
        return [f'    {text}']
    if text == 'Loop':
        return ['    Err.Clear', '    Loop']
    if text.startswith('Say '):
        return [f'    Err.Clear: v = {text[4:]}: If Err.Number Then Mark s, Err.Number, Err.Description Else s = s & "|" & Esc(CStrOf(v))']
    return [f'    Err.Clear: {text}: If Err.Number Then Mark s, Err.Number, Err.Description']


def _probe() -> str:
    lines = [_HELPERS]
    names = list(CASES)
    for index, name in enumerate(names):
        lines += [f'Private Function Case{index}() As String', '    Dim s As String, v As Variant, t As Variant, n As Variant, old As Variant',
                  '    Dim guard As Long', '    On Error Resume Next']
        for text in CASES[name]:
            lines += _statement(text)
        lines += ['    Close', '    Err.Clear', f'    Case{index} = s', 'End Function', '']
    lines += ['Public Function Probe(ByVal folder As String, ByVal sep As String, ByVal which As Long) As String',
              '    gFolder = folder: gSep = sep', '    gCell(1) = Empty', '    Select Case which']
    lines += [f'        Case {index}: Probe = Case{index}()' for index in range(len(names))]
    lines += ['    End Select', 'End Function']
    return '\n'.join(lines) + '\n'


def seed(folder: Path, body: list[str]) -> None:
    """The files a case starts with."""
    for text in body:
        if text.startswith('@seed'):
            _, name, *data = text.split(' ')
            target = folder / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(bytes.fromhex(data[0] if data else ''))


def files_of(folder: Path) -> str:
    """Every file under the folder, with its bytes in hex."""
    found = []
    for path in sorted(folder.rglob('*')):
        relative = path.relative_to(folder).as_posix()
        found.append(f'{relative}/' if path.is_dir() else f'{relative}={path.read_bytes().hex()}')
    return ' '.join(found)


def tidy(folder: Path) -> None:
    """Take read-only off anything left so the temporary folder can go."""
    for path in folder.rglob('*'):
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD | (stat.S_IEXEC if path.is_dir() else 0))


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
                print(f'FAILED {name}: {result.outcome} {result.error} {result.dialogs}', flush=True)
                answers.append([name, 'failed', files_of(folder)])
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
