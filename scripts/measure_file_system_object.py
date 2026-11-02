"""Scripting.FileSystemObject in live Excel: the file, folder and text stream members, and the errors they raise.

Built like scripts/measure_file_statements.py, and run the same way: each case has its own empty folder, the answer is
what it says and every file it leaves, the folder's path is masked, and the harness's line numbers stay off. fso is made
for each case, and the object variables st, o and o2 are there for it to use.

    python scripts/measure_file_system_object.py

writes tests/fixtures/file_system_object.json, which tests/test_vba_file_system_object.py replays.
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

OUT = Path(__file__).resolve().parents[1] / 'tests/fixtures/file_system_object.json'


def stream(data: str, *body: str, opening: str = 'Set st = fso.OpenTextFile(P("a.txt"))') -> list[str]:
    return [f'@seed a.txt {data}', opening, *body, 'st.Close']


READ_ALL = ['Say st.AtEndOfStream', 'Say st.ReadAll', 'Say st.AtEndOfStream']


def walk(*names: str) -> list[str]:
    return [f'@seed {name} 61' for name in names]


CASES: dict[str, list[str]] = {
    # exists
    'exists': ['@seed a.txt 61', 'MkDir P("d")', 'Say fso.FileExists(P("a.txt"))', 'Say fso.FileExists(P("A.TXT"))', 'Say fso.FileExists(P("d"))',
               'Say fso.FileExists(P("none"))', 'Say fso.FolderExists(P("d"))', 'Say fso.FolderExists(P("a.txt"))', 'Say fso.FolderExists(P("none"))',
               'Say fso.FolderExists(P("d\\"))', 'Say fso.FileExists("")', 'Say fso.FolderExists("")', 'Say fso.FileExists(Null)', 'Say fso.FolderExists(Null)'],
    # text streams: writing
    'create': ['Set st = fso.CreateTextFile(P("a.txt"))', 'Say TypeName(st)', 'st.Write "abc"', 'st.WriteLine "def"', 'st.WriteLine', 'st.WriteBlankLines 2', 'st.WriteLine 5',
               'st.Write 1.5', 'st.Write True', 'st.Close'],
    'create line and column': ['Set st = fso.CreateTextFile(P("a.txt"))', 'Say st.Line', 'Say st.Column', 'st.Write "abc"', 'Say st.Line', 'Say st.Column', 'st.WriteLine "d"', 'Say st.Line', 'Say st.Column',
                               'st.Write "x" & vbCrLf & "yz"', 'Say st.Line', 'Say st.Column', 'st.Write "x" & vbLf', 'Say st.Line', 'Say st.Column', 'st.Write "x" & vbCr', 'Say st.Line', 'Say st.Column', 'st.Close'],
    'create overwrite': ['@seed a.txt 61', 'Set st = fso.CreateTextFile(P("a.txt"))', 'st.Close'],
    'create no overwrite': ['@seed a.txt 61', 'Set st = fso.CreateTextFile(P("a.txt"), False)', 'Say TypeName(st)'],
    'create overwrite named': ['@seed a.txt 61', 'Set st = fso.CreateTextFile(P("a.txt"), Overwrite:=True)', 'st.Close'],
    'create no folder': ['Set st = fso.CreateTextFile(P("none\\a.txt"))'],
    'create over folder': ['MkDir P("d")', 'Set st = fso.CreateTextFile(P("d"))'],
    'create read only': ['@seed a.txt 61', 'SetAttr P("a.txt"), vbReadOnly', 'Set st = fso.CreateTextFile(P("a.txt"))', 'SetAttr P("a.txt"), vbNormal'],
    'create empty name': ['Set st = fso.CreateTextFile("")'],
    'create unicode': ['Set st = fso.CreateTextFile(P("a.txt"), True, True)', 'st.WriteLine "a" & ChrW(233) & ChrW(1234)', 'st.Close'],
    'create accented': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.WriteLine ChrW(233)', 'st.WriteLine ChrW(8364)', 'st.WriteLine ChrW(1234)', 'st.WriteLine "after"', 'st.Close'],
    'create line feed': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.Write "a" & vbLf & "b" & vbCr & "c"', 'st.Close'],
    'create write null': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.Write Null', 'st.WriteLine Null', 'st.Write Empty', 'st.WriteLine Empty', 'st.Close'],
    'create negative blanks': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.WriteBlankLines -1', 'st.WriteBlankLines 0', 'st.Close'],
    'create read': ['Set st = fso.CreateTextFile(P("a.txt"))', 'Say st.ReadLine', 'Say st.Read(1)', 'Say st.ReadAll', 'st.Skip 1', 'st.SkipLine', 'Say st.AtEndOfStream', 'Say st.AtEndOfLine', 'st.Close'],
    'create twice': ['Set st = fso.CreateTextFile(P("a.txt"))', 'Set o = fso.CreateTextFile(P("a.txt"))', 'st.Close'],
    'visible before close': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.Write "abc"', 'Say FileLen(P("a.txt"))', 'st.Close', 'Say FileLen(P("a.txt"))'],
    'after close': ['Set st = fso.CreateTextFile(P("a.txt"))', 'st.Close', 'st.Write "a"', 'Say st.Line', 'Say st.AtEndOfStream', 'st.Close'],
    # text streams: reading
    'read all': stream('610D0A620A630D64', *READ_ALL),
    'read all ctrl z': stream('611A62', *READ_ALL),
    'read all nul': stream('610062', *READ_ALL),
    'read all empty': stream('', 'Say st.AtEndOfStream', 'Say st.AtEndOfLine', 'Say st.ReadAll'),
    'read lines': stream('610D0A620A630D640D0A0D0A65', 'Do While Not st.AtEndOfStream', 'Say st.ReadLine', 'Say st.Line', 'Loop', 'Say st.ReadLine', 'Say st.Line'),
    'read line latin': stream('E98081FF', 'Say st.ReadLine'),
    'read chars': stream('6162630D0A64', 'Say st.Read(2)', 'Say st.Column', 'Say st.Read(2)', 'Say st.Line', 'Say st.Column', 'Say st.Read(1)', 'Say st.Line', 'Say st.Column', 'Say st.Read(9)', 'Say st.AtEndOfStream', 'Say st.Read(1)'),
    'read bad count': stream('61', 'Say st.Read(0)', 'Say st.Read(-1)'),
    'skip': stream('6162630D0A64650D0A66', 'st.Skip 2', 'Say st.Column', 'Say st.Read(1)', 'st.SkipLine', 'Say st.Line', 'Say st.ReadLine', 'st.Skip 0', 'st.Skip -1', 'Say st.Read(1)', 'st.SkipLine', 'st.SkipLine', 'st.Skip 1', 'Say st.AtEndOfStream'),
    'at end of line': stream('610D0A0D0A62', 'Say st.AtEndOfLine', 'Say st.Read(1)', 'Say st.AtEndOfLine', 'st.SkipLine', 'Say st.AtEndOfLine', 'st.SkipLine', 'Say st.AtEndOfLine', 'Say st.Read(1)', 'Say st.AtEndOfLine', 'Say st.AtEndOfStream'),
    'at end lf': stream('610A62', 'Say st.Read(1)', 'Say st.AtEndOfLine', 'Say st.ReadLine', 'Say st.Line'),
    'open missing': ['Set st = fso.OpenTextFile(P("none"))', 'Say TypeName(st)'],
    'open no folder': ['Set st = fso.OpenTextFile(P("none\\a.txt"))'],
    'open folder': ['MkDir P("d")', 'Set st = fso.OpenTextFile(P("d"))'],
    'open empty': ['Set st = fso.OpenTextFile("")'],
    'open write': stream('61', 'st.Write "b"', 'Say st.ReadLine', opening='Set st = fso.OpenTextFile(P("a.txt"), 2)'),
    'open append': stream('61', 'st.Write "b"', 'Say st.Line', 'Say st.Column', 'Say st.ReadLine', opening='Set st = fso.OpenTextFile(P("a.txt"), 8)'),
    'open read write': stream('61', 'st.Write "b"', opening='Set st = fso.OpenTextFile(P("a.txt"), 1)'),
    'open bad mode': ['@seed a.txt 61', 'Set st = fso.OpenTextFile(P("a.txt"), 3)', 'Set st = fso.OpenTextFile(P("a.txt"), 0)', 'Set st = fso.OpenTextFile(P("a.txt"), 4)'],
    'open create': ['Set st = fso.OpenTextFile(P("a.txt"), 2, True)', 'st.Write "b"', 'st.Close', 'Set st = fso.OpenTextFile(P("b.txt"), 8, True)', 'st.Close', 'Set st = fso.OpenTextFile(P("c.txt"), 1, True)', 'Say st.AtEndOfStream', 'st.Close'],
    'open no create': ['Set st = fso.OpenTextFile(P("a.txt"), 2, False)', 'Set st = fso.OpenTextFile(P("a.txt"), 8)'],
    'open formats': ['@seed a.txt FFFE61006200', '@seed b.txt 6162', 'Set st = fso.OpenTextFile(P("a.txt"), 1, False, -1)', 'Say st.ReadAll', 'st.Close', 'Set st = fso.OpenTextFile(P("a.txt"), 1, False, 0)', 'Say st.ReadAll', 'st.Close',
                     'Set st = fso.OpenTextFile(P("a.txt"), 1, False, -2)', 'Say st.ReadAll', 'st.Close', 'Set st = fso.OpenTextFile(P("b.txt"), 1, False, -2)', 'Say st.ReadAll', 'st.Close', 'Set st = fso.OpenTextFile(P("b.txt"), 1, False, -1)', 'Say st.ReadAll', 'st.Close',
                     'Set st = fso.OpenTextFile(P("a.txt"), 1, False, 5)'],
    'open unicode write': ['Set st = fso.OpenTextFile(P("a.txt"), 2, True, -1)', 'st.Write "a"', 'st.Close', 'Set st = fso.OpenTextFile(P("a.txt"), 8, True, -1)', 'st.Write "b"', 'st.Close'],
    'open default write': ['@seed a.txt FFFE6100', 'Set st = fso.OpenTextFile(P("a.txt"), 8, False, -2)', 'st.Write "b"', 'st.Close', 'Set st = fso.OpenTextFile(P("a.txt"), 2, False, -2)', 'st.Write "b"', 'st.Close'],
    'open twice': ['@seed a.txt 61', 'Set st = fso.OpenTextFile(P("a.txt"))', 'Set o = fso.OpenTextFile(P("a.txt"))', 'Say o.ReadAll', 'Set o2 = fso.OpenTextFile(P("a.txt"), 2)', 'o.Close', 'st.Close'],
    'open with statement': ['@seed a.txt 61', 'Open P("a.txt") For Output As #1', 'Set st = fso.OpenTextFile(P("a.txt"))', 'Close #1', 'Kill P("a.txt")', 'Set st = fso.OpenTextFile(P("a.txt"), 2, True)', 'Kill P("a.txt")', 'st.Close'],
    # GetFile and File
    'get file': ['@seed a.txt 616263', 'Set o = fso.GetFile(P("a.txt"))', 'Say TypeName(o)', 'Say o.Name', 'Say o.Path', 'Say o', 'Say o.Size', 'Say TypeName(o.Size)', 'Say o.Attributes', 'Say TypeName(o.Attributes)', 'Say TypeName(o.DateCreated)',
                 'Say TypeName(o.DateLastModified)', 'Say TypeName(o.DateLastAccessed)', 'Say TypeName(o.ParentFolder)', 'Say o.ParentFolder.Path'],
    'get file case': ['@seed a.txt 61', 'Say fso.GetFile(P("A.TXT")).Name', 'Say fso.GetFile(P("A.TXT")).Path'],
    'get file missing': ['Set o = fso.GetFile(P("none"))', 'Say TypeName(o)'],
    'get file folder': ['MkDir P("d")', 'Set o = fso.GetFile(P("d"))'],
    'get file empty': ['Set o = fso.GetFile("")'],
    'file rename': ['@seed a.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'o.Name = "b.txt"', 'Say o.Name', 'Say o.Path', 'Say o.Size', 'o.Name = "b.txt"', 'o.Name = ""', 'o.Name = "x\\y"'],
    'file rename exists': ['@seed a.txt 61', '@seed b.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'o.Name = "b.txt"', 'Say o.Name'],
    'file attributes': ['@seed a.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'o.Attributes = 1', 'Say o.Attributes', 'Say GetAttr(P("a.txt"))', 'o.Attributes = 0', 'Say o.Attributes', 'o.Attributes = 16', 'Say o.Attributes', 'o.Attributes = 2048', 'Say o.Attributes', 'o.Attributes = 0'],
    'file gone': ['@seed a.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'Kill P("a.txt")', 'Say o.Name', 'Say o.Path', 'Say o.Size', 'Say o.Attributes', 'o.Delete', 'Say TypeName(o.ParentFolder)'],
    'file grows': ['@seed a.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'Say o.Size', 'Set st = fso.OpenTextFile(P("a.txt"), 8)', 'st.Write "bc"', 'st.Close', 'Say o.Size'],
    'file delete': ['@seed a.txt 61', 'Set o = fso.GetFile(P("a.txt"))', 'o.Delete', 'Say fso.FileExists(P("a.txt"))', 'Say o.Name', 'o.Delete'],
    'file delete read only': ['@seed a.txt 61', 'SetAttr P("a.txt"), vbReadOnly', 'Set o = fso.GetFile(P("a.txt"))', 'o.Delete', 'Say fso.FileExists(P("a.txt"))', '@seed b.txt 61', 'SetAttr P("b.txt"), vbReadOnly', 'Set o = fso.GetFile(P("b.txt"))', 'o.Delete True', 'Say fso.FileExists(P("b.txt"))', 'SetAttr P("a.txt"), vbNormal'],
    'file copy': ['@seed a.txt 616263', 'MkDir P("d")', 'Set o = fso.GetFile(P("a.txt"))', 'o.Copy P("b.txt")', 'o.Copy P("d\\")', 'o.Copy P("d")', 'o.Copy P("b.txt"), False', 'o.Copy P("b.txt"), True', 'o.Copy P("none\\b.txt")', 'Say o.Name'],
    'file move': ['@seed a.txt 616263', 'MkDir P("d")', 'Set o = fso.GetFile(P("a.txt"))', 'o.Move P("b.txt")', 'Say o.Name', 'Say o.Path', 'o.Move P("d\\")', 'Say o.Path', 'o.Move P("none\\b.txt")', 'Say o.Path'],
    'file stream': ['@seed a.txt 616263', 'Set o = fso.GetFile(P("a.txt"))', 'Set st = o.OpenAsTextStream', 'Say st.ReadAll', 'st.Close', 'Set st = o.OpenAsTextStream(2)', 'st.Write "z"', 'st.Close', 'Set st = o.OpenAsTextStream(8, -1)', 'st.Write "y"', 'st.Close', 'Set st = o.OpenAsTextStream(3)'],
    # GetFolder and Folder
    'get folder': [*walk('d/a.txt', 'd/b.txt', 'd/e/c.txt'), 'Set o = fso.GetFolder(P("d"))', 'Say TypeName(o)', 'Say o.Name', 'Say o.Path', 'Say o', 'Say o.Size', 'Say TypeName(o.Size)', 'Say o.Attributes', 'Say o.IsRootFolder', 'Say TypeName(o.ParentFolder)', 'Say o.ParentFolder.Path',
                   'Say TypeName(o.DateCreated)', 'Say TypeName(o.Files)', 'Say TypeName(o.SubFolders)', 'Say o.Files.Count', 'Say o.SubFolders.Count', 'Say TypeName(o.Files.Count)'],
    'get folder slash': ['MkDir P("d")', 'Say fso.GetFolder(P("d\\")).Path', 'Say fso.GetFolder(P("d\\")).Name', 'Say fso.GetFolder(P("D")).Name'],
    'get folder missing': ['Set o = fso.GetFolder(P("none"))'],
    'get folder file': ['@seed a.txt 61', 'Set o = fso.GetFolder(P("a.txt"))'],
    'get folder empty': ['Set o = fso.GetFolder("")'],
    'files': [*walk('d/b.txt', 'd/a.txt', 'd/C.txt', 'd/e/x.txt'), 'Set o = fso.GetFolder(P("d"))', 'Dim f', 'For Each f In o.Files', 'Say f.Name', 'Next', 'Say o.Files.Item("a.txt").Name', 'Say o.Files("A.TXT").Name', 'Say o.Files.Item("none").Name', 'Say o.Files.Item(1).Name', 'Say o.Files(1).Name'],
    'sub folders': [*walk('d/b/x.txt', 'd/a/x.txt', 'd/C/x.txt', 'd/f.txt'), 'Set o = fso.GetFolder(P("d"))', 'Dim f', 'For Each f In o.SubFolders', 'Say f.Name', 'Next', 'Say o.SubFolders.Item("a").Name', 'Say o.SubFolders("A").Path', 'Say o.SubFolders.Item("none").Name', 'Say o.SubFolders.Item(1).Name', 'Say TypeName(o.SubFolders("a").Files)'],
    'files live': ['@seed d/a.txt 61', 'Set o = fso.GetFolder(P("d"))', 'Set o2 = o.Files', 'Say o2.Count', '@seed-now', 'Kill P("d\\a.txt")', 'Say o2.Count'],
    'sub folders add': ['MkDir P("d")', 'Set o = fso.GetFolder(P("d"))', 'Set o2 = o.SubFolders.Add("e")', 'Say TypeName(o2)', 'Say o2.Name', 'Say o.SubFolders.Count', 'Set o2 = o.SubFolders.Add("e")', 'Set o2 = o.SubFolders.Add("")', 'Set o2 = o.SubFolders.Add("x\\y")'],
    'folder create file': ['MkDir P("d")', 'Set o = fso.GetFolder(P("d"))', 'Set st = o.CreateTextFile("a.txt")', 'Say TypeName(st)', 'st.Write "x"', 'st.Close', 'Set st = o.CreateTextFile("a.txt", False)', 'Set st = o.CreateTextFile("b.txt", True, True)', 'st.Close'],
    'folder rename': ['MkDir P("d")', 'Set o = fso.GetFolder(P("d"))', 'o.Name = "e"', 'Say o.Name', 'Say o.Path', 'Say fso.FolderExists(P("e"))', 'MkDir P("f")', 'o.Name = "f"'],
    'folder attributes': ['MkDir P("d")', 'Set o = fso.GetFolder(P("d"))', 'o.Attributes = 1', 'Say o.Attributes', 'o.Attributes = 0', 'Say o.Attributes'],
    'folder delete': ['@seed d/e/a.txt 61', 'Set o = fso.GetFolder(P("d"))', 'o.Delete', 'Say fso.FolderExists(P("d"))', 'Say o.Name', 'o.Delete'],
    'folder copy': ['@seed d/a.txt 61', 'Set o = fso.GetFolder(P("d"))', 'o.Copy P("e")', 'o.Copy P("e")', 'o.Copy P("e"), False', 'o.Copy P("f\\")', 'o.Copy P("g\\h")'],
    'folder move': ['@seed d/a.txt 61', 'MkDir P("e")', 'Set o = fso.GetFolder(P("d"))', 'o.Move P("f")', 'Say o.Path', 'o.Move P("e\\")', 'Say o.Path', 'o.Move P("none\\g")'],
    'root': ['Set o = fso.GetFolder(Left(P(""), 3))', 'Say o.IsRootFolder', 'Say o.Name', 'Say TypeName(o.ParentFolder)', 'Say o.ParentFolder Is Nothing'],
    # the rest of FileSystemObject
    'create folder': ['Set o = fso.CreateFolder(P("d"))', 'Say TypeName(o)', 'Say o.Name', 'Set o = fso.CreateFolder(P("d"))', 'Set o = fso.CreateFolder(P("none\\e"))', 'Set o = fso.CreateFolder("")', 'Set o = fso.CreateFolder(P("f\\"))', 'Say fso.FolderExists(P("f"))'],
    'create folder over file': ['@seed a.txt 61', 'Set o = fso.CreateFolder(P("a.txt"))'],
    'delete file': ['@seed a.txt 61', 'fso.DeleteFile P("a.txt")', 'fso.DeleteFile P("a.txt")', 'fso.DeleteFile ""'],
    'delete file pattern': [*walk('a.txt', 'b.txt', 'c.log', 'd/x.txt'), 'fso.DeleteFile P("*.txt")', 'fso.DeleteFile P("*.txt")', 'fso.DeleteFile P("?.log")', 'fso.DeleteFile P("d")', 'fso.DeleteFile P("none\\*")'],
    'delete file read only': ['@seed a.txt 61', 'SetAttr P("a.txt"), vbReadOnly', 'fso.DeleteFile P("a.txt")', 'fso.DeleteFile P("a.txt"), False', 'fso.DeleteFile P("a.txt"), True'],
    'delete file open': ['@seed a.txt 61', 'Open P("a.txt") For Input As #1', 'fso.DeleteFile P("a.txt")', 'Close #1', '@seed b.txt 61', 'Set st = fso.OpenTextFile(P("b.txt"))', 'fso.DeleteFile P("b.txt")', 'st.Close'],
    'delete folder': ['@seed d/e/a.txt 61', 'fso.DeleteFolder P("d")', 'fso.DeleteFolder P("d")', 'fso.DeleteFolder ""', 'fso.DeleteFolder P("d\\")'],
    'delete folder pattern': [*walk('da/x.txt', 'db/x.txt', 'e/x.txt'), 'fso.DeleteFolder P("d*")', 'fso.DeleteFolder P("d*")'],
    'delete folder slash': ['MkDir P("d")', 'fso.DeleteFolder P("d\\")', 'Say fso.FolderExists(P("d"))'],
    'delete folder read only': ['@seed d/a.txt 61', 'SetAttr P("d\\a.txt"), vbReadOnly', 'fso.DeleteFolder P("d")', 'Say fso.FolderExists(P("d"))', 'fso.DeleteFolder P("d"), True', 'Say fso.FolderExists(P("d"))'],
    'copy file': [*walk('a.txt', 'b.txt'), 'MkDir P("d")', 'fso.CopyFile P("a.txt"), P("c.txt")', 'fso.CopyFile P("a.txt"), P("c.txt")', 'fso.CopyFile P("a.txt"), P("c.txt"), False', 'fso.CopyFile P("a.txt"), P("d\\")', 'fso.CopyFile P("a.txt"), P("d")', 'fso.CopyFile P("none"), P("e.txt")', 'fso.CopyFile P("a.txt"), P("none\\e.txt")', 'fso.CopyFile P("a.txt"), ""', 'fso.CopyFile "", P("e.txt")'],
    'copy file pattern': [*walk('a.txt', 'b.txt', 'c.log'), 'MkDir P("d")', 'fso.CopyFile P("*.txt"), P("d")', 'fso.CopyFile P("*.txt"), P("d"), False', 'fso.CopyFile P("*.none"), P("d")', 'fso.CopyFile P("*.txt"), P("none\\")', 'fso.CopyFile P("*.txt"), P("e.txt")'],
    'copy file read only': ['@seed a.txt 61', '@seed b.txt 62', 'SetAttr P("b.txt"), vbReadOnly', 'fso.CopyFile P("a.txt"), P("b.txt")', 'SetAttr P("b.txt"), vbNormal', 'SetAttr P("a.txt"), vbReadOnly', 'fso.CopyFile P("a.txt"), P("c.txt")', 'Say GetAttr(P("c.txt"))', 'SetAttr P("a.txt"), vbNormal', 'SetAttr P("c.txt"), vbNormal'],
    'copy file over itself': ['@seed a.txt 61', 'fso.CopyFile P("a.txt"), P("a.txt")'],
    'copy file over folder': ['@seed a.txt 61', 'MkDir P("d")', 'fso.CopyFile P("a.txt"), P("d")', 'MkDir P("d\\a.txt")', 'fso.CopyFile P("a.txt"), P("d\\")'],
    'copy folder': ['@seed d/a.txt 61', '@seed d/e/b.txt 62', 'fso.CopyFolder P("d"), P("f")', 'fso.CopyFolder P("d"), P("f")', 'fso.CopyFolder P("d"), P("f"), False', 'fso.CopyFolder P("d"), P("g\\")', 'fso.CopyFolder P("none"), P("h")', 'fso.CopyFolder P("d"), P("none\\x\\y")', 'fso.CopyFolder "", P("h")', 'fso.CopyFolder P("d"), ""'],
    'copy folder pattern': [*walk('da/x.txt', 'db/y.txt', 'e/z.txt'), 'MkDir P("t")', 'fso.CopyFolder P("d*"), P("t")', 'fso.CopyFolder P("*none"), P("t")'],
    'copy folder over itself': ['@seed d/a.txt 61', 'fso.CopyFolder P("d"), P("d")', 'fso.CopyFolder P("d"), P("d\\e")'],
    'move file': ['@seed a.txt 61', '@seed b.txt 62', 'MkDir P("d")', 'fso.MoveFile P("a.txt"), P("c.txt")', 'fso.MoveFile P("a.txt"), P("c.txt")', 'fso.MoveFile P("b.txt"), P("c.txt")', 'fso.MoveFile P("b.txt"), P("d\\")', 'fso.MoveFile P("c.txt"), P("none\\e.txt")', 'fso.MoveFile "", P("e.txt")', 'fso.MoveFile P("c.txt"), P("d")'],
    'move file pattern': [*walk('a.txt', 'b.txt', 'c.log'), 'MkDir P("d")', 'fso.MoveFile P("*.txt"), P("d")', 'fso.MoveFile P("*.txt"), P("d")'],
    'move folder': ['@seed d/a.txt 61', 'MkDir P("e")', 'fso.MoveFolder P("d"), P("f")', 'fso.MoveFolder P("d"), P("f")', 'fso.MoveFolder P("f"), P("e")', 'fso.MoveFolder P("f"), P("e\\")', 'fso.MoveFolder P("e\\f"), P("none\\g")', 'fso.MoveFolder "", P("g")'],
    'move folder pattern': [*walk('da/x.txt', 'db/y.txt'), 'MkDir P("t")', 'fso.MoveFolder P("d*"), P("t")'],
    'absolute': ['old = CurDir', 'ChDir P("")', 'Say fso.GetAbsolutePathName("a.txt")', 'Say fso.GetAbsolutePathName(".")', 'Say fso.GetAbsolutePathName("..\\b")', 'Say fso.GetAbsolutePathName("")', 'Say fso.GetAbsolutePathName("a\\..\\b\\")', 'Say fso.GetAbsolutePathName(P("x"))', 'ChDir old'],
    'temp name': ['t = fso.GetTempName', 'Say Len(t)', 'Say Left(t, 3)', 'Say Right(t, 4)', 'Say TypeName(t)'],
    'special folder': ['Say TypeName(fso.GetSpecialFolder(0))', 'Say TypeName(fso.GetSpecialFolder(1))', 'Say fso.GetSpecialFolder(0).Name'],
    'temp folder': ['Say TypeName(fso.GetSpecialFolder(2))', 'Say fso.GetSpecialFolder(3)', 'Say fso.FolderExists(fso.GetSpecialFolder(2).Path)'],
    'drive exists': ['Say TypeName(fso.DriveExists("C:"))', 'Say fso.DriveExists("")', 'Say fso.DriveExists(Left(P(""), 2))', 'Say fso.DriveExists("?:")', ],
    'nulls': ['Set o = fso.GetFile(Null)', 'Set o = fso.GetFolder(Null)', 'Set o = fso.CreateFolder(Null)', 'fso.DeleteFile Null', 'Set st = fso.OpenTextFile(Null)', 'Set st = fso.CreateTextFile(Null)'],
    'names of folders': ['@seed d/a.txt 61', 'fso.DeleteFile P("d")', 'fso.DeleteFile P("none\\*")', 'fso.DeleteFile P("none\\x")', 'fso.DeleteFile P("d\\")'],
    'folder onto itself': ['@seed d/a.txt 61', 'fso.CopyFolder P("d"), P("d")', 'fso.CopyFolder P("d"), P("d\\e")', 'fso.CopyFolder P("d"), P("d\\")', 'fso.CopyFolder P("d\\"), P("f")',
                           'fso.MoveFolder P("d"), P("d\\e")', 'fso.MoveFolder P("d"), P("d")'],
    'kinds': ['Say TypeName(fso.GetFolder(P("")).Files)', 'Say TypeName(fso)', 'Say TypeName(fso.GetFileName("a"))', 'Say TypeName(fso.FileExists("a"))'],
}


def _probe() -> str:
    lines = [_HELPERS]
    names = list(CASES)
    for index, name in enumerate(names):
        lines += [f'Private Function Case{index}() As String', '    Dim s As String, v As Variant, t As Variant, old As Variant',
                  '    Dim fso As Object, st As Object, o As Object, o2 As Object', '    On Error Resume Next',
                  '    Set fso = CreateObject("Scripting.FileSystemObject")']
        for text in CASES[name]:
            if text == '@seed-now':
                continue
            if text.startswith(('For Each', 'Next')):
                lines.append(f'    {text}')
            else:
                lines += _statement(text)
        lines += ['    If Not st Is Nothing Then st.Close', '    Close', '    Err.Clear', f'    Case{index} = s', 'End Function', '']
    lines += ['Public Function Probe(ByVal folder As String, ByVal sep As String, ByVal which As Long) As String',
              '    gFolder = folder: gSep = sep', '    Select Case which']
    lines += [f'        Case {index}: Probe = Case{index}()' for index in range(len(names))]
    lines += ['    End Select', 'End Function']
    return '\n'.join(lines) + '\n'


def masked(text: str, folder: Path) -> str:
    """The answer with the folder and its parent taken out, and one kind of separator."""
    for path, word in ((str(folder), 'FOLDER'), (str(folder.parent), 'ROOT')):
        text = text.replace(path, word).replace(path.replace('\\', '\\u005C'), word)
    return text.replace('\\u005C', '/')


def main() -> None:
    source = _probe()
    names = list(CASES)
    answers: list[list[str]] = []
    with tempfile.TemporaryDirectory() as root, ExcelSession(HarnessConfig(lock_wait_s=1200.0)) as excel:
        excel.new_document()
        for index, name in enumerate(names):
            folder = Path(root) / f'c{index}'
            folder.mkdir()
            seed(folder, [text for text in CASES[name] if text != '@seed-now'])
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
              'seeds': {name: [text for text in body if text.startswith('@seed') and text != '@seed-now'] for name, body in CASES.items()},
              'cases': answers}
    OUT.write_text(json.dumps(record, indent=1) + '\n', encoding='utf-8')
    for name, answer, files in answers:
        print(f'{name}: {answer}  ## {files}')


if __name__ == '__main__':
    main()
