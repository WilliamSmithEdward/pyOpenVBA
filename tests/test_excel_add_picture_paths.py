"""Shapes.AddPicture's path, as a macro writes it with backslashes, on every system.

The gate runs before anything on disk is looked at, and the file opened is final_path's one string: off Windows a backslash
is a separator, and .. is resolved before the read, so nothing turns the path round between the check and the open
(see tests/test_vba_sandbox.py and the sandbox's path rule)."""
from __future__ import annotations

import struct
import zlib
from pathlib import Path

import pytest

from pyopenvba.apps.excel import ExcelApplication
from pyopenvba.interpreter import VBASandboxError

SOURCE = '''Public Function Add(ByVal path As String) As String
    Dim sh As Object
    Set sh = ActiveSheet.Shapes.AddPicture(path, msoFalse, msoTrue, 10, 10, -1, -1)
    Add = sh.Name & " " & sh.Width & " " & sh.Height
End Function
'''


def png(path: Path) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    rows = b''.join(b'\x00' + b'\xff\x00\x00' * 8 for _ in range(4))
    path.write_bytes(b''.join([b'\x89PNG\r\n\x1a\n', chunk(b'IHDR', struct.pack('>IIBBBBB', 8, 4, 8, 2, 0, 0, 0)),
                               chunk(b'IDAT', zlib.compress(rows, 9)), chunk(b'IEND', b'')]))


def app(sandbox: bool) -> ExcelApplication:
    made = ExcelApplication(sandbox=sandbox)
    made.add_workbook()
    made.add_module(SOURCE, name='P')
    return made


@pytest.fixture
def folders(tmp_path: Path) -> tuple[Path, Path]:
    granted, away = tmp_path / 'granted', tmp_path / 'away'
    granted.mkdir()
    away.mkdir()
    png(away / 'a.png')
    return granted, away


def test_a_backslashed_path_opens_the_file_it_names(folders: tuple[Path, Path]) -> None:
    _, away = folders
    assert app(False).run('P.Add', str(away) + '\\a.png') == 'Picture 2 6 3'


def test_a_backslashed_route_out_is_the_file_it_reaches_not_a_name_in_the_folder(folders: tuple[Path, Path]) -> None:
    # Off Windows, granted/..\away\a.png would be one odd name inside granted if the backslashes were not separators.
    granted, _ = folders
    assert app(False).run('P.Add', str(granted) + '\\..\\away\\a.png') == 'Picture 2 6 3'


def test_with_the_sandbox_on_a_backslashed_route_out_is_refused_before_the_disk(folders: tuple[Path, Path],
                                                                                monkeypatch: pytest.MonkeyPatch) -> None:
    granted, _ = folders
    looked: list[str] = []

    def read_bytes(self: Path) -> bytes:
        looked.append(str(self))
        return b''

    monkeypatch.setattr(Path, 'read_bytes', read_bytes)
    with pytest.raises(VBASandboxError):
        app(True).run('P.Add', str(granted) + '\\..\\away\\a.png')
    assert not looked
