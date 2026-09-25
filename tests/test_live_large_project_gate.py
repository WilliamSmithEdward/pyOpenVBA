"""Live gate: a project past 7.1 MB the library saves opens in Excel and Word (opt-in).

A UserForm holding a large picture is the usual way a project gets that
big. Each application makes one, an Image showing a 9.7 MB bitmap, and saves
it. The library adds a module and saves the project again, past the 109 FAT
sectors a compound file's header can list, so the writer needs DIFAT
sectors. The application opens that file, lists the new module and shows
the picture at the size it had.

Opt-in per application: RUN_LIVE_EXCEL=1 or RUN_LIVE_WORD=1, on Windows
with it installed. Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import struct
import sys
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, WordFile

#: Each application: the switch that opts in, the library's file type, the extension, the project part, the
#: harness session, the open file as its VBA names it, and how its VBA makes, saves and closes a file.
HOSTS: dict[str, tuple[str, Any, str, str, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "xlsm", "xl/vbaProject.bin", "ExcelSession", "ActiveWorkbook",
              "Application.DisplayAlerts = False\nSet d = Workbooks.Add(-4167)",
              'Application.DisplayAlerts = False\nSet d = Workbooks("{name}")\n'
              'd.SaveAs Filename:="{path}", FileFormat:=52\nd.Close False'),
    "word": ("RUN_LIVE_WORD", WordFile, "docm", "word/vbaProject.bin", "WordSession", "ActiveDocument",
             "Application.DisplayAlerts = 0\nSet d = Documents.Add",
             'd.SaveAs2 FileName:="{path}", FileFormat:=13\nd.Close SaveChanges:=0'),
}
#: What the header of a compound file with 512-byte sectors can map: 109 FAT sectors of 128 sectors each.
HEADER_REACH = 109 * 128 * 512


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live large project gate: set {switch}=1 on Windows with the application installed"))


def _bitmap(path: Path, side: int = 1800) -> None:
    """A 24-bit bitmap of random pixels, ``side`` pixels square, which nothing compresses."""
    row = side * 3
    pixels = os.urandom(row * side)
    header = struct.pack("<2sIHHI", b"BM", 54 + len(pixels), 0, 0, 54)
    info = struct.pack("<IiiHHIIiiII", 40, side, side, 1, 24, 0, len(pixels), 3780, 3780, 0, 0)
    path.write_bytes(header + info + pixels)


def _run(office: Any, body: str) -> str:
    done = office.run_vba("Public Function Step() As String\nDim d As Object, c As Object, k As Object, "
                          f"out As String\n{body}\nStep = out\nEnd Function\n", "Step", timeout=600.0)
    assert done.ok, f"{done.outcome}: {done.message}"
    return str(done.value)


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_a_project_past_the_headers_reach_saves_and_opens(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, suffix, part, session_name, active, making, saving = HOSTS[host]
    picture, target = tmp_path / "large.bmp", tmp_path / f"large.{suffix}"
    _bitmap(picture)
    adding = ("Set c = d.VBProject.VBComponents.Add(3)\n"
              'Set k = c.Designer.Controls.Add("Forms.Image.1", "Image1")\n'
              f'Set k.Picture = LoadPicture("{picture}")\nout = d.Name & "|" & k.Picture.Width')
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if host == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            name, width = _run(office, f"{making}\n{adding}").split("|")
            _run(office, saving.format(name=name, path=target))
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            _, width = _run(office, f"{making}\n{adding}\n{saving.format(path=target)}").split("|")
    with zipfile.ZipFile(target) as package:
        assert len(package.read(part)) > HEADER_REACH
    with opener(target) as office_file:
        office_file.vba_project().add_module("Touched", 'Attribute VB_Name = "Touched"\r\n')
        office_file.save()
    with zipfile.ZipFile(target) as package:
        project = package.read(part)
    assert struct.unpack_from("<I", project, 72)[0] >= 1  # DIFAT sectors the library wrote
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.open_document(target)
        components = f"{active}.VBProject.VBComponents"
        assert office.eval(f'{components}("Touched").Name') == "Touched"
        assert str(office.eval(f'{components}("UserForm1").Designer.Controls("Image1").Picture.Width')) == width
