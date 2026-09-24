"""No template the library ships names a folder on the machine that made it.

Every file create_new() writes, and every workbook the in-memory Excel
starts, is one of the blobs in pyopenvba._templates. Excel records the
folder it saved a file in -- x15ac:absPath in xl/workbook.xml, a
BrtAbsPath15 record in a .xlsb's xl/workbook.bin -- and a VBA project
records where its referenced libraries were found, which can be a user's
Temp folder. Each blob is decoded here, each of its parts decompressed,
and a VBA project's dir stream too, since compression can hide a path
from a scan of the raw bytes.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from pyopenvba import _templates
from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress

NAMES = sorted(name for name in vars(_templates) if name.startswith("EMPTY_") and name.endswith("_BYTES"))
NEEDLES = [text.encode(encoding) for text in ("Users\\", "AppData", "GitHub") for encoding in ("ascii", "utf-16-le")]


def _pieces(data: bytes) -> dict[str, bytes]:
    """The blob's parts, decompressed, and each VBA project's PROJECT and decompressed dir stream."""
    if not data.startswith(b"PK\x03\x04"):
        return {"(the file)": data}
    pieces: dict[str, bytes] = {}
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        for name in package.namelist():
            pieces[name] = package.read(name)
            if name.endswith("vbaProject.bin"):
                cfb = CFB.from_bytes(pieces[name])
                pieces[f"{name}:PROJECT"] = cfb.get_stream("PROJECT")
                pieces[f"{name}:dir"] = decompress(cfb.get_stream("dir"))
    return pieces


def test_every_template_is_looked_at() -> None:
    assert NAMES == ["EMPTY_ACCDB_BYTES", "EMPTY_DOCM_BYTES", "EMPTY_PPTM_BYTES", "EMPTY_XLAM_BYTES",
                     "EMPTY_XLSB_BYTES", "EMPTY_XLSM_BYTES", "EMPTY_XLSX_BYTES"]


@pytest.mark.parametrize("name", NAMES)
def test_no_template_names_a_users_folder(name: str) -> None:
    found = {part: [needle.decode("latin-1") for needle in NEEDLES if needle in data]
             for part, data in _pieces(getattr(_templates, name)).items()}
    assert {part: needles for part, needles in found.items() if needles} == {}
