"""The Microsoft Forms reference Excel, Word and PowerPoint write when a UserForm is inserted.

Each application makes a file with no VBA, inserts a UserForm through its
editor (VBComponents.Add 3) and saves. The record lists the references the
saved project declares and takes the Microsoft Forms reference apart record
by record: REFERENCENAME and its Unicode partner, REFERENCEORIGINAL,
REFERENCECONTROL with its twiddled libid, the name again, and the extended
record with its libid, the type library's GUID and the cookie. The extended
libid names the editor's .exd cache in the Temp folder of whoever saved the
file, so the Temp folder is recorded as %TEMP% and no workbook is kept.

    python scripts/measure_form_reference.py

writes tests/fixtures/form_reference.json, which tests/test_form_reference.py
replays.
"""

from __future__ import annotations

import json
import os
import struct
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, PowerPointSession, WordSession

from pyopenvba.cfb import CFB
from pyopenvba.vba import decompress

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "form_reference.json"
TEMP = str(Path(os.environ["TEMP"]).resolve())
#: Each application: its session, the file's extension and project part, and how its VBA makes and saves a file.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "suffix": "xlsm", "entry": "xl/vbaProject.bin",
              "add": "Workbooks.Add(-4167)", "item": 'Workbooks("{name}")',
              "save_as": 'SaveAs Filename:="{path}", FileFormat:=52', "close": "Close False",
              "quiet": "Application.DisplayAlerts = False"},
    "word": {"session": WordSession, "suffix": "docm", "entry": "word/vbaProject.bin",
             "add": "Documents.Add", "item": 'Documents("{name}")',
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:=13', "close": "Close SaveChanges:=0",
             "quiet": "Application.DisplayAlerts = 0"},
    "powerpoint": {"session": PowerPointSession, "suffix": "pptm", "entry": "ppt/vbaProject.bin",
                   "add": "Presentations.Add(0)", "item": 'Presentations("{name}")',
                   "save_as": 'SaveAs "{path}", 25', "close": "Close", "quiet": "Application.DisplayAlerts = 1"},
}


def dir_records(raw: bytes) -> list[tuple[int, bytes]]:
    """The dir stream's records up to the module section, as (id, body)."""
    records: list[tuple[int, bytes]] = []
    at = 0
    while at + 6 <= len(raw):
        rid, size = struct.unpack_from("<HI", raw, at)
        if rid == 0x000F:
            return records
        end = at + (12 if rid == 0x0009 else 6 + size)
        records.append((rid, raw[at + 6:end]))
        at = end
    raise ValueError("no PROJECTMODULES record")


def sized(body: bytes) -> tuple[str, int]:
    size = int.from_bytes(body[:4], "little")
    return body[4:4 + size].decode("latin-1"), 4 + size


def unhomed(text: str) -> str:
    return text.replace(TEMP, "%TEMP%")


def forms_block(path: Path, entry: str) -> dict[str, Any]:
    with zipfile.ZipFile(path) as package:
        raw = decompress(CFB.from_bytes(package.read(entry)).get_stream("dir"))
    records = dir_records(raw)
    names = [body.decode("latin-1") for rid, body in records if rid == 0x0016]
    start = next(i for i, (rid, body) in enumerate(records) if rid == 0x0016 and body == b"MSForms")
    end = next(i for i in range(start, len(records)) if records[i][0] == 0x0030) + 1
    block = records[start:end]
    ids = [rid for rid, _ in block]
    twiddled, after_twiddled = sized(block[3][1])
    extended, after_extended = sized(block[-1][1])
    body = block[-1][1]
    return {
        "reference_names": list(dict.fromkeys(names)),
        "record_ids": [f"0x{rid:04X}" for rid in ids],
        "name": block[0][1].decode("latin-1"), "name_unicode": block[1][1].decode("utf-16-le"),
        "original_libid": unhomed(block[2][1].decode("latin-1")),
        "twiddled_libid": twiddled, "twiddled_reserved": block[3][1][after_twiddled:].hex(),
        "extended_libid": unhomed(extended),
        "extended_reserved": body[after_extended:after_extended + 6].hex(),
        "type_library_guid": body[after_extended + 6:after_extended + 22].hex(),
        "cookie": int.from_bytes(body[after_extended + 22:after_extended + 26], "little"),
        "extended_size": len(body),
        "last_reference": all(rid != 0x0016 for rid, _ in records[end:]),
    }


def measure(name: str, host: dict[str, Any]) -> dict[str, Any]:
    target = Path(tempfile.mkdtemp()) / f"form.{host['suffix']}"
    making = f"Set d = {host['add']}\nd.VBProject.VBComponents.Add 3\nout = d.Name"
    saving = f"d.{host['save_as'].format(path=target)}\nd.{host['close']}"
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call makes a project and saves it: a call each.
            made = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n"
                                  f"{host['quiet']}\n{making}\nStep = out\nEnd Function\n", "Step", timeout=120.0)
            assert made.ok, f"{name}: {made.outcome} {made.message}"
            steps = [f"Set d = {host['item'].format(name=made.value)}\n{saving}"]
        else:
            # Word and PowerPoint lose the harness's procedures once another file's project changed: one call.
            steps = [f"{making}\n{saving}"]
        for step in steps:
            done = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n"
                                  f"{host['quiet']}\n{step}\nStep = \"ok\"\nEnd Function\n", "Step", timeout=120.0)
            assert done.ok, f"{name}: {done.outcome} {done.message}"
    return forms_block(target, host["entry"])


def main() -> None:
    record = {name: measure(name, host) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(json.dumps(record, indent=1))


if __name__ == "__main__":
    main()
