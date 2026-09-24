"""Files Excel, Word and PowerPoint save before the first macro exists.

Each application makes a new file and saves it, macro-enabled and in its
binary format, without a line of VBA ever written. The record says what
each file holds where a project would be: the zip entry vbaProject.bin,
or the storages at the root of the binary file.

    python scripts/measure_no_vba.py

writes tests/fixtures/no_vba/: the six files and no_vba.json, which
tests/test_no_vba_project.py replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, PowerPointSession, WordSession

from fixture_workbook import copy_saved
from pyopenvba.cfb import CFB

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "no_vba"
#: Each application: its session, how it makes a file, and each format it saves as, with the file format number.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "add": "Workbooks.Add(1)", "quiet": "Application.DisplayAlerts = False",
              "save": 'SaveAs Filename:="{path}", FileFormat:={format}', "close": "Close False",
              "formats": {"workbook.xlsm": 52, "workbook.xls": 56}},
    "word": {"session": WordSession, "add": "Documents.Add", "quiet": "Application.DisplayAlerts = 0",
             "save": 'SaveAs2 FileName:="{path}", FileFormat:={format}', "close": "Close SaveChanges:=0",
             "formats": {"document.docm": 13, "document.doc": 0}},
    "powerpoint": {"session": PowerPointSession, "add": "Presentations.Add(0)", "quiet": "Application.DisplayAlerts = 1",
                   "save": 'SaveAs "{path}", {format}', "close": "Close",
                   "formats": {"presentation.pptm": 25, "presentation.ppt": 1}},
}


def holds(path: Path) -> dict[str, Any]:
    """What the file holds where a VBA project would be."""
    # Not zipfile.is_zipfile: a binary file from Office 2007 on carries its theme as a zip inside it.
    if path.read_bytes().startswith(b"PK\x03\x04"):
        with zipfile.ZipFile(path) as package:
            return {"container": "zip", "vba_parts": [name for name in package.namelist() if "vba" in name.lower()]}
    cfb = CFB.from_bytes(path.read_bytes())
    return {"container": "cfb", "root_storages": sorted(cfb.list_storages_at()),
            "root_streams": sorted(cfb.list_streams_at())}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp())
    record: dict[str, Any] = {}
    for host in HOSTS.values():
        lines = ["Public Function Make() As String", "Dim d As Object", host["quiet"]]
        for name, number in host["formats"].items():
            lines += [f"Set d = {host['add']}", "d." + host["save"].format(path=folder / name, format=number),
                      f"d.{host['close']}"]
        lines += ['Make = "made"', "End Function"]
        with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
            office.new_document()
            made = office.run_vba("\n".join(lines) + "\n", "Make", timeout=180.0)
            assert made.ok, f"{made.outcome} {made.message}"
        for name in host["formats"]:
            copy_saved(folder / name, OUT / name)
            record[name] = holds(OUT / name)
    (OUT / "no_vba.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, found in record.items():
        print(name, found)


if __name__ == "__main__":
    main()
