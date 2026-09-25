"""A UserForm in a binary .xls and .doc: where Excel and Word keep it, and what they write for it.

Each application makes a binary file whose project holds a standard module
and nothing else, and saves it; then inserts a UserForm through its editor
(VBComponents.Add 3) and saves a second file. The first file is the one the
offline tests start from, tests/fixtures/binary_project/workbook.xls and
document.doc. The record takes the second apart: the storages and streams
at the root and under the project's storage, the class id of each storage
there, the form's designer streams, and the PROJECT stream's lines. The
form keeps the editor's name, UserForm1, so the library can make the same
form.

    python scripts/measure_binary_forms.py

writes tests/fixtures/binary_project/ and tests/fixtures/binary_forms.json,
which tests/test_binary_forms.py replays.
"""

from __future__ import annotations

import json
import shutil
import tempfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, WordSession

from pyopenvba.cfb import CFB

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures" / "binary_project"
OUT = ROOT / "tests" / "fixtures" / "binary_forms.json"
FORM = "UserForm1"
#: Each application: its session, the file it keeps, the storage its project lives in, and its VBA.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "file": "workbook.xls", "storage": "_VBA_PROJECT_CUR",
              "quiet": "Application.DisplayAlerts = False", "add": "Workbooks.Add(-4167)",
              "item": 'Workbooks("{name}")', "save_as": 'SaveAs Filename:="{path}", FileFormat:=56',
              "close": "Close False"},
    "word": {"session": WordSession, "file": "document.doc", "storage": "Macros",
             "quiet": "Application.DisplayAlerts = 0", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:=0', "close": "Close SaveChanges:=0"},
}


def entry(raw: bytes, name: str, kind: int) -> bytes:
    """The 128-byte directory entry named ``name`` of object type ``kind`` (1 storage, 2 stream).

    Entries are sector-aligned, with the UTF-16LE name at 0, its byte length at 64 and the type at 66.
    """
    needle = name.encode("utf-16-le") + b"\x00\x00"
    found = [raw[at:at + 128] for at in range(0, len(raw) - 127, 128)
             if raw[at:at + len(needle)] == needle and raw[at + 64] == len(needle) and raw[at + 66] == kind]
    if len(found) != 1:
        raise ValueError(f"{len(found)} directory entries named {name!r} of type {kind}")
    return found[0]


def tree(cfb: CFB, path: tuple[str, ...]) -> list[str]:
    """Every storage (ending in /) and stream under ``path``, relative to it, depth first in name order."""
    lines = sorted(cfb.list_streams_at(path))
    for storage in sorted(cfb.list_storages_at(path)):
        lines.append(storage + "/")
        lines += [f"{storage}/{line}" for line in tree(cfb, (*path, storage))]
    return lines


def take_apart(path: Path, storage: str) -> dict[str, Any]:
    raw = path.read_bytes()
    cfb = CFB.from_bytes(raw)
    project = (storage,)
    storages = [line[:-1] for line in tree(cfb, project) if line.endswith("/")]
    form = entry(raw, FORM, 1)
    return {
        "root": sorted(cfb.list_storages_at()) + sorted(cfb.list_streams_at()),
        "tree": tree(cfb, project),
        # Only the storages directly under the project's storage: a name repeats deeper down.
        "clsids": {name: entry(raw, name, 1)[80:96].hex() for name in storages if "/" not in name},
        "form_entry": {"state": form[96:100].hex(), "created_set": any(form[100:108]),
                       "modified_set": any(form[108:116])},
        "form_streams": {name: cfb.get_stream_at((*project, FORM), name).hex()
                         for name in sorted(cfb.list_streams_at((*project, FORM)))},
        "project_lines": cfb.get_stream_at(project, "PROJECT").decode("latin-1").split("\r\n"),
    }


def run(office: Any, body: str) -> str:
    done = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n{body}\n"
                          "Step = out\nEnd Function\n", "Step", timeout=120.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    base, formed = folder / host["file"], folder / f"form_{host['file']}"
    quiet = host["quiet"]
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            made = run(office, f"{quiet}\nSet d = {host['add']}\nd.VBProject.VBComponents.Add 1\nout = d.Name")
            item = host["item"].format(name=made)
            run(office, f"{quiet}\nSet d = {item}\nd.{host['save_as'].format(path=base)}\nout = d.Name")
            item = host["item"].format(name=base.name)
            run(office, f"{quiet}\nSet d = {item}\nout = d.VBProject.VBComponents.Add(3).Name")
            run(office, f"{quiet}\nSet d = {item}\nd.{host['save_as'].format(path=formed)}\nd.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            run(office, "\n".join([
                quiet, f"Set d = {host['add']}", "d.VBProject.VBComponents.Add 1",
                f"d.{host['save_as'].format(path=base)}", "out = d.VBProject.VBComponents.Add(3).Name",
                f"d.{host['save_as'].format(path=formed)}", f"d.{host['close']}"]))
    shutil.copyfile(base, FIXTURES / host["file"])
    return {"project_storage": host["storage"], "base_tree": tree(CFB.from_bytes(base.read_bytes()),
                                                                   (host["storage"],)),
            **take_apart(formed, host["storage"])}


def main() -> None:
    FIXTURES.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print(json.dumps(record, indent=1))
    print("kept", folder)


if __name__ == "__main__":
    main()
