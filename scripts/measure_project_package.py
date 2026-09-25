"""The PROJECT stream's Package line: when Excel's and Word's editors write it, where, and when it goes.

Each application builds a project step by step through its editor
(VBComponents.Add and Remove) and saves after every step, in its binary
format: a standard module; the first UserForm; a module after it; a
second form later; the first form removed; the last form removed; a form
again. A second project gets its module and two forms before its only
save, a third leaves its Package line last among the declarations and
adds a module after it, and a fourth starts from a project that declares
a form with no Package line, as pyOpenVBA 6.1.2 and earlier wrote one.
The record keeps every saved state's PROJECT stream line by line, the
storages beside VBA/ and the references, and, as a check that the file
type does not matter, the first two states saved macro-enabled too.

    python scripts/measure_project_package.py

writes tests/fixtures/project_package.json, which
tests/test_project_package.py replays. No file is kept: a project with a
form names the Temp folder of whoever saved it.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, WordSession

from pyopenvba import ExcelFile, WordFile
from pyopenvba.cfb import CFB

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "project_package.json"
#: Each application: its session and file type, the storage its binary file keeps the project in, and its VBA.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "opener": ExcelFile, "storage": "_VBA_PROJECT_CUR",
              "binary": ("xls", 56), "macro": ("xlsm", 52, "xl/vbaProject.bin"),
              "quiet": "Application.DisplayAlerts = False", "add": "Workbooks.Add(-4167)",
              "item": 'Workbooks("{name}")', "save_as": 'SaveAs Filename:="{path}", FileFormat:={format}',
              "close": "Close False"},
    "word": {"session": WordSession, "opener": WordFile, "storage": "Macros",
             "binary": ("doc", 0), "macro": ("docm", 13, "word/vbaProject.bin"),
             "quiet": "Application.DisplayAlerts = 0", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:={format}', "close": "Close SaveChanges:=0"},
}
#: The steps, each saved as the state it names; a state in MACRO is saved macro-enabled as well.
CHAIN = [
    ("module", ["Add 1"]),
    ("first_form", ["Add 3"]),
    ("module_after_form", ["Add 1"]),
    ("second_form_later", ["Add 3"]),
    ("first_form_removed", ["Remove UserForm1"]),
    ("last_form_removed", ["Remove UserForm2"]),
    ("form_after_all_removed", ["Add 3"]),
]
AT_ONCE = [("two_forms_at_once", ["Add 1", "Add 3", "Add 3"])]
#: The Package line left last among the declarations, then a module added after it.
PACKAGE_LAST = [("form_declared_last", ["Add 1", "Add 3"]), ("package_left_last", ["Remove UserForm1"]),
                ("module_after_package", ["Add 1"])]
#: Starting from a project pyOpenVBA 6.1.2 and earlier wrote: a form declared with no Package line.
LEGACY = [("legacy_module_added", ["Add 1"]), ("legacy_second_form", ["Add 3"])]
MACRO = {"module", "first_form", "legacy_module_added", "legacy_second_form"}


def vba(action: str) -> str:
    verb, _, what = action.partition(" ")
    if verb == "Add":
        return f"d.VBProject.VBComponents.Add {what}"
    return f'd.VBProject.VBComponents.Remove d.VBProject.VBComponents("{what}")'


def state(path: Path, host: dict[str, Any], macro: bool) -> dict[str, Any]:
    if macro:
        with zipfile.ZipFile(path) as package:
            cfb, root = CFB.from_bytes(package.read(host["macro"][2])), ()
    else:
        cfb, root = CFB.from_bytes(path.read_bytes()), (host["storage"],)
    with host["opener"](path) as office_file:
        references = [reference.name for reference in office_file.references()]
    return {
        "project": cfb.get_stream_at(root, "PROJECT").decode("latin-1").split("\r\n"),
        "storages": sorted(name for name in cfb.list_storages_at(root) if name.casefold() != "vba"),
        "references": references,
    }


def run(office: Any, body: str) -> str:
    done = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n{body}\n"
                          "Step = out\nEnd Function\n", "Step", timeout=120.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def saves(host: dict[str, Any], folder: Path, name: str) -> list[tuple[Path, bool, str]]:
    """The files a state is saved as: the binary one, unless it is a legacy state, and the macro-enabled one
    where MACRO asks."""
    binary, macro = host["binary"], host["macro"]
    out: list[tuple[Path, bool, str]] = []
    if name in MACRO:
        path = folder / f"{name}.{macro[0]}"
        out.append((path, True, host["save_as"].format(path=path, format=macro[1])))
    if not name.startswith("legacy"):
        path = folder / f"{name}.{binary[0]}"
        out.append((path, False, host["save_as"].format(path=path, format=binary[1])))
    return out


def legacy_file(host: dict[str, Any], folder: Path) -> Path:
    """A macro-enabled file whose project declares a form, BaseClass=UserForm1, and no Package line."""
    path = folder / f"legacy.{host['macro'][0]}"
    with host["opener"].create_new(path) as office_file:
        office_file.add_form("UserForm1")
        office_file.save()
    entry = host["macro"][2]
    with zipfile.ZipFile(path) as package:
        parts = {info.filename: package.read(info) for info in package.infolist()}
    cfb = CFB.from_bytes(parts[entry])
    lines = cfb.get_stream("PROJECT").split(b"\r\n")
    cfb.write_stream("PROJECT", b"\r\n".join(line for line in lines if not line.startswith(b"Package=")))
    parts[entry] = cfb.to_bytes()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for part, data in parts.items():
            package.writestr(part, data)
    return path


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    quiet = host["quiet"]
    legacy = legacy_file(host, folder)
    opens = {"excel": f'Workbooks.Open("{legacy}")', "word": f'Documents.Open(FileName:="{legacy}")'}
    chains = [(host["add"], CHAIN), (host["add"], AT_ONCE), (host["add"], PACKAGE_LAST), (opens[name], LEGACY)]
    files: list[tuple[str, Path, bool]] = []
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            for start, chain in chains:
                opened = ""
                for label, actions in chain:
                    target = f"Set d = {start}" if not opened else f"Set d = {host['item'].format(name=opened)}"
                    opened = run(office, "\n".join([quiet, target, *map(vba, actions), "out = d.Name"]))
                    for path, macro, save_as in saves(host, folder, label):
                        opened = run(office, "\n".join([quiet, f"Set d = {host['item'].format(name=opened)}",
                                                        f"d.{save_as}", "out = d.Name"]))
                        files.append((label + (".macro" if macro else ""), path, macro))
                run(office, f"{quiet}\nSet d = {host['item'].format(name=opened)}\nd.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            lines = [quiet]
            for start, chain in chains:
                lines.append(f"Set d = {start}")
                for label, actions in chain:
                    lines += map(vba, actions)
                    for path, macro, save_as in saves(host, folder, label):
                        lines.append(f"d.{save_as}")
                        files.append((label + (".macro" if macro else ""), path, macro))
                lines.append(f"d.{host['close']}")
            run(office, "\n".join(lines))
    return {"project_storage": host["storage"],
            "states": {label: state(path, host, macro) for label, path, macro in files}}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, host_record in record.items():
        for label, saved in host_record["states"].items():
            declarations = saved["project"][1:next(i for i, line in enumerate(saved["project"])
                                                   if line.startswith("Name="))]
            print(f"{name} {label}: {declarations} | storages {saved['storages']} | {saved['references']}")


if __name__ == "__main__":
    main()
