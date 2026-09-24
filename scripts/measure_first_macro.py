"""What Excel, Word and PowerPoint write when a file gets its first macro.

Each application saves a file with no VBA in it -- Excel's with renamed,
reordered worksheets and a chart sheet -- then saves it three more times:
opened and saved untouched; after a standard module is added; and after
a module is added and removed again, which leaves the project the module
made and nothing in it. The record names, for each save, the parts that
differ from the file saved with no VBA, and reads back the project each
application made: its components with their types, and in Excel each
sheet's code name.

    python scripts/measure_first_macro.py

writes tests/fixtures/first_macro/: for each application the file with no
VBA (before), the file with the module (module) and the file with the
emptied project (emptied), and first_macro.json, which
tests/test_add_vba_project.py replays.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, PowerPointSession, WordSession

from fixture_workbook import copy_saved

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "first_macro"
MODULE = '"Public Sub Hello()" & vbCrLf & "End Sub"'
#: Each application: its session, the file's extension, how its VBA makes the file, saves, opens and closes one
#: with no alert in the way, and what it reads back about the file's own objects.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {
        "session": ExcelSession, "suffix": "xlsm", "quiet": "Application.DisplayAlerts = False",
        "make": ("Set d = Workbooks.Add(-4167)\nd.Worksheets(1).Name = \"Data\"\n"
                 "d.Worksheets.Add(After:=d.Worksheets(1)).Name = \"Summary\"\n"
                 "d.Worksheets.Add(After:=d.Worksheets(2)).Name = \"Notes\"\n"
                 "d.Worksheets(\"Notes\").Move Before:=d.Worksheets(\"Data\")\n"
                 "d.Charts.Add After:=d.Sheets(d.Sheets.Count)\n"
                 "d.Worksheets(\"Data\").Range(\"A1\").Value = 1\nd.Worksheets(\"Data\").Activate"),
        "save_as": 'SaveAs Filename:="{path}", FileFormat:=52', "open": 'Workbooks.Open("{path}")',
        "item": 'Workbooks("{name}")', "close": "Close False",
        "names": ('Dim s As Object\nFor Each s In d.Sheets\nout = out & s.Name & "=" & s.CodeName & ";"\nNext\n'
                  'out = out & "|" & d.CodeName'),
    },
    "word": {
        "session": WordSession, "suffix": "docm", "quiet": "Application.DisplayAlerts = 0",
        "make": 'Set d = Documents.Add\nd.Content.Text = "A document"',
        "save_as": 'SaveAs2 FileName:="{path}", FileFormat:=13', "open": 'Documents.Open("{path}")',
        "item": 'Documents("{name}")', "close": "Close SaveChanges:=0", "names": 'out = out & "|"',
    },
    "powerpoint": {
        "session": PowerPointSession, "suffix": "pptm", "quiet": "Application.DisplayAlerts = 1",
        "make": "Set d = Presentations.Add(0)\nd.Slides.Add 1, 12",
        "save_as": 'SaveAs "{path}", 25', "open": 'Presentations.Open("{path}", 0, 0, 0)',
        "item": 'Presentations("{name}")', "close": "Close", "names": 'out = out & "|"',
    },
}


def components() -> str:
    return ('Dim c As Object\nFor Each c In d.VBProject.VBComponents\nout = out & c.Name & ":" & c.Type & ";"\n'
            "Next\n")


def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def differing(one: Path, other: Path) -> list[str]:
    """The parts one file has and the other lacks, or has with other bytes."""
    first, second = parts(one), parts(other)
    return sorted(name for name in first.keys() | second.keys() if first.get(name) != second.get(name))


def measure(name: str, host: dict[str, Any]) -> dict[str, Any]:
    folder = Path(tempfile.mkdtemp())
    suffix = host["suffix"]
    before, resaved = folder / f"before.{suffix}", folder / f"resaved.{suffix}"
    module, emptied = folder / f"module.{suffix}", folder / f"emptied.{suffix}"

    def opening(path: Path) -> str:
        return f"Set d = {host['open'].format(path=path)}"

    def saving(path: Path) -> str:
        return f"d.{host['save_as'].format(path=path)}\nd.{host['close']}"

    open_before = f"Set d = {host['item'].format(name=before.name)}"
    adding = [f"{opening(before)}\nd.VBProject.VBComponents.Add(1).CodeModule.AddFromString {MODULE}",
              f"{open_before}\n{saving(module)}"]
    emptying = [f"{opening(before)}\nd.VBProject.VBComponents.Add 1",
                f'{open_before}\nd.VBProject.VBComponents.Remove d.VBProject.VBComponents("Module1")',
                f"{open_before}\n{saving(emptied)}"]
    if name != "excel":
        # Excel stops at a VBA dialog when one call makes a project and saves it; Word and PowerPoint lose the
        # harness's own procedures when a call ends after another file's project changed. So one each way.
        adding, emptying = ["\n".join(adding)], ["\n".join(emptying)]
    steps = [
        f"{host['make']}\n{saving(before)}",
        f"{opening(before)}\n{saving(resaved)}",
        *adding,
        *emptying,
        f"{opening(module)}\n{components()}out = out & \"#\"\n{host['names']}\nd.{host['close']}",
        f"{opening(emptied)}\n{components()}out = out & \"#\"\n{host['names']}\nd.{host['close']}",
    ]
    answers: list[str] = []
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        for step in steps:
            result = office.run_vba(f"Public Function Step() As String\nDim d As Object, out As String\n"
                                    f"{host['quiet']}\n{step}\nStep = out\nEnd Function\n", "Step", timeout=120.0)
            assert result.ok, f"{name}: {result.outcome} {result.message}\n{step}"
            answers.append(str(result.value or ""))
    saves = {"before": before, "module": module, "emptied": emptied}
    for label, path in saves.items():
        copy_saved(path, OUT / f"{name}_{label}.{suffix}")
    module_read, emptied_read = answers[-2], answers[-1]
    return {
        "suffix": suffix,
        "resave_changes": differing(before, resaved),
        "module_changes": differing(resaved, module),
        "emptied_changes": differing(resaved, emptied),
        "module": _read(module_read),
        "emptied": _read(emptied_read),
    }


def _read(text: str) -> dict[str, Any]:
    listed, _, names = text.partition("#")
    sheets, _, workbook = names.partition("|")
    return {"components": [one.split(":") for one in listed.split(";") if one],
            "code_names": [one.split("=") for one in sheets.split(";") if one], "workbook_code_name": workbook}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    record = {name: measure(name, host) for name, host in HOSTS.items()}
    (OUT / "first_macro.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, found in record.items():
        print(name, json.dumps(found, indent=1))


if __name__ == "__main__":
    main()
