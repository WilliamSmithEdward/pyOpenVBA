"""What Excel, Word and PowerPoint do with a VBA signature kept beside vbaProject.bin.

Each application makes a file with one module and saves it. The three
signature parts of EUROTOOL.XLAM, the add-in Office installs with Excel,
then go beside its VBA project with their relationships and Overrides,
as Office lays them out. The application opens the file and reads
VBASigned, saves it as it is, changes one line of the module, saves it
again, and reads VBASigned from both saves.

    python scripts/measure_signature_parts.py

writes tests/fixtures/signature_parts/: signature_parts.json with the
readings and, from both saves, [Content_Types].xml and the project's
relationships part; and the untouched save of each file as excel.xlsm,
word.docm and powerpoint.pptm, with placeholder bytes in the signature
parts, which are Microsoft's. The application reads VBASigned from that
file too. tests/test_signature_parts.py replays them.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, PowerPointSession, WordSession

from fixture_workbook import without_save_path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "signature_parts"
EUROTOOL = Path(r"C:\Program Files\Microsoft Office\root\Office16\Library\EUROTOOL.XLAM")
PARTS = ("vbaProjectSignature.bin", "vbaProjectSignatureAgile.bin", "vbaProjectSignatureV3.bin")
TYPES = {"vbaProjectSignature.bin": "application/vnd.ms-office.vbaProjectSignature",
         "vbaProjectSignatureAgile.bin": "application/vnd.ms-office.vbaProjectSignatureAgile",
         "vbaProjectSignatureV3.bin": "application/vnd.ms-office.vbaProjectSignatureV3"}
PLACEHOLDER = b"A placeholder for a signature part, which is Microsoft's.\r\n"

#: Each host: its session, the file's extension, the folder of its VBA project, and how its VBA makes, saves,
#: opens and closes a file with no alert in the way.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "suffix": "xlsm", "folder": "xl", "add": "Workbooks.Add(1)",
              "save_as": 'SaveAs Filename:="{path}", FileFormat:=52', "open": 'Workbooks.Open("{path}")',
              "close": "Close False", "quiet": "Application.DisplayAlerts = False"},
    "word": {"session": WordSession, "suffix": "docm", "folder": "word", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:=13', "open": 'Documents.Open("{path}")',
             "close": "Close SaveChanges:=0", "quiet": "Application.DisplayAlerts = 0"},
    "powerpoint": {"session": PowerPointSession, "suffix": "pptm", "folder": "ppt", "add": "Presentations.Add(0)",
                   "save_as": 'SaveAs "{path}", 25', "open": 'Presentations.Open("{path}", 0, 0, 0)',
                   "close": "Close", "quiet": "Application.DisplayAlerts = 1"},
}


def signed(base: Path, target: Path, folder: str) -> None:
    """``base`` with EUROTOOL's signature parts beside its VBA project: related, overridden, as Office has them."""
    with zipfile.ZipFile(EUROTOOL) as source:
        parts = {name: source.read(f"xl/{name}") for name in PARTS}
        rels = source.read("xl/_rels/vbaProject.bin.rels").decode("utf-8")
    rels_name = f"{folder}/_rels/vbaProject.bin.rels"
    with zipfile.ZipFile(base) as original, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        if rels_name in original.namelist():
            # Word relates vbaData.xml from here already; the signature's relationships join it under new ids.
            kept = original.read(rels_name).decode("utf-8")
            taken = len(re.findall(r"<Relationship\b", kept))
            added = [re.sub(r'Id="rId\d+"', f'Id="rId{taken + index}"', one)
                     for index, one in enumerate(re.findall(r"<Relationship\b[^>]*/>", rels), 1)]
            rels = kept.replace("</Relationships>", "".join(added) + "</Relationships>")
        for info in original.infolist():
            if info.filename == rels_name:
                continue
            data = original.read(info.filename)
            if info.filename == "[Content_Types].xml":
                overrides = "".join(f'<Override PartName="/{folder}/{name}" ContentType="{TYPES[name]}"/>'
                                    for name in PARTS)
                data = data.decode("utf-8").replace("</Types>", overrides + "</Types>").encode("utf-8")
            out.writestr(info, data)
        for name, data in parts.items():
            out.writestr(f"{folder}/{name}", data)
        out.writestr(rels_name, rels.encode("utf-8"))


def with_placeholders(source: Path, target: Path) -> None:
    """``source`` with a placeholder in each signature part and everything else as the application saved it."""
    with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, "w") as out:
        for info in original.infolist():
            data = PLACEHOLDER if info.filename.rsplit("/", 1)[-1] in PARTS else original.read(info.filename)
            out.writestr(info, data)
    target.write_bytes(without_save_path(target.read_bytes()))


def texts(path: Path, folder: str) -> dict[str, str | None]:
    """The file's [Content_Types].xml and its project's relationships part, None where it has none."""
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        rels = f"{folder}/_rels/vbaProject.bin.rels"
        return {"content_types": package.read("[Content_Types].xml").decode("utf-8"),
                "relationships": package.read(rels).decode("utf-8") if rels in names else None,
                "parts": [name for name in names if name.rsplit("/", 1)[-1] in PARTS]}  # type: ignore[dict-item]


def measure(name: str, host: dict[str, Any]) -> dict[str, Any]:
    folder = Path(tempfile.mkdtemp())
    suffix, project = host["suffix"], host["folder"]
    base, signed_path = folder / f"base.{suffix}", folder / f"signed.{suffix}"
    same, edited = folder / f"same.{suffix}", folder / f"edited.{suffix}"
    fixture = OUT / f"{name}.{suffix}"

    def opening(path: Path) -> str:
        return f"Set d = {host['open'].format(path=path)}"

    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        made = office.run_vba(
            f"Public Function Make() As String\nDim d As Object\n{host['quiet']}\nSet d = {host['add']}\n"
            'd.VBProject.VBComponents.Add(1).CodeModule.AddFromString "Public Sub Hello()" & vbCrLf & "End Sub"\n'
            f"d.{host['save_as'].format(path=base)}\nd.{host['close']}\nMake = \"made\"\nEnd Function\n",
            "Make", timeout=120.0)
        assert made.ok, f"{name}: {made.outcome} {made.message}"
        signed(base, signed_path, project)
        probe = office.run_vba(
            f"Public Function Probe() As String\nDim d As Object, out As String\n{host['quiet']}\n"
            f"{opening(signed_path)}\nout = CStr(CBool(d.VBASigned))\n"
            f"d.{host['save_as'].format(path=same)}\nd.{host['close']}\n{opening(signed_path)}\n"
            "d.VBProject.VBComponents(\"Module1\").CodeModule.InsertLines 1, \"' edited\"\n"
            'out = out & "|" & CStr(CBool(d.VBASigned))\n'
            f"d.{host['save_as'].format(path=edited)}\nd.{host['close']}\n"
            f"{opening(edited)}\nout = out & \"|\" & CStr(CBool(d.VBASigned))\nd.{host['close']}\n"
            f"{opening(same)}\nout = out & \"|\" & CStr(CBool(d.VBASigned))\nd.{host['close']}\n"
            "Probe = out\nEnd Function\n", "Probe", timeout=180.0)
        assert probe.ok, f"{name}: {probe.outcome} {probe.message}"
        with_placeholders(same, fixture)
        # A placeholder is no signature: what the application makes of one is read, an error included.
        placeholder = office.run_vba(
            f"Public Function Placeholder() As String\nDim d As Object\n{host['quiet']}\nOn Error Resume Next\n"
            f"{opening(fixture)}\nIf Err.Number <> 0 Then\n"
            'Placeholder = "error " & Err.Number & ": " & Err.Description\nExit Function\nEnd If\n'
            f"Placeholder = CStr(CBool(d.VBASigned))\nd.{host['close']}\nEnd Function\n", "Placeholder",
            timeout=120.0)
        assert placeholder.ok, f"{name}: {placeholder.outcome} {placeholder.message}"
    opened, after_edit, edited_reopened, same_reopened = str(probe.value).split("|")
    return {"file": fixture.name, "project": project, "module": "Module1",
            "vba_signed": {"opened": opened, "edited": after_edit, "edited_save": edited_reopened,
                           "untouched_save": same_reopened, "placeholder": str(placeholder.value)},
            "untouched_save": texts(same, project), "edited_save": texts(edited, project)}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    record = {name: measure(name, host) for name, host in HOSTS.items()}
    (OUT / "signature_parts.json").write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    for name, found in record.items():
        print(name, found["vba_signed"], "edited save keeps", found["edited_save"]["parts"],
              "relationships", found["edited_save"]["relationships"] is not None)


if __name__ == "__main__":
    main()
