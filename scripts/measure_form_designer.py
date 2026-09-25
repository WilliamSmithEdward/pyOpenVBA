"""What Excel's and Word's form designers write when they add controls.

Each application adds UserForms through its editor and fills them through
the designer (Designer.Controls.Add, Pages.Add, font and size settings):

    Plain      one control of every kind, on a form with the default font
    Fonted     a form set to Arial 10 with controls on it, a Frame set to
               Courier New 9 with controls in it, a MultiPage with a
               control on its first page and a page added, and a second
               MultiPage resized to 300 x 200 points
    Neighbour  a CheckBox given its own font, then a ToggleButton beside it
    Frames     an empty Frame, and a Frame with one control in it
    Effects    a form set bold and struck through, with controls on it and
               a Frame set italic and underlined with a control in it
    Charset    a form set to Arial with the Greek character set
    Tabs       Images and other controls added in turn
    Nested     a Frame with a control in it on a MultiPage's first page,
               two pages added, and a control on the second page

and saves the file macro-enabled. The record keeps, for every storage of
every form, its streams byte for byte, its class id and its directory
entry's timestamps, and the form's \\x03VBFrame text.

    python scripts/measure_form_designer.py

writes tests/fixtures/form_designer.json, which tests/test_form_designer.py
replays. No file is kept: a project with a form names the Temp folder of
whoever saved it.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, WordSession

from pyopenvba.cfb import CFB

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "form_designer.json"
#: Each application: its session, its file type with format number and project part, and its VBA.
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "suffix": "xlsm", "format": 52, "part": "xl/vbaProject.bin",
              "quiet": "Application.DisplayAlerts = False", "add": "Workbooks.Add(-4167)",
              "item": 'Workbooks("{name}")', "save_as": 'SaveAs Filename:="{path}", FileFormat:={format}',
              "close": "Close False"},
    "word": {"session": WordSession, "suffix": "docm", "format": 13, "part": "word/vbaProject.bin",
             "quiet": "Application.DisplayAlerts = 0", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:={format}', "close": "Close SaveChanges:=0"},
}
#: Every control kind, by the name the designer would give the first one.
KINDS = ["Label", "CommandButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
         "Frame", "MultiPage", "Image", "SpinButton", "ScrollBar", "TabStrip"]


def _form(name: str) -> list[str]:
    return ["Set c = d.VBProject.VBComponents.Add(3)", f'c.Name = "{name}"', "Set f = c.Designer"]


def _add(kind: str, name: str, into: str = "f") -> str:
    return f'{into}.Controls.Add "Forms.{kind}.1", "{name}"'


FORMS = [
    *_form("Plain"), *[_add(kind, f"{kind}1") for kind in KINDS],
    *_form("Fonted"), 'f.Font.Name = "Arial"', "f.Font.Size = 10",
    _add("Label", "Label1"), _add("CommandButton", "CommandButton1"), _add("CheckBox", "CheckBox1"),
    _add("ToggleButton", "ToggleButton1"), _add("TextBox", "TextBox1"),
    'Set g = f.Controls.Add("Forms.Frame.1", "Frame1")', 'g.Font.Name = "Courier New"', "g.Font.Size = 9",
    _add("Label", "Label2", "g"), _add("ToggleButton", "ToggleButton2", "g"),
    'Set m = f.Controls.Add("Forms.MultiPage.1", "MultiPage1")', _add("Label", "Label3", "m.Pages(0)"),
    'm.Pages.Add "Page3"',
    'Set m = f.Controls.Add("Forms.MultiPage.1", "MultiPage2")', "m.Width = 300", "m.Height = 200",
    *_form("Neighbour"), 'Set k = f.Controls.Add("Forms.CheckBox.1", "CheckBox1")', 'k.Font.Name = "Courier New"',
    _add("ToggleButton", "ToggleButton1"),
    *_form("Frames"), _add("Frame", "Frame1"), 'Set g = f.Controls.Add("Forms.Frame.1", "Frame2")',
    _add("Label", "Label1", "g"),
    *_form("Effects"), "f.Font.Bold = True", "f.Font.Strikethrough = True",
    _add("Label", "Label1"), _add("CommandButton", "CommandButton1"), _add("CheckBox", "CheckBox1"),
    'Set g = f.Controls.Add("Forms.Frame.1", "Frame1")', "g.Font.Italic = True", "g.Font.Underline = True",
    _add("Label", "Label2", "g"),
    *_form("Charset"), 'f.Font.Name = "Arial"', "f.Font.Charset = 161", _add("Label", "Label1"),
    _add("TextBox", "TextBox1"),
    *_form("Tabs"), _add("Image", "Image1"), _add("Label", "Label1"), _add("Image", "Image2"),
    _add("CommandButton", "CommandButton1"), _add("TextBox", "TextBox1"),
    *_form("Nested"), 'Set m = f.Controls.Add("Forms.MultiPage.1", "MultiPage1")',
    'Set g = m.Pages(0).Controls.Add("Forms.Frame.1", "Frame1")', _add("Label", "Label1", "g"),
    'm.Pages.Add "Page3"', 'm.Pages.Add "Page4"', _add("CommandButton", "CommandButton1", "m.Pages(1)"),
]


def storages(cfb: CFB, path: tuple[str, ...]) -> dict[str, Any]:
    """Every storage from ``path`` down: its streams, class id and entry fields, keyed by path."""
    index = cfb._resolve_path(path)  # noqa: SLF001 - a measurement reads the entry as the file has it
    raw = cfb._directory[index].raw  # noqa: SLF001
    out = {"/".join(path): {
        "clsid": raw[80:96].hex(), "state": raw[96:100].hex(), "created": raw[100:108].hex(),
        "modified": raw[108:116].hex(),
        "streams": {name: cfb.get_stream_at(path, name).hex() for name in sorted(cfb.list_streams_at(path))},
    }}
    for child in sorted(cfb.list_storages_at(path)):
        out.update(storages(cfb, (*path, child)))
    return out


def run(office: Any, body: str) -> str:
    done = office.run_vba("Public Function Step() As String\nDim d As Object, c As Object, f As Object, "
                          f"g As Object, m As Object, k As Object, out As String\n{body}\nStep = out\nEnd Function\n",
                          "Step", timeout=300.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    target = folder / f"designer.{host['suffix']}"
    save = host["save_as"].format(path=target, format=host["format"])
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            made = run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *FORMS, "out = d.Name"]))
            item = host["item"].format(name=made)
            run(office, f"{host['quiet']}\nSet d = {item}\nd.{save}\nd.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *FORMS, f"d.{save}",
                                   f"d.{host['close']}"]))
    with zipfile.ZipFile(target) as package:
        cfb = CFB.from_bytes(package.read(host["part"]))
    return {form: storages(cfb, (form,))
            for form in ("Plain", "Fonted", "Neighbour", "Frames", "Effects", "Charset", "Tabs", "Nested")}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("wrote", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
