"""How a form's font shapes what the designer adds to it: a MultiPage's tab band and the default sizes.

Each application makes one UserForm per font, sets the form's font through
the designer, and adds a MultiPage and one control of each other text kind
at their default sizes. The record keeps, per form, the StdFont the form
stores, each control's size, the MultiPage's size, its TabStrip's size and
each page's position and size.

    python scripts/measure_form_fonts.py

writes tests/fixtures/form_fonts.json, which tests/test_form_designer.py
replays.
"""

from __future__ import annotations

import json
import struct
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pyvbaharness import ExcelSession, HarnessConfig, WordSession

from pyopenvba._oforms_records import SPECS_BY_CACHE_INDEX, parse_record
from pyopenvba.cfb import CFB
from pyopenvba.forms import _read_container  # noqa: PLC2701 - the measurement reads the streams as they are

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "form_fonts.json"
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "suffix": "xlsm", "format": 52, "part": "xl/vbaProject.bin",
              "quiet": "Application.DisplayAlerts = False", "add": "Workbooks.Add(-4167)",
              "item": 'Workbooks("{name}")', "save_as": 'SaveAs Filename:="{path}", FileFormat:={format}',
              "close": "Close False"},
    "word": {"session": WordSession, "suffix": "docm", "format": 13, "part": "word/vbaProject.bin",
             "quiet": "Application.DisplayAlerts = 0", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:={format}', "close": "Close SaveChanges:=0"},
}
#: The fonts, each at a size the designer stores as it is (a multiple of 0.375 pt).
FONTS = [("Tahoma", 7.875), ("Tahoma", 8.25), ("Tahoma", 9), ("Tahoma", 9.75), ("Tahoma", 10.125), ("Tahoma", 12),
         ("Arial", 8.25), ("Arial", 9), ("Arial", 9.75), ("Arial", 10.125), ("Arial", 12),
         ("Segoe UI", 8.25), ("Segoe UI", 9), ("Segoe UI", 9.75), ("Calibri", 9.75), ("Calibri", 11.25),
         ("Aptos", 10.125), ("Aptos", 11.25), ("Microsoft Sans Serif", 8.25), ("Verdana", 8.25),
         ("Times New Roman", 12), ("Courier New", 9)]
KINDS = ["Label", "CommandButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton"]


def forms_vba() -> list[str]:
    lines: list[str] = []
    for number, (face, size) in enumerate(FONTS, 1):
        lines += ["Set c = d.VBProject.VBComponents.Add(3)", f'c.Name = "F{number:02d}"', "Set f = c.Designer",
                  f'f.Font.Name = "{face}"', f"f.Font.Size = {size}", 'f.Controls.Add "Forms.MultiPage.1", "MultiPage1"',
                  *[f'f.Controls.Add "Forms.{kind}.1", "{kind}1"' for kind in KINDS]]
    return lines


def stdfont(blob: bytes) -> dict[str, Any]:
    if not blob:
        return {}
    _, charset, flags, weight, height, length = struct.unpack_from("<BHBHIB", blob, 16)
    return {"name": blob[27:27 + length].decode("latin-1"), "cy_size": height, "charset": charset,
            "flags": flags, "weight": weight}


def take_apart(cfb: CFB, form: str) -> dict[str, Any]:
    stream, _ = _read_container(cfb.get_stream_at((form,), "f"), "cp1252")
    o = cfb.get_stream_at((form,), "o")
    sizes: dict[str, list[int]] = {}
    at = 0
    multipage = 0
    for site in stream.sites:
        size = site.values.get("ObjectStreamSize", 0)
        name = site.strings["Name"].text
        spec = SPECS_BY_CACHE_INDEX.get(site.values["ClsidCacheIndex"])
        if spec is not None and size:
            record = parse_record(o[at:at + size], spec, "cp1252")
            sizes[name] = [record.sizes["Size"].width, record.sizes["Size"].height]
        if name == "MultiPage1":
            multipage = site.values["ID"]
        at += size
    storage = (form, f"i{multipage:02d}")
    inner, _ = _read_container(cfb.get_stream_at(storage, "f"), "cp1252")
    inner_o = cfb.get_stream_at(storage, "o")
    tabstrip = inner.sites[0]
    tab_record = parse_record(inner_o[:tabstrip.values["ObjectStreamSize"]], SPECS_BY_CACHE_INDEX[18], "cp1252")
    pages = []
    for site in inner.sites[1:]:
        page, _ = _read_container(cfb.get_stream_at((*storage, f"i{site.values['ID']:02d}"), "f"), "cp1252")
        size = page.record.sizes["DisplayedSize"]
        pages.append({"position": list(site.position), "size": [size.width, size.height]})
    shown = inner.record.sizes["DisplayedSize"]
    return {"font": stdfont(getattr(stream, "font_raw", b"")), "sizes": sizes,
            "multipage": {"size": [shown.width, shown.height],
                          "tabstrip": [tab_record.sizes["Size"].width, tab_record.sizes["Size"].height],
                          "tab_font": tab_record.text_props.properties() if tab_record.text_props else {},
                          "pages": pages}}


def run(office: Any, body: str) -> str:
    done = office.run_vba("Public Function Step() As String\nDim d As Object, c As Object, f As Object, "
                          f"out As String\n{body}\nStep = out\nEnd Function\n", "Step", timeout=600.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    target = folder / f"fonts.{host['suffix']}"
    save = host["save_as"].format(path=target, format=host["format"])
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            made = run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *forms_vba(), "out = d.Name"]))
            run(office, f"{host['quiet']}\nSet d = {host['item'].format(name=made)}\nd.{save}\nd.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *forms_vba(), f"d.{save}",
                                   f"d.{host['close']}"]))
    with zipfile.ZipFile(target) as package:
        cfb = CFB.from_bytes(package.read(host["part"]))
    return {f"{face} {size}": take_apart(cfb, f"F{number:02d}") for number, (face, size) in enumerate(FONTS, 1)}


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("wrote", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
