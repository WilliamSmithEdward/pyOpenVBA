"""Where Excel's and Word's designers store a property set through VBA, and what else setting it changes.

Each application adds one UserForm and sets a property per control through
the designer (Designer.Controls(...), the form's own Font and scroll bars):
the ones whose stored field is not named the way VBA names them, and the
ones whose setting changes a second field. The record keeps, per control,
its site's fields, its record's mask and its record's bytes; and the
form's own record and font, taken apart.

    python scripts/measure_form_properties.py

writes tests/fixtures/form_properties.json, which
tests/test_form_properties.py replays.
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
OUT = ROOT / "tests" / "fixtures" / "form_properties.json"
HOSTS: dict[str, dict[str, Any]] = {
    "excel": {"session": ExcelSession, "suffix": "xlsm", "format": 52, "part": "xl/vbaProject.bin",
              "quiet": "Application.DisplayAlerts = False", "add": "Workbooks.Add(-4167)",
              "item": 'Workbooks("{name}")', "save_as": 'SaveAs Filename:="{path}", FileFormat:={format}',
              "close": "Close False"},
    "word": {"session": WordSession, "suffix": "docm", "format": 13, "part": "word/vbaProject.bin",
             "quiet": "Application.DisplayAlerts = 0", "add": "Documents.Add",
             "save_as": 'SaveAs2 FileName:="{path}", FileFormat:={format}', "close": "Close SaveChanges:=0"},
}
#: Each control: its kind, its name, and the settings made on it, in order.
CONTROLS: list[tuple[str, str, list[str]]] = [
    ("CommandButton", "Plain", []),
    ("CommandButton", "NoFocus", ["TakeFocusOnClick = False"]),
    ("Label", "AlignLeft", ["TextAlign = 1"]),
    ("Label", "AlignCenter", ["TextAlign = 2"]),
    ("Label", "AlignRight", ["TextAlign = 3"]),
    ("CheckBox", "BoxRight", ["Alignment = 1"]),
    ("CheckBox", "BoxLeft", ["Alignment = 0"]),
    ("ComboBox", "Combo", ["Style = 0"]),
    ("ComboBox", "DropList", ["Style = 2"]),
    ("CheckBox", "Triple", ["TripleState = True"]),
    ("ListBox", "Columns", ["ColumnCount = 3", 'ColumnWidths = "20 pt;30 pt;40 pt"']),
    ("Label", "Disabled", ["Enabled = False"]),
    ("ListBox", "DisabledList", ["Enabled = False"]),
    ("Label", "Italic", ["Font.Italic = True"]),
    ("Label", "Underline", ["Font.Underline = True"]),
    ("Label", "Strike", ["Font.Strikethrough = True"]),
    ("Label", "Bold", ["Font.Bold = True"]),
    ("CheckBox", "Face", ['Font.Name = "Arial"']),
    ("TextBox", "BorderThenEffect", ["BorderStyle = 1", "SpecialEffect = 2"]),
    ("TextBox", "EffectThenBorder", ["SpecialEffect = 2", "BorderStyle = 1"]),
    ("Image", "ImageBorderThenEffect", ["BorderStyle = 1", "SpecialEffect = 2"]),
    ("Image", "ImageEffectThenBorder", ["SpecialEffect = 2", "BorderStyle = 1"]),
    ("ScrollBar", "BarDisabled", ["Enabled = False"]),
    ("SpinButton", "SpinDisabled", ["Enabled = False"]),
    ("ScrollBar", "BarMin", ["Min = 50"]),
]
FORM = ["f.ScrollBars = 3", "f.KeepScrollBarsVisible = 1", "f.Font.Italic = True", "f.Font.Underline = True"]


def body() -> list[str]:
    lines = ["Set c = d.VBProject.VBComponents.Add(3)", 'c.Name = "Probe"', "Set f = c.Designer", *FORM]
    for kind, name, settings in CONTROLS:
        lines.append(f'Set k = f.Controls.Add("Forms.{kind}.1", "{name}")')
        lines += [f"k.{setting}" for setting in settings]
    return lines


def stdfont(blob: bytes) -> dict[str, Any]:
    _, charset, flags, weight, height, length = struct.unpack_from("<BHBHIB", blob, 16)
    return {"name": blob[27:27 + length].decode("latin-1"), "cy_size": height, "charset": charset,
            "flags": flags, "weight": weight}


def take_apart(cfb: CFB) -> dict[str, Any]:
    stream, _ = _read_container(cfb.get_stream_at(("Probe",), "f"), "cp1252")
    o = cfb.get_stream_at(("Probe",), "o")
    controls: dict[str, Any] = {}
    at = 0
    for site in stream.sites:
        size = site.values.get("ObjectStreamSize", 0)
        piece = o[at:at + size]
        at += size
        record = parse_record(piece, SPECS_BY_CACHE_INDEX[site.values["ClsidCacheIndex"]], "cp1252")
        # The record's bytes are the measurement; the tests take them apart with the reader as it stands.
        controls[site.strings["Name"].text] = {
            "site": {name: value for name, value in site.values.items() if name != "NameData"},
            "mask": record.mask, "record": piece.hex()}
    form = stream.record
    return {"form": {"properties": {name: value if isinstance(value, (int, str)) else repr(value)
                                    for name, value in form.properties().items()},
                     "font": stdfont(stream.font_raw)},
            "controls": controls}


def run(office: Any, text: str) -> str:
    done = office.run_vba("Public Function Step() As String\nDim d As Object, c As Object, f As Object, "
                          f"k As Object, out As String\n{text}\nStep = out\nEnd Function\n", "Step", timeout=300.0)
    assert done.ok, f"{done.outcome} {done.message}"
    return str(done.value)


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    target = folder / f"properties.{host['suffix']}"
    save = host["save_as"].format(path=target, format=host["format"])
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            made = run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *body(), "out = d.Name"]))
            run(office, f"{host['quiet']}\nSet d = {host['item'].format(name=made)}\nd.{save}\nd.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *body(), f"d.{save}",
                                   f"d.{host['close']}"]))
    with zipfile.ZipFile(target) as package:
        return take_apart(CFB.from_bytes(package.read(host["part"])))


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("wrote", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
