"""Where Excel's and Word's form designers put a MultiPage's pages when it changes.

A MultiPage from Designer.Controls.Add has the default size, so VBA gets
another size by setting Width and Height afterwards. Each application adds
MultiPages through its designer and changes them, on these forms:

    Sizes      resized to twelve sizes, on a form with the default font
    Arial      six of those sizes on a form set to Arial 10
    Tahoma9    the same on a form set to Tahoma 9
    Courier    the same on a form set to Courier New 9
    Tahoma12   the same on a form set to Tahoma 12
    Order      resized with another page selected, resized twice, and
               resized in one direction only
    Added      a page added after a resize, before one, and between two
    Refont     the MultiPage's font changed before and after a resize and
               before a page is added

Then the library makes a MultiPage whose pages sit where a designer at 192
DPI puts them (top 542, as in #31's sample), and each application opens it,
adds a page to it through its designer, edits the form beside a second such
MultiPage, and saves.

The record keeps every storage of every form, as
scripts/measure_form_designer.py does.

    python scripts/measure_multipage_resize.py

writes tests/fixtures/multipage_resize.json, which
tests/test_multipage_resize.py replays. No file is kept.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from measure_form_designer import HOSTS, run, storages
from pyvbaharness import HarnessConfig

from pyopenvba import ExcelFile, WordFile
from pyopenvba.cfb import CFB

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "fixtures" / "multipage_resize.json"

#: Width and height in points, the way VBA sets them.
SIZES = [(300, 200), (150, 100), (200.5, 133.3), (123.45, 98.76), (400, 300), (96, 72), (250, 180.25),
         (333.33, 222.22), (180, 120), (512, 384), (100, 60), (144.7, 108.3)]
SOME = [SIZES[0], SIZES[1], SIZES[3], SIZES[4], SIZES[7], SIZES[10]]
#: The page a planted MultiPage's pages sit under, as a designer at 192 DPI puts them for Tahoma 7.875.
PLANTED_TOP = 542


def _form(name: str, *font: str) -> list[str]:
    return ["Set c = d.VBProject.VBComponents.Add(3)", f'c.Name = "{name}"', "Set f = c.Designer",
            *[f"f.Font.{setting}" for setting in font]]


def _multipage(name: str, *steps: str) -> list[str]:
    return [f'Set m = f.Controls.Add("Forms.MultiPage.1", "{name}")', *steps]


def _resize(width: float, height: float) -> list[str]:
    return [f"m.Width = {width}", f"m.Height = {height}"]


def _sized(name: str, font: tuple[str, ...], sizes: list[tuple[float, float]]) -> list[str]:
    return [*_form(name, *font),
            *[line for index, size in enumerate(sizes, 1) for line in _multipage(f"MP{index:02d}", *_resize(*size))]]


FORMS: dict[str, list[str]] = {
    "Sizes": _sized("Sizes", (), SIZES),
    "Arial": _sized("Arial", ('Name = "Arial"', "Size = 10"), SOME),
    "Tahoma9": _sized("Tahoma9", ("Size = 9",), SOME),
    "Courier": _sized("Courier", ('Name = "Courier New"', "Size = 9"), SOME),
    "Tahoma12": _sized("Tahoma12", ("Size = 12",), SOME),
    "Order": [
        *_form("Order"),
        *_multipage("MP01", "m.Value = 1", *_resize(300, 200)),
        *_multipage("MP02", 'm.Pages.Add "Page3"', "m.Value = 2", *_resize(300, 200)),
        *_multipage("MP03", *_resize(300, 200), *_resize(180, 120)),
        *_multipage("MP04", "m.Width = 250"),
        *_multipage("MP05", "m.Height = 250"),
        *_multipage("MP06", "m.Value = 1", *_resize(300, 200), "m.Value = 0", *_resize(180, 120)),
    ],
    "Added": [
        *_form("Added"),
        *_multipage("MP01", *_resize(300, 200), 'm.Pages.Add "Page3"'),
        *_multipage("MP02", *_resize(300, 200), 'm.Pages.Add "Page3"', *_resize(180, 120)),
        *_multipage("MP03", 'm.Pages.Add "Page3"', *_resize(300, 200)),
        *_multipage("MP04", *_resize(100, 60), 'm.Pages.Add "Page3"'),
    ],
    "Refont": [
        *_form("Refont"),
        *_multipage("MP01", "m.Font.Size = 14"),
        *_multipage("MP02", *_resize(300, 200), "m.Font.Size = 14"),
        *_multipage("MP03", "m.Font.Size = 14", 'm.Pages.Add "Page3"'),
        *_multipage("MP04", "m.Font.Size = 14", *_resize(300, 200)),
    ],
}


def planted(path: Path, kind: type[Any]) -> None:
    """A file whose form holds two MultiPages with their pages where a designer at 192 DPI puts them."""
    with kind.create_new(path) as document:
        form = document.add_form("Planted")
        for name in ("MPA", "MPB"):
            form.add_control("MultiPage", name)
        levels = form._levels  # a measurement plants a layout no public call makes
        for level in levels:
            for site in level.sites:
                if site.clsid_cache_index != 7 or site.position is None:
                    continue
                left, top = site.position
                page = next(child for child in levels if child.path[-1] == f"i{site.id:02d}")
                size = page.record.sizes["DisplayedSize"]
                page.record.set_size(size.width, size.height - (PLANTED_TOP - top))
                site.position = (left, PLANTED_TOP)
        document.save()


def measure(name: str, host: dict[str, Any], folder: Path) -> dict[str, Any]:
    target = folder / f"resize.{host['suffix']}"
    save = host["save_as"].format(path=target, format=host["format"])
    body = [line for lines in FORMS.values() for line in lines]
    plant = folder / f"planted.{host['suffix']}"
    planted(plant, ExcelFile if name == "excel" else WordFile)
    opened = "Workbooks.Open(\"{0}\")" if name == "excel" else "Documents.Open(FileName:=\"{0}\")"
    edit = ['Set f = d.VBProject.VBComponents("Planted").Designer', 'f.Controls("MPA").Pages.Add "Page3"',
            'f.Controls.Add "Forms.Label.1", "Touch"']
    with host["session"](HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if name == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            made = run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *body, "out = d.Name"]))
            item = host["item"].format(name=made)
            run(office, f"{host['quiet']}\nSet d = {item}\nd.{save}\nd.{host['close']}")
            run(office, "\n".join([host["quiet"], f"Set d = {opened.format(plant)}", *edit, "out = d.Name"]))
            run(office, f"{host['quiet']}\nSet d = {host['item'].format(name=plant.name)}\nd.Save\n"
                        f"d.{host['close']}")
        else:
            # Word loses the harness's procedures once another file's project changed: one call each.
            run(office, "\n".join([host["quiet"], f"Set d = {host['add']}", *body, f"d.{save}",
                                   f"d.{host['close']}"]))
            run(office, "\n".join([host["quiet"], f"Set d = {opened.format(plant)}", *edit, "d.Save",
                                   f"d.{host['close']}"]))
    record: dict[str, Any] = {}
    for path, forms in ((target, list(FORMS)), (plant, ["Planted"])):
        with zipfile.ZipFile(path) as package:
            cfb = CFB.from_bytes(package.read(host["part"]))
        record.update({form: storages(cfb, (form,)) for form in forms})
    return record


def main() -> None:
    folder = Path(tempfile.mkdtemp())
    record = {name: measure(name, host, folder) for name, host in HOSTS.items()}
    OUT.write_text(json.dumps(record, indent=1) + "\n", encoding="utf-8")
    print("wrote", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main()
