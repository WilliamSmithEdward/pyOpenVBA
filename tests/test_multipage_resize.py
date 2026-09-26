"""Where Excel's and Word's form designers lay a MultiPage's pages out when it changes, and the library beside them.

tests/fixtures/multipage_resize.json is what scripts/measure_multipage_resize.py
saw each designer write: MultiPages that VBA sized after Designer.Controls.Add
added them, at twelve sizes and under five fonts; sized with another page
shown, and sized twice; pages added after a resize; fonts set on the
MultiPage; and, in a file the library made, pages laid out as a designer at
192 DPI lays them out, which a designer at 96 DPI opened and added a page to.

The designer lays out only the page it shows, in its TabStrip. Every other
page keeps the layout it was given when it was added, which is always the
default size's.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile
from pyopenvba._oforms_records import Size
from pyopenvba.cfb import CFB
from pyopenvba.forms import FormControl, VBAForm

RECORD: dict[str, dict[str, dict[str, Any]]] = json.loads(
    (Path(__file__).parent / "fixtures" / "multipage_resize.json").read_text(encoding="utf-8"))
HOSTS = list(RECORD)
PART = "xl/vbaProject.bin"
#: The forms whose MultiPages were only sized, each under its own font.
SIZED = ["Sizes", "Arial", "Tahoma9", "Courier", "Tahoma12"]


def _office_form(host: str, form: str, tmp_path: Path) -> Path:
    """A workbook whose form ``form`` is the one the application's designer wrote, storage for storage."""
    target = tmp_path / f"{form}.xlsm"
    with ExcelFile.create_new(target) as workbook:
        workbook.add_form(form)
        workbook.save()
    with zipfile.ZipFile(target) as package:
        parts = {info.filename: package.read(info) for info in package.infolist()}
    cfb = CFB.from_bytes(parts[PART])
    cfb.remove_storage_at((), form)
    for path, level in RECORD[host][form].items():
        *parent, name = path.split("/")
        cfb.add_substorage_at(parent, name, bytes.fromhex(level["clsid"]))
        for stream, data in level["streams"].items():
            cfb.add_stream_at([*parent, name], stream, bytes.fromhex(data))
    parts[PART] = cfb.to_bytes()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for name, data in parts.items():
            package.writestr(name, data)
    return target


def _form(workbook: ExcelFile, name: str) -> VBAForm:
    return next(form for form in workbook.forms() if form.name == name)


def _level(form: VBAForm, control_id: int) -> Any:
    """The storage level of the container with that id."""
    levels = form._levels  # pyright: ignore[reportPrivateUsage]
    return next(level for level in levels if level.path[-1] == f"i{control_id:02d}")


def _tabstrip(level: Any) -> FormControl:
    return next(control for control in level.controls if control.clsid_cache_index == 18)


def _layout(form: VBAForm, multipage: str) -> list[tuple[object, object]]:
    """A MultiPage's TabStrip size, then where each of its pages is sited and its size. Every MultiPage has
    a Page1, so the pages are found through the MultiPage rather than by name."""
    level = _level(form, form.control(multipage).id)
    return [("TabStrip", _tabstrip(level).get("Size")),
            *[(site.position, _level(form, site.id).record.sizes.get("DisplayedSize"))
              for site in level.sites if site.clsid_cache_index == 7]]


def _multipages(form: VBAForm) -> list[str]:
    return [control.name for control in form.controls if control.clsid_cache_index == 57]


def _size(control: FormControl) -> tuple[float, float]:
    """A control's size in points, which add_control takes."""
    size = control.get("DisplayedSize")
    assert isinstance(size, Size)
    return size.width * 72 / 2540, size.height * 72 / 2540


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", SIZED)
def test_a_multipage_made_at_a_size_is_the_one_the_designer_sizes(host: str, form: str, tmp_path: Path) -> None:
    # Each MultiPage was added at the default size and sized after; the TabStrip and the first page, the
    # one shown, followed it, and the page's top moved a unit or two with the size.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        for index, name in enumerate(_multipages(office)):
            width, height = _size(office.control(name))
            office.add_control("MultiPage", f"New{index}", left=0, top=0, width=width, height=height)
            assert _layout(office, f"New{index}") == _layout(office, name), name


@pytest.mark.parametrize("host", HOSTS)
def test_a_page_goes_under_tabs_in_the_font_set_on_its_multipage(host: str, tmp_path: Path) -> None:
    # MP01 was set to Tahoma 14 pt, which its TabStrip keeps; MP03 was set the same, and the designer then
    # added its Page3.
    with ExcelFile(_office_form(host, "Refont", tmp_path)) as workbook:
        office = _form(workbook, "Refont")
        office.add_page("MP01", name="Page3")
        assert _layout(office, "MP01")[-1] == _layout(office, "MP03")[-1]


ROWS = [(host, form) for host in HOSTS for form in RECORD[host]]


@pytest.mark.parametrize(("host", "form"), ROWS, ids=[f"{host}-{form}" for host, form in ROWS])
def test_the_page_shown_is_laid_out_in_the_tabstrip(host: str, form: str, tmp_path: Path) -> None:
    # Whatever changed -- the size, the page shown, the font -- the page the TabStrip's ListIndex names sits
    # where the library lays a page out in a TabStrip of that size. Word leaves the TabStrip at the default
    # size once a font is set on the MultiPage (the Refont form), and the page follows the TabStrip.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        for name in _multipages(office):
            level = _level(office, office.control(name).id)
            tabstrip = _tabstrip(level)
            shown, size = tabstrip.get("ListIndex") or 0, tabstrip.get("Size")
            assert isinstance(shown, int) and isinstance(size, Size)
            left, top, width, height = office._page_box(level, size)  # pyright: ignore[reportPrivateUsage]
            assert _layout(office, name)[1 + shown] == ((left, top), Size(width, height)), name