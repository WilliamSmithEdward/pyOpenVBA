"""A control the library adds is the one Excel's and Word's form designers add.

tests/fixtures/form_designer.json is what scripts/measure_form_designer.py
saw when each application's designer filled forms through
Designer.Controls.Add: every storage of every form, byte for byte. The two
applications wrote the same bytes throughout. These tests rebuild a form
from those storages and have the library add the same controls to it, then
compare what it writes with what the designer wrote in the same place.
"""

from __future__ import annotations

import json
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile
from pyopenvba._oforms_pages import parse_string_array, serialize_string_array
from pyopenvba._oforms_records import ParsedRecord, Size, serialize_record
from pyopenvba.cfb import CFB
from pyopenvba.forms import FormControl, VBAForm

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, dict[str, dict[str, Any]]] = json.loads(
    (FIXTURES / "form_designer.json").read_text(encoding="utf-8"))
HOSTS = list(RECORD)
PART = "xl/vbaProject.bin"


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


def _record(control: FormControl) -> ParsedRecord:
    assert control.record is not None, control.name
    return control.record


def _bytes(record: ParsedRecord) -> bytes:
    """The record's bytes with its alignment padding zeroed: the designer leaves whatever was in memory
    there (the three bytes after a five-letter face name can read "a\\0\\0"), which is not a rule to
    reproduce."""
    padded = replace(record, pads={}, text_props=replace(record.text_props, pads={}) if record.text_props else None)
    return serialize_record(padded, "cp1252")


def _text_props(control: FormControl) -> bytes:
    text_props = _record(control).text_props
    assert text_props is not None, control.name
    return _bytes(text_props)


#: Where the designer added a control of a kind: the form, the container (None for the form itself), the kind,
#: and the control it added there.
FONT_CASES = [
    *[("Plain", None, kind, f"{kind}1") for kind in
      ("Label", "CommandButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton", "ToggleButton",
       "TabStrip")],
    ("Fonted", None, "Label", "Label1"),
    ("Fonted", None, "CommandButton", "CommandButton1"),
    ("Fonted", None, "CheckBox", "CheckBox1"),
    ("Fonted", None, "ToggleButton", "ToggleButton1"),
    ("Fonted", None, "TextBox", "TextBox1"),
    ("Fonted", "Frame1", "Label", "Label2"),
    ("Fonted", "Frame1", "ToggleButton", "ToggleButton2"),
    ("Fonted", "Page1", "Label", "Label3"),
    ("Neighbour", None, "ToggleButton", "ToggleButton1"),
    ("Effects", None, "Label", "Label1"),
    ("Effects", None, "CommandButton", "CommandButton1"),
    ("Effects", None, "CheckBox", "CheckBox1"),
    ("Effects", "Frame1", "Label", "Label2"),
    ("Charset", None, "Label", "Label1"),
    ("Charset", None, "TextBox", "TextBox1"),
]


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize(("form", "container", "kind", "theirs"), FONT_CASES,
                         ids=[f"{form}-{container or 'form'}-{kind}" for form, container, kind, _ in FONT_CASES])
def test_a_new_control_takes_its_containers_font(host: str, form: str, container: str | None, kind: str,
                                                  theirs: str, tmp_path: Path) -> None:
    # The face, charset, effects and weight of the form or Frame the control lands on, its size in whole
    # twips, and centred text only on a button; nothing from the controls already there. The Neighbour
    # form's ToggleButton sits beside a CheckBox in Courier New and still takes the form's Tahoma.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        added = office.add_control(kind, "Added", container=container, left=0, top=0)
        assert _text_props(added) == _text_props(office.control(theirs))


def _site(form: VBAForm, name: str) -> Any:
    """The site the form's ``f`` streams give a control: its flags, id and tab index, which the public
    FormControl does not all carry."""
    for level in form._levels:  # pyright: ignore[reportPrivateUsage]
        for site in level.sites:
            if site.name == name:
                return site
    raise KeyError(name)


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("kind", ["Label", "CommandButton", "TextBox", "CheckBox", "Image", "ScrollBar"])
def test_a_new_control_is_sited_with_the_designers_flags(host: str, kind: str, tmp_path: Path) -> None:
    # A Label takes no focus: the designer stores its flags without TabStop (0x32), where every other kind
    # leaves them at the default and stores none.
    with ExcelFile(_office_form(host, "Plain", tmp_path)) as workbook:
        office = _form(workbook, "Plain")
        office.add_control(kind, "Added", left=0, top=0)
        ours, theirs = _site(office, "Added"), _site(office, f"{kind}1")
        assert (ours.mask, ours.values.get("BitFlags")) == (theirs.mask, theirs.values.get("BitFlags"))


def _composed(tmp_path: Path, form: str) -> tuple[ExcelFile, VBAForm]:
    """A new workbook and a new form of that name, to compose as the designer composed its own."""
    workbook = ExcelFile.create_new(tmp_path / f"{form}.xlsm")
    return workbook, workbook.add_form(form)


def _kind(name: str) -> str:
    """The kind of a control the designer named: Label1 is a Label."""
    return name.rstrip("0123456789")


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", ["Tabs", "Plain"])
def test_images_stay_last_in_the_tab_order(host: str, form: str, tmp_path: Path) -> None:
    # The designer's forms had their controls added in the order their sites list them; the Images end up
    # last in the tab order, whenever they were added.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        names = [control.name for control in office.controls]
        theirs = [_site(office, name).tab_index for name in names]
    workbook, composed = _composed(tmp_path, form)
    with workbook:
        for name in names:
            composed.add_control(_kind(name), name, left=0, top=0)
        assert [_site(composed, name).tab_index for name in names] == theirs


#: How the designer built a form, as the library's calls: (kind, name, container), a kind of "Page" adding a
#: page named ``name`` to the MultiPage named in place of a container.
STEPS: dict[str, list[tuple[str, str, str | None]]] = {
    "Frames": [("Frame", "Frame1", None), ("Frame", "Frame2", None), ("Label", "Label1", "Frame2")],
    "Nested": [("MultiPage", "MultiPage1", None), ("Frame", "Frame1", "Page1"), ("Label", "Label1", "Frame1"),
               ("Page", "Page3", "MultiPage1"), ("Page", "Page4", "MultiPage1"),
               ("CommandButton", "CommandButton1", "Page2")],
    "Plain": [(kind, f"{kind}1", None) for kind in
              ("Label", "CommandButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton",
               "ToggleButton", "Frame", "MultiPage", "Image", "SpinButton", "ScrollBar", "TabStrip")],
    "Tabs": [("Image", "Image1", None), ("Label", "Label1", None), ("Image", "Image2", None),
             ("CommandButton", "CommandButton1", None), ("TextBox", "TextBox1", None)],
}


def _compose(form: VBAForm, steps: list[tuple[str, str, str | None]]) -> None:
    for kind, name, container in steps:
        if kind == "Page":
            assert container is not None
            form.add_page(container, name=name)
        else:
            form.add_control(kind, name, container=container, left=0, top=0)


def _counters(form: VBAForm) -> dict[str, tuple[object, ...]]:
    """What the form and each Frame, MultiPage and page store of their counters and their effect."""
    fields = ("NextAvailableID", "ShapeCookie", "SpecialEffect")
    out = {"": tuple(form.get(name) for name in fields)}
    for control in form.walk():
        if control.is_container:
            out[control.name] = tuple(control.get(name) for name in fields)
    return out


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", list(STEPS))
def test_containers_count_what_the_designer_counts(host: str, form: str, tmp_path: Path) -> None:
    # A new Frame is etched and stores no NextAvailableID or ShapeCookie until something is added to it; a
    # control counts one in its container and every container above it, a page two in its MultiPage alone.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        theirs = _counters(_form(workbook, form))
    workbook, composed = _composed(tmp_path, form)
    with workbook:
        _compose(composed, STEPS[form])
        assert _counters(composed) == theirs


def _size(control: FormControl) -> object:
    return control.get("DisplayedSize" if control.is_container else "Size")


@pytest.mark.parametrize("host", HOSTS)
def test_a_new_control_takes_the_designers_size(host: str, tmp_path: Path) -> None:
    with ExcelFile(_office_form(host, "Plain", tmp_path)) as workbook:
        theirs = {control.name: _size(control) for control in _form(workbook, "Plain").controls}
    workbook, composed = _composed(tmp_path, "Plain")
    with workbook:
        _compose(composed, STEPS["Plain"])
        assert {control.name: _size(control) for control in composed.controls} == theirs


FONTS: dict[str, dict[str, dict[str, Any]]] = json.loads((FIXTURES / "form_fonts.json").read_text(encoding="utf-8"))
_STDFONT = bytes.fromhex("0352e30b918fce119de300aa004bb851")


def _with_font(form: VBAForm, face: str, size: int) -> None:
    """Give a composed form the StdFont the designer stored for one of its own; the library sets no form
    font, so this writes the form's record as the designer left it."""
    level = form._levels[0]  # pyright: ignore[reportPrivateUsage]
    name = face.encode("cp1252")
    level.stream.font_raw = _STDFONT + bytes([1, 0, 0, 0]) + (400).to_bytes(2, "little") \
        + size.to_bytes(4, "little") + bytes([len(name)]) + name
    level.record.mask |= 1 << 20
    level.record.values["Font"] = 0xFFFF


FONT_ROWS = [(host, key) for host in FONTS for key in FONTS[host]]


@pytest.mark.parametrize(("host", "key"), FONT_ROWS, ids=[f"{host}-{key}" for host, key in FONT_ROWS])
def test_a_new_controls_size_follows_the_font(host: str, key: str, tmp_path: Path) -> None:
    # Each form was set to one font; a CheckBox or OptionButton is at least 24 pixels tall and grows with
    # it, every other kind keeps its size.
    measured = FONTS[host][key]
    stored: dict[str, Any] = measured["font"] or {"name": "Tahoma", "cy_size": 82500}
    workbook, composed = _composed(tmp_path, "Fonted")
    with workbook:
        _with_font(composed, str(stored["name"]), int(stored["cy_size"]))
        for name in measured["sizes"]:
            added = composed.add_control(_kind(name), name, left=0, top=0)
            size = added.get("Size")
            assert size == Size(*measured["sizes"][name]), name


def _pages(form: VBAForm, multipage: str) -> list[tuple[object, object]]:
    """A MultiPage's TabStrip size, then where each page is sited and its size."""
    tabstrip, *pages = form.control(multipage).children
    return [("TabStrip", tabstrip.get("Size")),
            *[(_site(form, page.name).position, page.get("DisplayedSize")) for page in pages]]


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", ["Plain", "Nested"])
def test_pages_sit_under_the_tabs_of_their_multipage(host: str, form: str, tmp_path: Path) -> None:
    # The TabStrip is as large as the MultiPage; the pages sit two pixels inside it and under the tabs.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        theirs = _pages(_form(workbook, form), "MultiPage1")
    workbook, composed = _composed(tmp_path, form)
    with workbook:
        _compose(composed, STEPS[form])
        assert _pages(composed, "MultiPage1") == theirs


@pytest.mark.parametrize(("host", "key"), FONT_ROWS, ids=[f"{host}-{key}" for host, key in FONT_ROWS])
def test_the_tabs_are_as_tall_as_the_font_makes_them(host: str, key: str, tmp_path: Path) -> None:
    measured = FONTS[host][key]
    stored: dict[str, Any] = measured["font"] or {"name": "Tahoma", "cy_size": 82500}
    workbook, composed = _composed(tmp_path, "Fonted")
    with workbook:
        _with_font(composed, str(stored["name"]), int(stored["cy_size"]))
        composed.add_control("MultiPage", "MultiPage1", left=0, top=0)
        multipage = measured["multipage"]
        assert _pages(composed, "MultiPage1") == [
            ("TabStrip", Size(*multipage["tabstrip"])),
            *[(tuple(page["position"]), Size(*page["size"])) for page in multipage["pages"]]]


@pytest.mark.parametrize("host", HOSTS)
def test_a_page_added_to_the_designers_multipage_is_the_one_it_adds(host: str, tmp_path: Path) -> None:
    # The Fonted form's first MultiPage shows Arial 9.75 pt, and the designer added its Page3 with Pages.Add.
    with ExcelFile(_office_form(host, "Fonted", tmp_path)) as workbook:
        office = _form(workbook, "Fonted")
        office.add_page("MultiPage1", name="Added")
        added, theirs = office.control("Added"), office.control("Page3")
        assert (_site(office, "Added").position, added.get("DisplayedSize")) == (
            _site(office, "Page3").position, theirs.get("DisplayedSize"))


def _tab_arrays(control: FormControl) -> dict[str, object]:
    """A TabStrip's five arrays: each entry's text and compression, then each array's bytes, whose
    alignment after an entry like "Page1" is whatever the designer had in memory."""
    record = _record(control)
    entries = {name: [(entry.text, entry.compressed) for entry in parse_string_array(blob, "cp1252")]
               for name, blob in record.arrays.items()}
    exact = {f"{name} bytes": blob for name, blob in record.arrays.items() if name != "Items"}
    return {**entries, **exact}


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", ["Plain", "Nested"])
def test_a_multipages_tabs_are_the_designers(host: str, form: str, tmp_path: Path) -> None:
    # A tab's empty tip, tag and accelerator are stored as a count of zero with no compression flag.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        theirs = _tab_arrays(_form(workbook, form).control("MultiPage1").children[0])
    workbook, composed = _composed(tmp_path, form)
    with workbook:
        _compose(composed, STEPS[form])
        assert _tab_arrays(composed.control("MultiPage1").children[0]) == theirs


@pytest.mark.parametrize("host", HOSTS)
def test_a_new_tabstrip_has_the_designers_two_tabs(host: str, tmp_path: Path) -> None:
    # Tab1 and Tab2, with room allocated for those two alone, where a MultiPage's TabStrip keeps two spare.
    with ExcelFile(_office_form(host, "Plain", tmp_path)) as workbook:
        office = _form(workbook, "Plain")
        added = office.add_control("TabStrip", "Added", left=0, top=0)
        assert _bytes(_record(added)) == _bytes(_record(office.control("TabStrip1")))


#: The kinds add_control gives a caption, their name, where the designer's Controls.Add gives none.
CAPTIONED = frozenset({"Label", "CommandButton", "ToggleButton", "CheckBox", "OptionButton", "Frame"})


def _unpadded_depths(raw: bytes, sites: int) -> bytes:
    """SiteDepthsAndTypes with the alignment after its entries zeroed."""
    at = covered = 0
    while covered < sites:
        counted = bool(raw[at + 1] & 0x80)
        covered += raw[at + 1] & 0x7F if counted else 1
        at += 3 if counted else 2
    return raw[:at] + bytes(len(raw) - at)


def _streams(form: VBAForm) -> dict[str, tuple[bytes, bytes]]:
    """Every container's ``f`` and ``o`` as the library writes them, by path below the form, with the
    alignment padding zeroed wherever the designer leaves whatever was in memory."""
    levels = form._levels  # pyright: ignore[reportPrivateUsage]
    below = len(levels[0].path)
    out: dict[str, tuple[bytes, bytes]] = {}
    for level in levels:
        level.record.pads = {}
        if level.stream.depths_raw:  # a container read from a file; a new one composes its own
            level.stream.depths_raw = _unpadded_depths(level.stream.depths_raw, len(level.sites))
        for site in level.sites:
            site.pads = {}
        for control in level.controls:
            record = control.record
            if record is None:
                continue
            record.pads = {}
            if record.text_props is not None:
                record.text_props = replace(record.text_props, pads={})
            for name, blob in record.arrays.items():
                entries = parse_string_array(blob, "cp1252")
                for entry in entries:
                    entry.pad = b""
                record.arrays[name] = serialize_string_array(entries, "cp1252")
        out["/".join(level.path[below:])] = level.serialize("cp1252")
    return out


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize("form", list(STEPS))
def test_a_composed_form_is_the_designers_byte_for_byte(host: str, form: str, tmp_path: Path) -> None:
    # Every container's f and o, the form's own included, as the designer wrote them for the same controls
    # added in the same order: all but the captions add_form and add_control give on purpose, cleared here,
    # and the alignment the designer leaves as whatever was in memory. A lone site is listed in the plain
    # form of SiteDepthsAndTypes, and a run of them in the counted form.
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        theirs = _streams(_form(workbook, form))
    workbook, composed = _composed(tmp_path, form)
    with workbook:
        _compose(composed, STEPS[form])
        composed.set_property("Caption", None)
        for control in composed.walk():
            if control.kind.rpartition(".")[2] in CAPTIONED:
                control.set_property("Caption", None)
        assert _streams(composed) == theirs


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize(("form", "theirs"), [("Plain", "MultiPage1"), ("Fonted", "MultiPage1")])
def test_a_new_multipages_tabs_show_the_containers_font(host: str, form: str, theirs: str,
                                                         tmp_path: Path) -> None:
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        added = office.add_control("MultiPage", "Added", left=0, top=0)
        # The first child of a MultiPage is its TabStrip, which draws the tabs.
        assert _text_props(added.children[0]) == _text_props(office.control(theirs).children[0])
