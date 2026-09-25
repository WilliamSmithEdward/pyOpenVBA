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
from pyopenvba._oforms_records import ParsedRecord, serialize_record
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


@pytest.mark.parametrize("host", HOSTS)
@pytest.mark.parametrize(("form", "theirs"), [("Plain", "MultiPage1"), ("Fonted", "MultiPage1")])
def test_a_new_multipages_tabs_show_the_containers_font(host: str, form: str, theirs: str,
                                                         tmp_path: Path) -> None:
    with ExcelFile(_office_form(host, form, tmp_path)) as workbook:
        office = _form(workbook, form)
        added = office.add_control("MultiPage", "Added", left=0, top=0)
        # The first child of a MultiPage is its TabStrip, which draws the tabs.
        assert _text_props(added.children[0]) == _text_props(office.control(theirs).children[0])
