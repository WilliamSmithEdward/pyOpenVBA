"""A UserForm in a binary .xls or .doc lives in the project's storage, where Excel and Word keep one.

tests/fixtures/binary_forms.json is what scripts/measure_binary_forms.py
saw when Excel and Word inserted a UserForm through their editor into a
binary file: the form's storage inside _VBA_PROJECT_CUR or Macros, beside
VBA/, and nothing new at the file's root. tests/fixtures/binary_project/
holds the files each application saved just before, which these tests give
a form of the library's. The library used to put the storage at the file's
root, which both applications refuse to open, and read no form from a
binary file at all.
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, VBAProjectError, WordFile
from pyopenvba._references import MSFORMS_GUID
from pyopenvba.cfb import CFB

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "binary_forms.json").read_text(encoding="utf-8"))
HOSTS: dict[str, tuple[type[ExcelFile] | type[WordFile], str]] = {
    "excel": (ExcelFile, "workbook.xls"), "word": (WordFile, "document.doc")}


def _copy(host: str, tmp_path: Path) -> Path:
    target = tmp_path / HOSTS[host][1]
    shutil.copyfile(FIXTURES / "binary_project" / HOSTS[host][1], target)
    return target


def _with_form(host: str, tmp_path: Path) -> Path:
    """The fixture given the form the editor made, UserForm1, through add_form."""
    target = _copy(host, tmp_path)
    with HOSTS[host][0](target) as office_file:
        office_file.add_form("UserForm1")
        office_file.save()
    return target


def _tree(cfb: CFB, path: tuple[str, ...]) -> list[str]:
    """Every storage (ending in /) and stream under ``path``, as the measurement lists them."""
    lines = sorted(cfb.list_streams_at(path))
    for storage in sorted(cfb.list_storages_at(path)):
        lines.append(storage + "/")
        lines += [f"{storage}/{line}" for line in _tree(cfb, (*path, storage))]
    return lines


def _storage_entry(raw: bytes, name: str) -> bytes:
    """The directory entry of the one storage named ``name``: sector-aligned, 128 bytes, the UTF-16LE
    name at 0, its byte length at 64 and the object type at 66."""
    needle = name.encode("utf-16-le") + b"\x00\x00"
    found = [raw[at:at + 128] for at in range(0, len(raw) - 127, 128)
             if raw[at:at + len(needle)] == needle and raw[at + 64] == len(needle) and raw[at + 66] == 1]
    assert len(found) == 1, f"{len(found)} storages named {name!r}"
    return found[0]


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_fixture_is_the_file_the_application_saved(host: str) -> None:
    measured = RECORD[host]
    cfb = CFB.from_bytes((FIXTURES / "binary_project" / HOSTS[host][1]).read_bytes())
    assert _tree(cfb, (measured["project_storage"],)) == measured["base_tree"]


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_new_form_lands_where_the_editor_puts_one(host: str, tmp_path: Path) -> None:
    measured = RECORD[host]
    raw = _with_form(host, tmp_path).read_bytes()
    cfb = CFB.from_bytes(raw)
    assert sorted(cfb.list_storages_at()) + sorted(cfb.list_streams_at()) == measured["root"]
    # The editor keeps its performance caches, __SRP_0 and on; a writer must not emit them ([MS-OVBA] 2.3.4.1),
    # and every save of the library drops them.
    assert _tree(cfb, (measured["project_storage"],)) == [
        line for line in measured["tree"] if not line.startswith("VBA/__SRP_")]
    for name, clsid in measured["clsids"].items():
        assert _storage_entry(raw, name)[80:96].hex() == clsid
    # The editor also stamps the storage with the time of the save, which is not a rule to reproduce.
    assert _storage_entry(raw, "UserForm1")[96:100].hex() == measured["form_entry"]["state"]


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_form_is_the_editors_but_for_the_caption_add_form_sets(host: str, tmp_path: Path) -> None:
    measured = RECORD[host]
    target = _with_form(host, tmp_path)
    path = (measured["project_storage"], "UserForm1")
    written = CFB.from_bytes(target.read_bytes())
    streams = {name: written.get_stream_at(path, name).hex() for name in written.list_streams_at(path)}
    assert set(streams) == set(measured["form_streams"])
    for name in ("\x01CompObj", "\x03VBFrame", "o"):
        assert streams[name] == measured["form_streams"][name], name
    # add_form sets the caption it is given, the form's name by default, and the editor stores none in f until
    # one is set. With it cleared the form is the editor's, byte for byte.
    with HOSTS[host][0](target) as office_file:
        form = office_file.forms()[0]
        assert form.get("Caption") == "UserForm1"
        form.set_property("Caption", None)
        office_file.save()
    cleared = CFB.from_bytes(target.read_bytes()).get_stream_at(path, "f")
    assert cleared.hex() == measured["form_streams"]["f"]


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_project_declares_it_as_a_designer(host: str, tmp_path: Path) -> None:
    measured = RECORD[host]
    cfb = CFB.from_bytes(_with_form(host, tmp_path).read_bytes())
    lines = cfb.get_stream_at((measured["project_storage"],), "PROJECT").decode("latin-1").split("\r\n")

    def declarations(project: list[str]) -> list[str]:
        return project[1:next(i for i, line in enumerate(project) if line.startswith("Name="))]

    assert declarations(lines) == declarations(measured["project_lines"])


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_form_is_read_edited_and_trimmed_in_place(host: str, tmp_path: Path) -> None:
    opener, _ = HOSTS[host]
    storage = RECORD[host]["project_storage"]
    target = _copy(host, tmp_path)
    with opener(target) as office_file:
        assert office_file.forms() == []
        form = office_file.add_form("Wizard", caption="Setup")
        form.add_control("CommandButton", "Go", left=12, top=12)
        form.add_control("Frame", "Group", left=12, top=48, width=200, height=90)
        form.add_control("OptionButton", "First", container="Group", left=6, top=12)
        office_file.save()
    cfb = CFB.from_bytes(target.read_bytes())
    # The frame's own storage nests in the form's, inside the project's storage; the root is untouched.
    assert cfb.list_storages_at((storage, "Wizard")) == ["i02"]
    assert "Wizard" not in cfb.list_storages_at()
    with opener(target) as office_file:
        form = office_file.forms()[0]
        assert (form.name, form.get("Caption")) == ("Wizard", "Setup")
        assert [control.name for control in form.walk()] == ["Go", "Group", "First"]
        form.set_property("Caption", "Settings")
        form.remove_control("Group")
        office_file.save()
    with opener(target) as office_file:
        form = office_file.forms()[0]
        assert form.get("Caption") == "Settings"
        assert [control.name for control in form.walk()] == ["Go"]
    cfb = CFB.from_bytes(target.read_bytes())
    assert cfb.list_storages_at((storage, "Wizard")) == []


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_editors_form_is_read_where_the_editor_keeps_it(host: str, tmp_path: Path) -> None:
    # The storage the editor wrote, with its bytes, placed where it placed it and with no help from add_form.
    measured = RECORD[host]
    storage = measured["project_storage"]
    target = _copy(host, tmp_path)
    cfb = CFB.from_bytes(target.read_bytes())
    cfb.add_substorage_at((storage,), "UserForm1")
    for name, data in measured["form_streams"].items():
        cfb.add_stream_at((storage, "UserForm1"), name, bytes.fromhex(data))
    target.write_bytes(cfb.to_bytes())
    with HOSTS[host][0](target) as office_file:
        assert [(form.name, form.get("Caption")) for form in office_file.forms()] == [("UserForm1", None)]
        office_file.add_reference("MSForms", MSFORMS_GUID, 2, 0)
        with pytest.raises(VBAProjectError, match="required by UserForms: UserForm1"):
            office_file.remove_reference("MSForms")


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_forms_reference_stays_while_a_form_does(host: str, tmp_path: Path) -> None:
    with HOSTS[host][0](_with_form(host, tmp_path)) as office_file:
        assert [ref.kind for ref in office_file.references() if ref.name == "MSForms"] == ["control"]
        with pytest.raises(VBAProjectError, match="required by UserForms: UserForm1"):
            office_file.remove_reference("MSForms")


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_second_form_of_the_same_name_is_refused(host: str, tmp_path: Path) -> None:
    with HOSTS[host][0](_with_form(host, tmp_path)) as office_file, pytest.raises(Exception, match="UserForm1"):
        office_file.add_form("UserForm1")
