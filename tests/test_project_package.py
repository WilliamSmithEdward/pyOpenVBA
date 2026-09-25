"""The PROJECT stream's Package line, written and kept as Excel's and Word's editors keep it.

tests/fixtures/project_package.json is what scripts/measure_project_package.py
saw when each application built a project through its editor and saved
after every step. The editor writes
Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F} once, immediately before the
BaseClass line of the form it adds to a stream that has no such line, and
never moves or removes it afterwards: not when a form or the last form goes,
and not for a project that declares its forms without one. These tests run
the same steps through the library, from the files each application saved
with one module (tests/fixtures/binary_project/), and compare the
declarations and the form storages after every save.
"""

from __future__ import annotations

import json
import shutil
import zipfile
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, WordFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import serialize_project_stream

FIXTURES = Path(__file__).parent / "fixtures"
RECORD: dict[str, dict[str, Any]] = json.loads((FIXTURES / "project_package.json").read_text(encoding="utf-8"))
HOSTS: dict[str, tuple[type[ExcelFile] | type[WordFile], str, str, str]] = {
    "excel": (ExcelFile, "workbook.xls", "xlsm", "xl/vbaProject.bin"),
    "word": (WordFile, "document.doc", "docm", "word/vbaProject.bin"),
}
#: The measured chains as the library runs them: each state and the steps that lead to it from the one before.
CHAIN = [("first_form", ["form UserForm1"]), ("module_after_form", ["module Module2"]),
         ("second_form_later", ["form UserForm2"])]
AT_ONCE = [("two_forms_at_once", ["form UserForm1", "form UserForm2"])]
LEGACY = [("legacy_module_added.macro", ["module Module2"]), ("legacy_second_form.macro", ["form UserForm2"])]


def _declarations(project: list[str]) -> list[str]:
    """The lines between ID= and Name=: the modules, the packages and the forms, in order."""
    return project[1:next(i for i, line in enumerate(project) if line.startswith("Name="))]


def _project(host: str, path: Path) -> tuple[list[str], list[str]]:
    """The PROJECT stream's lines and the storages beside VBA/ of a binary file."""
    cfb = CFB.from_bytes(path.read_bytes())
    root = (RECORD[host]["project_storage"],)
    storages = sorted(name for name in cfb.list_storages_at(root) if name.casefold() != "vba")
    return cfb.get_stream_at(root, "PROJECT").decode("latin-1").split("\r\n"), storages


def _copy(host: str, tmp_path: Path) -> Path:
    target = tmp_path / HOSTS[host][1]
    shutil.copyfile(FIXTURES / "binary_project" / HOSTS[host][1], target)
    return target


def _run(host: str, path: Path, steps: list[str]) -> None:
    """Take the steps through the library and save once, as the editor saved once for each state."""
    with HOSTS[host][0](path) as office_file:
        for step in steps:
            verb, _, name = step.partition(" ")
            if verb == "form":
                office_file.add_form(name)
            elif verb == "module":
                office_file.vba_project().add_module(name, f'Attribute VB_Name = "{name}"\r\n')
            else:
                office_file.vba_project().delete_module(name)
        office_file.save()


def _replay(host: str, path: Path, chain: list[tuple[str, list[str]]]) -> None:
    for state, steps in chain:
        _run(host, path, steps)
        project, storages = _project(host, path)
        measured = RECORD[host]["states"][state]
        assert _declarations(project) == _declarations(measured["project"]), state
        assert storages == measured["storages"], state


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_fixture_declares_what_the_measured_project_did(host: str) -> None:
    project, _ = _project(host, FIXTURES / "binary_project" / HOSTS[host][1])
    assert _declarations(project) == _declarations(RECORD[host]["states"]["module"]["project"])


@pytest.mark.parametrize("host", list(HOSTS))
def test_the_file_type_does_not_matter(host: str) -> None:
    states = RECORD[host]["states"]
    for state in ("module", "first_form"):
        assert _declarations(states[f"{state}.macro"]["project"]) == _declarations(states[state]["project"])


@pytest.mark.parametrize("host", list(HOSTS))
def test_forms_come_and_go_as_the_editor_declares_them(host: str, tmp_path: Path) -> None:
    _replay(host, _copy(host, tmp_path), CHAIN)


@pytest.mark.parametrize("host", list(HOSTS))
def test_two_forms_in_one_save_share_the_package(host: str, tmp_path: Path) -> None:
    _replay(host, _copy(host, tmp_path), AT_ONCE)


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_module_goes_after_a_package_left_last(host: str) -> None:
    # The editor's own stream with the Package line last among its declarations, then a module added.
    states = RECORD[host]["states"]
    raw = "\r\n".join(states["package_left_last"]["project"]).encode("latin-1")
    written = serialize_project_stream(raw, {}, add_modules=[("Module2", "Module")])
    assert _declarations(written.decode("latin-1").split("\r\n")) == _declarations(
        states["module_after_package"]["project"])


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_project_without_the_line_gets_it_with_its_next_form(host: str, tmp_path: Path) -> None:
    # A form declared with no Package line, as pyOpenVBA 6.1.2 and earlier wrote one.
    target = _copy(host, tmp_path)
    _run(host, target, ["form UserForm1"])
    cfb = CFB.from_bytes(target.read_bytes())
    root = (RECORD[host]["project_storage"],)
    lines = cfb.get_stream_at(root, "PROJECT").split(b"\r\n")
    cfb.write_stream_at(root, "PROJECT", b"\r\n".join(line for line in lines if not line.startswith(b"Package=")))
    target.write_bytes(cfb.to_bytes())
    _replay(host, target, LEGACY)


@pytest.mark.parametrize("host", list(HOSTS))
def test_a_macro_enabled_file_gets_the_line_too(host: str, tmp_path: Path) -> None:
    opener, _, suffix, entry = HOSTS[host]
    target = tmp_path / f"form.{suffix}"
    with opener.create_new(target) as office_file:
        office_file.add_form("UserForm1")
        office_file.save()
    with zipfile.ZipFile(target) as package:
        project = CFB.from_bytes(package.read(entry)).get_stream("PROJECT").decode("latin-1").split("\r\n")
    assert _declarations(project) == _declarations(RECORD[host]["states"]["first_form.macro"]["project"])

