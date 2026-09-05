"""The Access surface that mirrors the other hosts: module_names / get_module /
set_module, vba_project().add_module and friends, pull and push of module
files, and forms edited through the same calls a UserForm takes."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from pyopenvba import AccessDatabase, VBAModuleKind, pull_access, push_access
from pyopenvba.access import AccessControl, AccessForm, AccessVBAProject
from pyopenvba.access._designs import OPEN_CONTROL, OPEN_SECTION, OPEN_SIBLING
from pyopenvba.access_read import AccessError

TEMPLATE = Path(__file__).parents[1] / "src" / "pyopenvba" / "_templates" / "blank_files" / "blank_database.accdb"


@pytest.fixture
def blank(tmp_path: Path) -> AccessDatabase:
    out = tmp_path / "blank.accdb"
    shutil.copyfile(TEMPLATE, out)
    return AccessDatabase(out)


# -- modules, the way the other hosts read and write them ----------------------


def test_module_names_get_module_and_set_module_read_like_the_other_hosts(blank: AccessDatabase) -> None:
    assert blank.module_names() == [m.name for m in blank.modules()]
    name = blank.module_names()[0]
    assert blank.get_module(name) == blank.module(name).source

    blank.set_module(name, "Option Compare Database\r\n\r\nPublic Sub Go()\r\nEnd Sub")

    assert "Public Sub Go()" in blank.get_module(name)


def test_set_module_refuses_a_module_that_is_not_there(blank: AccessDatabase) -> None:
    with pytest.raises(AccessError):
        blank.set_module("Nowhere", "Option Compare Database")


def test_add_module_takes_the_kinds_the_other_hosts_take(blank: AccessDatabase) -> None:
    standard = blank.add_module("Helpers", "Option Compare Database\r\nPublic Sub A()\r\nEnd Sub")
    klass = blank.add_module("Widget", "Option Compare Database", kind=VBAModuleKind.other)
    by_string = blank.add_module("Gadget", kind="class")

    assert (standard.kind, klass.kind, by_string.kind) == ("module", "class", "class")
    assert {"Helpers", "Widget", "Gadget"} <= set(blank.module_names())
    with pytest.raises(AccessError):
        blank.add_module("Odd", kind="document")


def test_vba_project_has_the_shape_of_a_host_project(blank: AccessDatabase) -> None:
    project = blank.vba_project()
    assert isinstance(project, AccessVBAProject)
    before = project.module_names()

    added = project.add_module("Tools", "Option Compare Database", kind=VBAModuleKind.standard)
    renamed = project.rename_module("Tools", "Kit")
    project.delete_module("Kit")

    assert added.name == "Tools" and renamed.name == "Kit"
    assert project.module_names() == before
    assert [m.name for m in project.modules] == before
    assert isinstance(project.references, list)


def test_pull_and_push_round_trip_the_module_files(blank: AccessDatabase, tmp_path: Path) -> None:
    blank.add_module("Helpers", "Option Compare Database\r\n\r\nPublic Sub A()\r\nEnd Sub")
    blank.add_module("Widget", "Option Compare Database\r\nPublic Total As Long", kind=VBAModuleKind.other)
    pulled = blank.pull_modules(tmp_path / "vba")

    names = {p.name for p in pulled}
    assert {"Helpers.bas", "Widget.cls"} <= names
    (tmp_path / "vba" / "Helpers.bas").write_text("Option Compare Database\n\nPublic Sub B()\nEnd Sub\n", encoding="utf-8")
    (tmp_path / "vba" / "Stray.bas").write_text("Public Sub C()\nEnd Sub\n", encoding="utf-8")

    updated = blank.push_modules(tmp_path / "vba")

    assert "Helpers" in updated and "Stray" not in updated
    assert "Public Sub B()" in blank.get_module("Helpers")
    with pytest.raises(KeyError):
        blank.push_modules(tmp_path / "vba", strict=True)


def test_push_drops_a_vbe_export_preamble_and_attribute_lines(blank: AccessDatabase, tmp_path: Path) -> None:
    blank.add_module("Widget", "Option Compare Database", kind=VBAModuleKind.other)
    folder = tmp_path / "vba"
    folder.mkdir()
    (folder / "Widget.cls").write_text(
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"Widget\"\n"
        "Attribute VB_Exposed = False\nOption Compare Database\nPublic Sub D()\nEnd Sub\n",
        encoding="utf-8",
    )

    blank.push_modules(folder)

    source = blank.get_module("Widget")
    assert source.startswith("Option Compare Database") and "VERSION" not in source and "Attribute VB_" not in source


def test_push_access_and_pull_access_work_on_files(tmp_path: Path) -> None:
    database = tmp_path / "db.accdb"
    shutil.copyfile(TEMPLATE, database)
    with AccessDatabase(database) as db:
        db.add_module("Helpers", "Option Compare Database\r\nPublic Sub A()\r\nEnd Sub")
        db.save()
    pulled = pull_access(database, tmp_path / "vba")
    assert any(p.name == "Helpers.bas" for p in pulled)
    (tmp_path / "vba" / "Helpers.bas").write_text("Option Compare Database\r\nPublic Sub Z()\r\nEnd Sub\r\n", encoding="utf-8")

    updated = push_access(tmp_path / "vba", database, out=tmp_path / "out.accdb")

    assert set(updated) == {"Helpers", "Module1"}  # every file that matches a module is pushed
    assert "Public Sub Z()" in AccessDatabase(tmp_path / "out.accdb").get_module("Helpers")
    assert "Public Sub A()" in AccessDatabase(database).get_module("Helpers")


def test_the_cli_pushes_into_an_access_database(tmp_path: Path) -> None:
    database = tmp_path / "db.accdb"
    shutil.copyfile(TEMPLATE, database)
    with AccessDatabase(database) as db:
        db.add_module("Helpers", "Option Compare Database")
        db.save()
    folder = tmp_path / "vba"
    folder.mkdir()
    (folder / "Helpers.bas").write_text("Option Compare Database\r\nPublic Sub Cli()\r\nEnd Sub\r\n", encoding="utf-8")

    done = subprocess.run(
        [sys.executable, "-m", "pyopenvba", "access-push", str(folder), str(database)],
        capture_output=True,
        text=True,
        cwd=Path(__file__).parents[1],
        timeout=300,
    )

    assert done.returncode == 0, done.stderr
    assert done.stdout.split() == ["Helpers"]
    assert "Public Sub Cli()" in AccessDatabase(database).get_module("Helpers")


# -- forms, the way a UserForm is edited ---------------------------------------


def test_add_form_returns_a_form_with_the_host_surface(blank: AccessDatabase) -> None:
    form = blank.add_form("Summary", caption="Totals", width=8000, height=3000)

    assert isinstance(form, AccessForm)
    assert form.name == "Summary" and form.kind == "form"
    assert form.get("Caption") == "Totals" and form.get("Width") == 8000
    assert form.control("Detail").get("Height") == 3000
    assert form.controls == () and form.walk() == []
    assert [f.name for f in blank.forms()] == ["Summary"]
    assert isinstance(blank.form("Summary"), AccessForm)


def test_controls_go_on_and_read_back_through_the_form(blank: AccessDatabase) -> None:
    form = blank.add_form("Built")
    title = form.add_control("Label", "Title", left=240, top=240, width=2000, height=300, caption="Hello")
    form.add_control("TextBox", "Box", top=700, caption="=1+1")

    assert isinstance(title, AccessControl)
    assert [c.name for c in form.controls] == ["Title", "Box"]
    assert [c.kind for c in form.walk()] == ["Label", "TextBox"]
    assert form.control("Title").get("Caption") == "Hello"
    assert form.control("Box").properties()["Left"] == 0
    assert form.control("box").name == "Box"


def test_a_property_set_through_the_control_lands_on_the_design(blank: AccessDatabase) -> None:
    form = blank.add_form("Props")
    form.add_control("Label", "Title", caption="Hello")

    form.control("Title").set_property("Caption", "Changed")
    form.set_property("Caption", "Form caption")

    fresh = blank.form("Props")
    assert fresh.control("Title").get("Caption") == "Changed"
    assert fresh.get("Caption") == "Form caption"


def test_removing_a_control_re_marks_the_ones_left(blank: AccessDatabase) -> None:
    form = blank.add_form("Marks")
    for name in ("A", "B", "C"):
        form.add_control("Label", name, caption=name)
    assert [c.marker for c in form.controls] == [OPEN_CONTROL, OPEN_SIBLING, OPEN_SIBLING]
    assert form.controls[0].code == 3

    form.remove_control("B")
    assert [c.name for c in form.controls] == ["A", "C"]
    assert [c.marker for c in form.controls] == [OPEN_CONTROL, OPEN_SIBLING]
    assert form.controls[0].code == 2

    form.remove_control("A")
    assert [c.name for c in form.controls] == ["C"]
    assert form.controls[0].marker == OPEN_SECTION

    form.remove_control("C")
    assert form.controls == ()
    with pytest.raises(AccessError):
        form.remove_control("C")


def test_the_tab_order_closes_up_behind_a_removed_control(blank: AccessDatabase) -> None:
    """The first control that takes the focus carries no TabIndex record
    and the rest are numbered, as Access writes them; a control after the
    removed one moves up, and the one that lands on 0 drops the record."""
    form = blank.add_form("Tabs")
    for i in range(3):
        form.add_control("TextBox", f"T{i}", top=240 + i * 400)
    form.add_control("Label", "Fixed", top=2000)
    assert [c.tab_index for c in form.controls] == [None, 1, 2, None]

    form.remove_control("T1")

    assert [c.name for c in form.controls] == ["T0", "T2", "Fixed"]
    assert [c.tab_index for c in form.controls] == [None, 1, None]
    form.remove_control("T0")
    assert [c.tab_index for c in form.controls] == [None, None]  # T2 is first now and carries no record


def test_a_page_comes_off_its_tab_control(blank: AccessDatabase) -> None:
    form = blank.add_form("Tabbed")
    form.add_control("Tab", "Holder", width=4000, height=3000)
    form.add_control("Page", "One", parent="Holder")
    form.add_control("Page", "Two", container="Holder")
    form.add_control("Page", "Three", parent="Holder")
    assert [c.name for c in form.controls] == ["Holder", "One", "Two", "Three"]

    form.remove_control("Two")

    assert [c.name for c in form.controls] == ["Holder", "One", "Three"]
    pages = [c for c in form.controls if c.kind == "Page"]
    assert [c.marker for c in pages] == [OPEN_CONTROL, OPEN_SIBLING] and pages[0].code == 2
    form.remove_control("Holder")
    assert form.controls == ()


def test_a_removed_control_stays_removed_across_a_save(blank: AccessDatabase, tmp_path: Path) -> None:
    form = blank.add_form("Kept")
    form.add_control("Label", "Gone", caption="x")
    form.add_control("Label", "Stays", caption="y")
    form.remove_control("Gone")
    blank.save(tmp_path / "kept.accdb")

    reopened = AccessDatabase(tmp_path / "kept.accdb")
    assert [c.name for c in reopened.form("Kept").controls] == ["Stays"]


def test_code_goes_behind_a_form_through_the_form(blank: AccessDatabase) -> None:
    form = blank.add_form("Coded")

    module = form.set_code("Option Compare Database\r\nPrivate Sub Form_Load()\r\nEnd Sub")

    assert module.name == "Form_Coded"
    assert "Form_Load" in blank.get_module("Form_Coded")
