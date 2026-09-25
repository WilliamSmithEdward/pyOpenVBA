"""Live gate: the editor takes the library's PROJECT declarations as its own (opt-in).

What only Excel and Word can answer: that a project the library gave two
forms, a module between them and then took the first form out of, loads
with exactly those components, and that when the application adds a
module and saves, it writes every declaration back where the library put
it, the Forms Package line included, and adds its module after them. Each
file starts as tests/fixtures/binary_project/, which the application saved
itself; the rules are tests/fixtures/project_package.json.

Opt-in per application: RUN_LIVE_EXCEL=1 or RUN_LIVE_WORD=1, on Windows
with it installed. Skipped everywhere else, including CI.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Any

import pytest

from pyopenvba import ExcelFile, WordFile
from pyopenvba.cfb import CFB

FIXTURES = Path(__file__).parent / "fixtures" / "binary_project"
#: Each application: the switch that opts in, the library's file type, the fixture, the storage the project lives
#: in, the harness session, and how its VBA opens, saves and closes a file.
HOSTS: dict[str, tuple[str, Any, str, str, str, str, str]] = {
    "excel": ("RUN_LIVE_EXCEL", ExcelFile, "workbook.xls", "_VBA_PROJECT_CUR", "ExcelSession",
              'Application.DisplayAlerts = False\nSet d = Workbooks.Open("{path}")',
              'Application.DisplayAlerts = False\nSet d = Workbooks("{name}")\nd.Save\nd.Close False'),
    "word": ("RUN_LIVE_WORD", WordFile, "document.doc", "Macros", "WordSession",
             'Application.DisplayAlerts = 0\nSet d = Documents.Open(FileName:="{path}", AddToRecentFiles:=False)',
             "d.Save\nd.Close SaveChanges:=0"),
}
#: What the application lists once the library is done: each component and its type (1 module, 3 form,
#: 100 document).
COMPONENTS = {"excel": "ThisWorkbook:100;Sheet1:100;Module1:1;Module2:1;UserForm2:3;",
              "word": "ThisDocument:100;Module1:1;Module2:1;UserForm2:3;"}


def _param(host: str) -> Any:
    switch = HOSTS[host][0]
    return pytest.param(host, id=host, marks=pytest.mark.skipif(
        os.environ.get(switch) != "1" or sys.platform != "win32",
        reason=f"live project package gate: set {switch}=1 on Windows with the application installed"))


def _run(office: Any, body: str) -> str:
    done = office.run_vba(f"Public Function Step() As String\nDim d As Object, c As Object, out As String\n{body}\n"
                          "Step = out\nEnd Function\n", "Step", timeout=120.0)
    assert done.ok, f"{done.outcome}: {done.message}"
    return str(done.value)


def _declarations(path: Path, storage: str) -> tuple[list[str], list[str]]:
    """The declarations between ID= and Name=, and the storages beside VBA/."""
    cfb = CFB.from_bytes(path.read_bytes())
    project = cfb.get_stream_at((storage,), "PROJECT").decode("latin-1").split("\r\n")
    end = next(i for i, line in enumerate(project) if line.startswith("Name="))
    return project[1:end], sorted(n for n in cfb.list_storages_at((storage,)) if n.casefold() != "vba")


@pytest.mark.parametrize("host", [_param(host) for host in HOSTS])
def test_the_editor_keeps_the_librarys_declarations(host: str, tmp_path: Path) -> None:
    harness = pytest.importorskip("pyvbaharness")
    _, opener, name, storage, session_name, opening, closing = HOSTS[host]
    target = tmp_path / name
    shutil.copyfile(FIXTURES / name, target)
    for step in ("form UserForm1", "module Module2", "form UserForm2", "remove UserForm1"):
        verb, _, what = step.partition(" ")
        with opener(target) as office_file:
            if verb == "form":
                office_file.add_form(what)
            elif verb == "module":
                office_file.vba_project().add_module(what, f'Attribute VB_Name = "{what}"\r\n')
            else:
                office_file.vba_project().delete_module(what)
            office_file.save()
    ours, storages = _declarations(target, storage)
    assert storages == ["UserForm2"]

    listing = ("For Each c In d.VBProject.VBComponents\nout = out & c.Name & \":\" & c.Type & \";\"\nNext\n"
               'out = out & "|" & d.VBProject.VBComponents.Add(1).Name')
    with getattr(harness, session_name)(harness.HarnessConfig(lock_wait_s=1200.0)) as office:
        office.new_document()
        if host == "excel":
            # Excel stops at a VBA dialog when one call changes a project and saves it: a call each.
            seen = _run(office, f"{opening.format(path=target)}\n{listing}")
            _run(office, closing.format(name=name))
        else:
            # Word loses the harness's procedures once another file's project changed: one call.
            seen = _run(office, f"{opening.format(path=target)}\n{listing}\n{closing}")
    components, added = seen.split("|")
    assert components == COMPONENTS[host]
    theirs, storages = _declarations(target, storage)
    assert theirs == [*ours, f"Module={added}"]
    assert storages == ["UserForm2"]
